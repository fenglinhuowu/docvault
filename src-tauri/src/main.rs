use std::fs;
use std::path::{Path, PathBuf};

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

// ==================== 数据模型 ====================

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct FileData {
    pub path: String,
    pub name: String,
    pub size: u64,
    pub mime_type: String,
    pub extension: String,
    pub content_base64: String,
    pub modified: Option<String>,
    pub created: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct FileInfo {
    pub path: String,
    pub name: String,
    pub size: u64,
    pub is_dir: bool,
    pub extension: String,
    pub modified: Option<String>,
    pub mime_type: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ConvertResult {
    pub success: bool,
    pub output_path: String,
    pub message: String,
    pub file_size: u64,
}

// ==================== 命令模块 ====================
// 将命令函数放在单独模块中避免宏名称冲突

pub mod commands {
    use super::*;
    use base64::Engine;

    /// 完全离线读取本地文件，返回 Base64 编码的内容
    #[tauri::command]
    pub async fn open_local_file(path: String) -> Result<FileData, String> {
        let path = validate_path(&path)?;

        if !path.exists() {
            return Err(format!("File not found: {}", path.display()));
        }

        if !path.is_file() {
            return Err(format!("Path is not a file: {}", path.display()));
        }

        let metadata = fs::metadata(&path).map_err(|e| e.to_string())?;

        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

        let mime_type = mime_guess::from_path(&path)
            .first_or_octet_stream()
            .to_string();

        let content = fs::read(&path).map_err(|e| e.to_string())?;
        let content_base64 = base64::engine::general_purpose::STANDARD.encode(&content);

        let modified = metadata
            .modified()
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| {
                let dt: DateTime<Utc> = DateTime::from_timestamp(d.as_secs() as i64, 0)
                    .unwrap_or_else(|| Utc::now());
                dt.to_rfc3339()
            });

        let created = metadata
            .created()
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| {
                let dt: DateTime<Utc> = DateTime::from_timestamp(d.as_secs() as i64, 0)
                    .unwrap_or_else(|| Utc::now());
                dt.to_rfc3339()
            });

        Ok(FileData {
            path: path.to_string_lossy().to_string(),
            name: path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string(),
            size: metadata.len(),
            mime_type,
            extension,
            content_base64,
            modified,
            created,
        })
    }

    /// 将内容写入本地文件
    #[tauri::command]
    pub async fn save_local_file(path: String, content: Vec<u8>) -> Result<(), String> {
        let path = validate_path(&path)?;

        if let Some(parent) = path.parent() {
            if !parent.exists() {
                fs::create_dir_all(parent).map_err(|e| e.to_string())?;
            }
        }

        fs::write(&path, &content).map_err(|e| {
            if e.kind() == std::io::ErrorKind::PermissionDenied {
                format!("Permission denied: {}", path.display())
            } else {
                e.to_string()
            }
        })?;

        log::info!("Saved file: {} ({} bytes)", path.display(), content.len());
        Ok(())
    }

    /// Word 文档转换为 PDF（离线实现）
    #[tauri::command]
    pub async fn convert_docx_to_pdf(
        docx_path: String,
        pdf_path: String,
    ) -> Result<ConvertResult, String> {
        let docx_path = validate_path(&docx_path)?;
        let pdf_path = validate_path(&pdf_path)?;

        if !docx_path.exists() {
            return Err(format!("File not found: {}", docx_path.display()));
        }

        let ext = docx_path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

        if ext != "docx" {
            return Err(format!("Expected .docx file, got .{}", ext));
        }

        if let Some(parent) = pdf_path.parent() {
            if !parent.exists() {
                fs::create_dir_all(parent).map_err(|e| e.to_string())?;
            }
        }

        match perform_docx_to_pdf_conversion(&docx_path, &pdf_path).await {
            Ok(file_size) => {
                let result = ConvertResult {
                    success: true,
                    output_path: pdf_path.to_string_lossy().to_string(),
                    message: "Conversion completed successfully".to_string(),
                    file_size,
                };
                log::info!(
                    "Converted {} -> {} ({} bytes)",
                    docx_path.display(),
                    pdf_path.display(),
                    file_size
                );
                Ok(result)
            }
            Err(e) => Err(format!("Conversion error: {}", e)),
        }
    }

    /// 扫描目录
    #[tauri::command]
    pub async fn scan_directory(dir_path: String) -> Result<Vec<FileInfo>, String> {
        let path = validate_path(&dir_path)?;

        if !path.exists() || !path.is_dir() {
            return Err(format!("Not a valid directory: {}", path.display()));
        }

        let mut files = Vec::new();

        for entry in walkdir::WalkDir::new(&path)
            .max_depth(3)
            .into_iter()
            .filter_map(|e| e.ok())
        {
            let entry_path = entry.path().to_path_buf();
            let metadata = match entry.metadata() {
                Ok(m) => m,
                Err(_) => continue,
            };

            let extension = entry_path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_lowercase();

            let mime_type = mime_guess::from_path(&entry_path)
                .first_or_octet_stream()
                .to_string();

            let modified = metadata
                .modified()
                .ok()
                .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                .map(|d| {
                    let dt: DateTime<Utc> = DateTime::from_timestamp(d.as_secs() as i64, 0)
                        .unwrap_or_else(|| Utc::now());
                    dt.to_rfc3339()
                });

            files.push(FileInfo {
                path: entry_path.to_string_lossy().to_string(),
                name: entry.file_name().to_string_lossy().to_string(),
                size: metadata.len(),
                is_dir: entry.file_type().is_dir(),
                extension,
                modified,
                mime_type,
            });
        }

        Ok(files)
    }

    /// 获取文件信息
    #[tauri::command]
    pub async fn get_file_info(path: String) -> Result<FileInfo, String> {
        let path = validate_path(&path)?;

        if !path.exists() {
            return Err(format!("File not found: {}", path.display()));
        }

        let metadata = fs::metadata(&path).map_err(|e| e.to_string())?;

        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

        let mime_type = mime_guess::from_path(&path)
            .first_or_octet_stream()
            .to_string();

        let modified = metadata
            .modified()
            .ok()
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| {
                let dt: DateTime<Utc> = DateTime::from_timestamp(d.as_secs() as i64, 0)
                    .unwrap_or_else(|| Utc::now());
                dt.to_rfc3339()
            });

        Ok(FileInfo {
            path: path.to_string_lossy().to_string(),
            name: path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string(),
            size: metadata.len(),
            is_dir: path.is_dir(),
            extension,
            modified,
            mime_type,
        })
    }

    // ==================== 辅助函数 ====================

    fn validate_path(path: &str) -> Result<PathBuf, String> {
        if path.is_empty() {
            return Err("Empty path".to_string());
        }
        Ok(PathBuf::from(path))
    }

    async fn perform_docx_to_pdf_conversion(
        docx_path: &Path,
        pdf_path: &Path,
    ) -> Result<u64, Box<dyn std::error::Error + Send + Sync>> {
        let docx_data = tokio::fs::read(docx_path).await?;
        let cursor = std::io::Cursor::new(&docx_data);
        let mut archive = zip::ZipArchive::new(cursor)?;

        let mut document_xml = String::new();
        {
            let mut file = archive.by_name("word/document.xml")?;
            std::io::Read::read_to_string(&mut file, &mut document_xml)?;
        }

        let parsed_content = parse_docx_xml(&document_xml);
        generate_pdf(&parsed_content, pdf_path)?;

        let output_size = tokio::fs::metadata(pdf_path).await?.len();
        Ok(output_size)
    }

    fn parse_docx_xml(xml: &str) -> Vec<String> {
        let mut paragraphs = Vec::new();
        let mut current = String::new();

        for line in xml.lines() {
            let trimmed = line.trim();
            if trimmed.contains("<w:p>") || trimmed.contains("<w:p ") {
                current.clear();
            }
            if trimmed.contains("<w:t") {
                if let Some(start) = trimmed.find(">") {
                    let text = &trimmed[start + 1..];
                    if let Some(end) = text.find("</w:t>") {
                        current.push_str(&text[..end]);
                    }
                }
            }
            if trimmed.contains("</w:p>") && !current.trim().is_empty() {
                paragraphs.push(current.clone());
            }
        }
        paragraphs
    }

    fn generate_pdf(
        paragraphs: &[String],
        output_path: &Path,
    ) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
        use printpdf::*;
        use std::io::BufWriter;

        let (doc_pdf, page1, layer1) =
            PdfDocument::new("Converted Document", Mm(210.0), Mm(297.0), "Layer 1");

        let current_layer = doc_pdf.get_page(page1).get_layer(layer1);
        let font = doc_pdf.add_builtin_font(BuiltinFont::Helvetica)?;

        let mut y_position = Mm(270.0);
        let line_height = Mm(7.0);
        let margin_left = Mm(20.0);

        for text in paragraphs {
            if y_position < Mm(20.0) {
                y_position = Mm(270.0);
            }
            current_layer.use_text(text.clone(), 12.0, margin_left, y_position, &font);
            y_position -= line_height * 2.0;
        }

        let file = std::fs::File::create(output_path)?;
        let mut buf_writer = BufWriter::new(file);
        doc_pdf.save(&mut buf_writer)?;
        Ok(())
    }
}

// ==================== Tauri 入口 ====================

fn main() {
    env_logger::init();

    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            commands::open_local_file,
            commands::save_local_file,
            commands::convert_docx_to_pdf,
            commands::scan_directory,
            commands::get_file_info,
        ])
        .run(tauri::generate_context!())
        .expect("error while running DocVault application");
}
