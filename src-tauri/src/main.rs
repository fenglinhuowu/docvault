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
    use calamine::Reader;

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

    /// 解析 Word 文档内容（返回 HTML）
    #[tauri::command]
    pub async fn parse_docx(path: String) -> Result<String, String> {
        let path = validate_path(&path)?;

        if !path.exists() {
            return Err(format!("File not found: {}", path.display()));
        }

        let content = fs::read(&path).map_err(|e| e.to_string())?;

        // 使用 docx-rs 解析
        match docx_rs::read_docx(&content) {
            Ok(docx) => {
                let mut html = String::new();
                html.push_str("<div class=\"docx-content\">");

                for child in &docx.document.children {
                    match child {
                        docx_rs::DocumentChild::Paragraph(p) => {
                            html.push_str("<p>");
                            for para_child in &p.children {
                                if let docx_rs::ParagraphChild::Run(run) = para_child {
                                    for run_child in &run.children {
                                        if let docx_rs::RunChild::Text(text) = run_child {
                                            html.push_str(&text.text);
                                        }
                                    }
                                }
                            }
                            html.push_str("</p>");
                        }
                        _ => {}
                    }
                }

                html.push_str("</div>");
                Ok(html)
            }
            Err(e) => Err(format!("Failed to parse docx: {}", e)),
        }
    }

    /// 解析 Excel 表格内容（返回 JSON）
    #[tauri::command]
    pub async fn parse_xlsx(path: String) -> Result<Vec<Vec<String>>, String> {
        let path = validate_path(&path)?;

        if !path.exists() {
            return Err(format!("File not found: {}", path.display()));
        }

        // 根据扩展名选择解析方式
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();

        let mut all_rows: Vec<Vec<String>> = Vec::new();

        if ext == "xlsx" || ext == "xlsm" {
            // 使用 calamine 读取 xlsx
            let mut workbook: calamine::Xlsx<_> = calamine::open_workbook(&path)
                .map_err(|e| format!("Failed to open xlsx: {:?}", e))?;

            let sheet_names = workbook.sheet_names();
            if !sheet_names.is_empty() {
                if let Ok(range) = workbook.worksheet_range(&sheet_names[0]) {
                    for row in range.rows() {
                        let row_data: Vec<String> = row
                            .iter()
                            .map(|cell| cell.to_string())
                            .collect();
                        all_rows.push(row_data);
                    }
                }
            }
        } else {
            // 尝试其他格式
            return Err(format!("Unsupported format: {}", ext));
        }

        Ok(all_rows)
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

        // 在阻塞线程中执行转换
        let docx_path_clone = docx_path.clone();
        let pdf_path_clone = pdf_path.clone();
        let file_size = tokio::task::spawn_blocking(move || {
            perform_docx_to_pdf_blocking(&docx_path_clone, &pdf_path_clone)
        })
        .await
        .map_err(|e| format!("Task join error: {}", e))?
        .map_err(|e| format!("Conversion error: {}", e))?;

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

    /// 同步执行 docx 转 PDF（用于 spawn_blocking）
    fn perform_docx_to_pdf_blocking(
        docx_path: &Path,
        pdf_path: &Path,
    ) -> Result<u64, Box<dyn std::error::Error + Send + Sync>> {
        use printpdf::*;
        use std::io::BufWriter;

        // 读取 docx 文件
        let docx_data = std::fs::read(docx_path)?;

        // 解析 docx 提取文本
        let paragraphs = extract_docx_text(&docx_data)
            .unwrap_or_else(|_| vec!["(无法解析文档内容)".to_string()]);

        // 生成 PDF
        let (doc_pdf, page1, layer1) =
            PdfDocument::new("Converted Document", Mm(210.0), Mm(297.0), "Layer 1");

        let current_layer = doc_pdf.get_page(page1).get_layer(layer1);

        // 加载支持中文的系统字体
        let font_data = load_chinese_font()?;
        let font = doc_pdf.add_external_font(font_data.as_slice())?;

        let mut y_position = Mm(270.0);
        let line_height = Mm(7.0);
        let margin_left = Mm(20.0);

        for text in &paragraphs {
            if y_position < Mm(20.0) {
                let (new_page, new_layer) = doc_pdf.add_page(Mm(210.0), Mm(297.0), "New Page");
                let layer = doc_pdf.get_page(new_page).get_layer(new_layer);
                y_position = Mm(270.0);

                let wrapped = wrap_text(text, 40);
                for line in wrapped {
                    layer.use_text(line, 12.0, margin_left, y_position, &font);
                    y_position -= line_height;
                }
            } else {
                let wrapped = wrap_text(text, 40);
                for line in wrapped {
                    current_layer.use_text(line, 12.0, margin_left, y_position, &font);
                    y_position -= line_height;
                }
            }
            y_position -= line_height;
        }

        let file = std::fs::File::create(pdf_path)?;
        let mut buf_writer = BufWriter::new(file);
        doc_pdf.save(&mut buf_writer)?;

        let output_size = std::fs::metadata(pdf_path)?.len();
        Ok(output_size)
    }

    /// 加载支持中文的系统字体
    fn load_chinese_font() -> Result<Vec<u8>, Box<dyn std::error::Error + Send + Sync>> {
        // 按优先级尝试加载系统字体
        let font_paths = if cfg!(target_os = "macos") {
            vec![
                "/System/Library/Fonts/STHeiti Medium.ttc",
                "/System/Library/Fonts/Hiragino Sans GB.ttc",
                "/System/Library/Fonts/PingFang.ttc",
                "/Library/Fonts/Arial Unicode.ttf",
            ]
        } else if cfg!(target_os = "windows") {
            vec![
                "C:\\Windows\\Fonts\\msyh.ttc",
                "C:\\Windows\\Fonts\\simsun.ttc",
                "C:\\Windows\\Fonts\\msyhbd.ttc",
            ]
        } else {
            vec![
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            ]
        };

        for path in &font_paths {
            if let Ok(data) = std::fs::read(path) {
                return Ok(data);
            }
        }

        Err("No Chinese font found on system".into())
    }

    /// 从 docx 文件中提取文本
    fn extract_docx_text(data: &[u8]) -> Result<Vec<String>, Box<dyn std::error::Error>> {
        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor)?;

        let mut document_xml = String::new();
        {
            let mut file = archive.by_name("word/document.xml")?;
            std::io::Read::read_to_string(&mut file, &mut document_xml)?;
        }

        Ok(parse_docx_xml(&document_xml))
    }

    fn parse_docx_xml(xml: &str) -> Vec<String> {
        let mut paragraphs = Vec::new();
        let mut current = String::new();

        for line in xml.lines() {
            let trimmed = line.trim();
            if trimmed.contains("<w:p>") || trimmed.contains("<w:p ") {
                if !current.trim().is_empty() {
                    paragraphs.push(current.clone());
                }
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
        }

        if !current.trim().is_empty() {
            paragraphs.push(current);
        }

        if paragraphs.is_empty() {
            paragraphs.push("(空文档)".to_string());
        }

        paragraphs
    }

    /// 简单的文本换行
    fn wrap_text(text: &str, max_chars: usize) -> Vec<String> {
        let mut result = Vec::new();
        let mut current_line = String::new();

        for word in text.split_whitespace() {
            if current_line.len() + word.len() + 1 > max_chars {
                if !current_line.is_empty() {
                    result.push(current_line.clone());
                    current_line.clear();
                }
                current_line.push_str(word);
            } else {
                if !current_line.is_empty() {
                    current_line.push(' ');
                }
                current_line.push_str(word);
            }
        }

        if !current_line.is_empty() {
            result.push(current_line);
        }

        if result.is_empty() {
            result.push(text.to_string());
        }

        result
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
            commands::parse_docx,
            commands::parse_xlsx,
        ])
        .run(tauri::generate_context!())
        .expect("error while running DocVault application");
}
