import { ref, reactive } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { open, save } from '@tauri-apps/plugin-dialog'
import { readTextFile, writeTextFile, exists } from '@tauri-apps/plugin-fs'

// ==================== 类型定义 ====================

interface FileData {
  path: string
  name: string
  size: number
  mime_type: string
  extension: string
  content_base64: string
  modified: string | null
  created: string | null
}

interface FileInfo {
  path: string
  name: string
  size: number
  is_dir: boolean
  extension: string
  modified: string | null
  mime_type: string
}

interface ConvertResult {
  success: boolean
  output_path: string
  message: string
  file_size: number
}

// ==================== 状态管理 ====================

const state = reactive({
  currentFile: null as FileData | null,
  fileList: [] as FileInfo[],
  loading: false,
  error: '',
  success: '',
  conversionProgress: 0,
})

// ==================== 核心功能 ====================

/**
 * 打开本地文件（完全离线）
 */
async function openFile() {
  try {
    state.loading = true
    state.error = ''

    const selected = await open({
      multiple: false,
      filters: [
        { name: '文档', extensions: ['docx', 'xlsx', 'pdf', 'doc', 'xls'] },
        { name: 'Word', extensions: ['docx', 'doc'] },
        { name: 'Excel', extensions: ['xlsx', 'xls'] },
        { name: 'PDF', extensions: ['pdf'] },
        { name: '所有文件', extensions: ['*'] },
      ],
    })

    if (!selected) {
      state.loading = false
      return
    }

    const fileData = await invoke<FileData>('open_local_file', {
      path: selected as string,
    })

    state.currentFile = fileData
    state.success = `已打开: ${fileData.name} (${formatSize(fileData.size)})`
  } catch (err) {
    state.error = `打开文件失败: ${err}`
  } finally {
    state.loading = false
  }
}

/**
 * 保存文件到本地
 */
async function saveFile() {
  if (!state.currentFile) {
    state.error = '没有可保存的文件'
    return
  }

  try {
    state.loading = true
    state.error = ''

    const savePath = await save({
      defaultPath: state.currentFile.name,
      filters: [
        { name: 'Word 文档', extensions: ['docx'] },
        { name: 'Excel 表格', extensions: ['xlsx'] },
        { name: 'PDF 文档', extensions: ['pdf'] },
        { name: '纯文本', extensions: ['txt'] },
      ],
    })

    if (!savePath) {
      state.loading = false
      return
    }

    // 将 Base64 转回 Uint8Array
    const binaryString = atob(state.currentFile.content_base64)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }

    await invoke('save_local_file', {
      path: savePath,
      content: Array.from(bytes),
    })

    state.success = `已保存到: ${savePath}`
  } catch (err) {
    state.error = `保存失败: ${err}`
  } finally {
    state.loading = false
  }
}

/**
 * Word 转 PDF（完全离线）
 */
async function convertToPdf() {
  if (!state.currentFile) {
    state.error = '请先打开一个 Word 文档'
    return
  }

  if (state.currentFile.extension !== 'docx') {
    state.error = '当前文件不是 .docx 格式'
    return
  }

  try {
    state.loading = true
    state.error = ''
    state.conversionProgress = 0

    // 选择输出路径
    const outputPath = await save({
      defaultPath: state.currentFile.name.replace('.docx', '.pdf'),
      filters: [{ name: 'PDF 文档', extensions: ['pdf'] }],
    })

    if (!outputPath) {
      state.loading = false
      return
    }

    state.conversionProgress = 30

    const result = await invoke<ConvertResult>('convert_docx_to_pdf', {
      docxPath: state.currentFile.path,
      pdfPath: outputPath,
    })

    state.conversionProgress = 100

    if (result.success) {
      state.success = `转换成功: ${result.output_path} (${formatSize(result.file_size)})`
    } else {
      state.error = `转换失败: ${result.message}`
    }
  } catch (err) {
    state.error = `转换错误: ${err}`
  } finally {
    state.loading = false
    state.conversionProgress = 0
  }
}

/**
 * 扫描目录
 */
async function scanDir(dirPath: string) {
  try {
    state.loading = true
    state.error = ''

    const files = await invoke<FileInfo[]>('scan_directory', {
      dirPath,
    })

    state.fileList = files.filter(f =>
      ['docx', 'xlsx', 'pdf', 'doc', 'xls'].includes(f.extension)
    )
  } catch (err) {
    state.error = `扫描目录失败: ${err}`
  } finally {
    state.loading = false
  }
}

// ==================== 工具函数 ====================

function formatSize(bytes: number): string {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

function getFileIcon(extension: string): string {
  const icons: Record<string, string> = {
    docx: '📄',
    doc: '📄',
    xlsx: '📊',
    xls: '📊',
    pdf: '📕',
    txt: '📝',
  }
  return icons[extension] || '📁'
}

// ==================== 导出 ====================

export function useDocVault() {
  return {
    state,
    openFile,
    saveFile,
    convertToPdf,
    scanDir,
    formatSize,
    getFileIcon,
  }
}
