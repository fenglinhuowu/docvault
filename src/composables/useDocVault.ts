import { ref, reactive } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { open, save } from '@tauri-apps/plugin-dialog'

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

interface ParsedContent {
  type: 'text' | 'html' | 'pdf' | 'spreadsheet' | 'unsupported'
  content: string
  rows?: string[][]
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

const parsedContent = ref<ParsedContent>({ type: 'unsupported', content: '' })

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
        { name: '文档', extensions: ['docx', 'xlsx', 'pdf', 'doc', 'xls', 'txt'] },
        { name: 'Word', extensions: ['docx', 'doc'] },
        { name: 'Excel', extensions: ['xlsx', 'xls'] },
        { name: 'PDF', extensions: ['pdf'] },
        { name: '文本', extensions: ['txt'] },
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

    // 解析文件内容
    await parseFileContent(fileData)
  } catch (err) {
    state.error = `打开文件失败: ${err}`
  } finally {
    state.loading = false
  }
}

/**
 * 解析文件内容
 */
async function parseFileContent(fileData: FileData) {
  const ext = fileData.extension

  if (ext === 'txt') {
    // 文本文件直接解码
    const text = atob(fileData.content_base64)
    parsedContent.value = { type: 'text', content: text }
  } else if (ext === 'pdf') {
    // PDF 使用 blob URL
    parsedContent.value = { type: 'pdf', content: fileData.content_base64 }
  } else if (ext === 'docx') {
    // 解析 docx 提取文本
    const text = await parseDocx(fileData.content_base64)
    parsedContent.value = { type: 'html', content: text }
  } else if (ext === 'xlsx' || ext === 'xls') {
    // 解析 xlsx 提取数据
    const rows = await parseXlsx(fileData.content_base64)
    parsedContent.value = { type: 'spreadsheet', content: '', rows }
  } else {
    parsedContent.value = { type: 'unsupported', content: '暂不支持该格式' }
  }
}

/**
 * 解析 docx 文件提取文本
 */
async function parseDocx(base64: string): Promise<string> {
  try {
    const binaryString = atob(base64)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }

    // 解压 zip (docx 本质是 zip)
    const zip = await decompressZip(bytes)
    const documentXml = zip['word/document.xml']

    if (!documentXml) return '<p>无法解析文档内容</p>'

    // 解析 XML 提取文本
    const text = extractTextFromDocxXml(documentXml)
    return text
  } catch {
    return '<p>文档解析失败</p>'
  }
}

/**
 * 简化的 zip 解压
 */
async function decompressZip(data: Uint8Array): Promise<Record<string, string>> {
  const result: Record<string, string> = {}

  // 使用简单的 ZIP 解析（仅支持未压缩的文件）
  const view = new DataView(data.buffer)
  let offset = 0

  while (offset < data.length) {
    // 查找本地文件头签名
    if (view.getUint32(offset, true) !== 0x04034b50) {
      offset++
      continue
    }

    const compressionMethod = view.getUint16(offset + 8, true)
    const compressedSize = view.getUint32(offset + 18, true)
    const fileNameLength = view.getUint16(offset + 26, true)
    const extraFieldLength = view.getUint16(offset + 28, true)

    const fileName = new TextDecoder().decode(
      data.slice(offset + 30, offset + 30 + fileNameLength)
    )

    const dataStart = offset + 30 + fileNameLength + extraFieldLength

    if (compressionMethod === 0) {
      // 未压缩
      const fileData = data.slice(dataStart, dataStart + compressedSize)
      result[fileName] = new TextDecoder().decode(fileData)
    } else if (compressionMethod === 8) {
      // DEFLATE 压缩 - 使用 pako 或类似库
      const compressedData = data.slice(dataStart, dataStart + compressedSize)
      try {
        const decompressed = await decompressDeflate(compressedData)
        result[fileName] = new TextDecoder().decode(decompressed)
      } catch {
        // 解压失败则跳过
      }
    }

    offset = dataStart + compressedSize
  }

  return result
}

/**
 * DEFLATE 解压（使用 CompressionStream）
 */
async function decompressDeflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Response(data).body!
    .pipeThrough(new DecompressionStream('deflate'))
  const response = await new Response(stream).arrayBuffer()
  return new Uint8Array(response)
}

/**
 * 从 docx XML 提取文本
 */
function extractTextFromDocxXml(xml: string): string {
  const paragraphs: string[] = []
  const parser = new DOMParser()
  const doc = parser.parseFromString(xml, 'text/xml')

  const pElements = doc.getElementsByTagName('w:p')
  for (let i = 0; i < pElements.length; i++) {
    const p = pElements[i]
    const texts: string[] = []

    const tElements = p.getElementsByTagName('w:t')
    for (let j = 0; j < tElements.length; j++) {
      texts.push(tElements[j].textContent || '')
    }

    const lineBreaks = p.getElementsByTagName('w:br')
    if (lineBreaks.length > 0) {
      // 处理换行
    }

    const paragraphText = texts.join('')
    if (paragraphText.trim()) {
      paragraphs.push(`<p>${escapeHtml(paragraphText)}</p>`)
    }
  }

  return paragraphs.join('') || '<p>文档为空</p>'
}

/**
 * HTML 转义
 */
function escapeHtml(text: string): string {
  const div = document.createElement('div')
  div.textContent = text
  return div.innerHTML
}

/**
 * 解析 xlsx 文件提取数据
 */
async function parseXlsx(base64: string): Promise<string[][]> {
  try {
    const binaryString = atob(base64)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }

    const zip = await decompressZip(bytes)
    const sheetXml = zip['xl/worksheets/sheet1.xml']
    const sharedStringsXml = zip['xl/sharedStrings.xml']

    if (!sheetXml) return []

    // 解析共享字符串
    const sharedStrings: string[] = []
    if (sharedStringsXml) {
      const parser = new DOMParser()
      const doc = parser.parseFromString(sharedStringsXml, 'text/xml')
      const siElements = doc.getElementsByTagName('si')
      for (let i = 0; i < siElements.length; i++) {
        const tElements = siElements[i].getElementsByTagName('t')
        let text = ''
        for (let j = 0; j < tElements.length; j++) {
          text += tElements[j].textContent || ''
        }
        sharedStrings.push(text)
      }
    }

    // 解析工作表数据
    const parser = new DOMParser()
    const doc = parser.parseFromString(sheetXml, 'text/xml')
    const rows: string[][] = []

    const rowElements = doc.getElementsByTagName('row')
    for (let i = 0; i < rowElements.length; i++) {
      const row: string[] = []
      const cElements = rowElements[i].getElementsByTagName('c')

      for (let j = 0; j < cElements.length; j++) {
        const cell = cElements[j]
        const type = cell.getAttribute('t')
        const vElement = cell.getElementsByTagName('v')[0]
        let value = vElement?.textContent || ''

        if (type === 's' && value) {
          // 共享字符串引用
          const index = parseInt(value)
          value = sharedStrings[index] || value
        }

        row.push(value)
      }

      rows.push(row)
    }

    return rows
  } catch {
    return []
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
      ['docx', 'xlsx', 'pdf', 'doc', 'xls', 'txt'].includes(f.extension)
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
    parsedContent,
    openFile,
    saveFile,
    convertToPdf,
    scanDir,
    formatSize,
    getFileIcon,
  }
}
