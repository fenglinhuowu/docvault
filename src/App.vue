<template>
  <div class="docvault-app">
    <!-- 顶部工具栏 -->
    <header class="toolbar">
      <div class="toolbar-left">
        <h1 class="app-title">📚 DocVault</h1>
        <span class="app-subtitle">完全离线的文档管理</span>
      </div>
      <div class="toolbar-actions">
        <button :disabled="state.loading" class="btn btn-primary" @click="openFile">
          📂 打开文件
        </button>
        <button :disabled="!state.currentFile || state.loading" class="btn btn-success" @click="saveFile">
          💾 保存
        </button>
        <button :disabled="!state.currentFile || state.loading" class="btn btn-warning" @click="convertToPdf">
          🔄 Word转PDF
        </button>
      </div>
    </header>

    <!-- 状态提示 -->
    <div v-if="state.error" class="alert alert-error">
      {{ state.error }}
      <button class="alert-close" @click="state.error = ''">×</button>
    </div>
    <div v-if="state.success" class="alert alert-success">
      {{ state.success }}
      <button class="alert-close" @click="state.success = ''">×</button>
    </div>

    <!-- 转换进度条 -->
    <div v-if="state.conversionProgress > 0" class="progress-bar">
      <div class="progress-fill" :style="{ width: state.conversionProgress + '%' }"></div>
    </div>

    <!-- 主内容区 -->
    <main class="main-content">
      <!-- 左侧文件列表 -->
      <aside class="sidebar">
        <div class="sidebar-header">
          <h3>文件列表</h3>
        </div>
        <div class="file-list">
          <div
            v-for="file in state.fileList"
            :key="file.path"
            class="file-item"
            :class="{ active: state.currentFile?.path === file.path }"
            @click="openFileByPath(file.path)"
          >
            <span class="file-icon">{{ getFileIcon(file.extension) }}</span>
            <div class="file-info">
              <span class="file-name">{{ file.name }}</span>
              <span class="file-meta">{{ formatSize(file.size) }}</span>
            </div>
          </div>
          <div v-if="state.fileList.length === 0" class="empty-state">
            点击"打开文件"开始
          </div>
        </div>
      </aside>

      <!-- 右侧编辑/预览区 -->
      <section class="editor-area">
        <div v-if="state.currentFile" class="document-viewer">
          <div class="doc-header">
            <h2>{{ state.currentFile.name }}</h2>
            <div class="doc-meta">
              <span>大小: {{ formatSize(state.currentFile.size) }}</span>
              <span>类型: {{ state.currentFile.mime_type }}</span>
            </div>
          </div>

          <!-- 文本文件 -->
          <div v-if="parsedContent.type === 'text'" class="text-viewer">
            <textarea
              v-model="textContent"
              class="text-editor"
              placeholder="输入文本内容..."
            ></textarea>
          </div>

          <!-- Word 文档 -->
          <div v-else-if="parsedContent.type === 'html'" class="word-editor">
            <div class="editor-toolbar">
              <button class="btn btn-small" @click="formatBlock('bold')" title="加粗"><b>B</b></button>
              <button class="btn btn-small" @click="formatBlock('italic')" title="斜体"><i>I</i></button>
              <button class="btn btn-small" @click="formatBlock('underline')" title="下划线"><u>U</u></button>
              <button class="btn btn-small" @click="formatBlock('strikeThrough')" title="删除线"><s>S</s></button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="formatBlock('justifyLeft')" title="左对齐">⬅</button>
              <button class="btn btn-small" @click="formatBlock('justifyCenter')" title="居中">↔</button>
              <button class="btn btn-small" @click="formatBlock('justifyRight')" title="右对齐">➡</button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="formatBlock('insertUnorderedList')" title="无序列表">• 列表</button>
              <button class="btn btn-small" @click="formatBlock('insertOrderedList')" title="有序列表">1. 列表</button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="formatBlock('undo')" title="撤销">↩</button>
              <button class="btn btn-small" @click="formatBlock('redo')" title="重做">↪</button>
            </div>
            <div class="editor-scroll-area">
              <div
                ref="editorRef"
                class="editor-content"
                contenteditable="true"
                @input="onEditorInput"
                v-html="parsedContent.content"
              ></div>
            </div>
          </div>

          <!-- Excel 表格 -->
          <div v-else-if="parsedContent.type === 'spreadsheet'" class="excel-viewer">
            <div class="table-scroll-area">
              <table class="data-table">
                <tbody>
                  <tr v-for="(row, rowIndex) in parsedContent.rows" :key="rowIndex">
                    <td v-for="(cell, cellIndex) in row" :key="cellIndex" class="cell">
                      <input v-model="parsedContent.rows![rowIndex][cellIndex]" class="cell-input" />
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <!-- PDF 预览 -->
          <div v-else-if="parsedContent.type === 'pdf'" class="pdf-viewer">
            <div class="pdf-toolbar">
              <button class="btn btn-small" @click="pdfZoomOut">➖</button>
              <span class="pdf-zoom">{{ Math.round(pdfZoom * 100) }}%</span>
              <button class="btn btn-small" @click="pdfZoomIn">➕</button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="pdfPrevPage" :disabled="pdfPage <= 1">◀</button>
              <span class="pdf-page">{{ pdfPage }} / {{ pdfTotalPages }}</span>
              <button class="btn btn-small" @click="pdfNextPage" :disabled="pdfPage >= pdfTotalPages">▶</button>
            </div>
            <div class="pdf-scroll-area">
              <canvas ref="pdfCanvas" class="pdf-canvas"></canvas>
            </div>
          </div>

          <!-- 其他格式 -->
          <div v-else class="unsupported">
            <p>暂不支持该格式的预览</p>
          </div>
        </div>

        <!-- 空状态 -->
        <div v-else class="empty-editor">
          <div class="empty-content">
            <span class="empty-icon">📄</span>
            <h3>欢迎使用 DocVault</h3>
            <p>完全离线、跨平台的文档管理</p>
            <button class="btn btn-primary btn-large" @click="openFile">
              打开文件
            </button>
          </div>
        </div>
      </section>
    </main>

    <!-- 底部状态栏 -->
    <footer class="status-bar">
      <span v-if="state.loading" class="status-loading">⏳ 处理中...</span>
      <span v-else class="status-ready">✅ 就绪</span>
      <span class="status-info">{{ state.currentFile ? state.currentFile.name : '未打开文件' }}</span>
    </footer>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, watch, nextTick } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { useDocVault } from './composables/useDocVault'
import * as pdfjsLib from 'pdfjs-dist'

// PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.mjs',
  import.meta.url
).toString()

const { state, parsedContent, openFile, saveFile, convertToPdf, scanDir, formatSize, getFileIcon } = useDocVault()

const editorRef = ref<HTMLDivElement | null>(null)
const textContent = ref('')
const pdfCanvas = ref<HTMLCanvasElement | null>(null)
const pdfDoc = ref<any>(null)
const pdfPage = ref(1)
const pdfTotalPages = ref(0)
const pdfZoom = ref(1.0)

// PDF 渲染
watch(
  () => parsedContent.value,
  async (content) => {
    if (content.type === 'pdf' && content.content) {
      await nextTick()
      await loadPdf(content.content)
    }
  }
)

async function loadPdf(base64: string) {
  try {
    const binaryString = atob(base64)
    const bytes = new Uint8Array(binaryString.length)
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i)
    }

    const loadingTask = pdfjsLib.getDocument({ data: bytes })
    pdfDoc.value = await loadingTask.promise
    pdfTotalPages.value = pdfDoc.value.numPages
    pdfPage.value = 1
    await renderPdfPage()
  } catch (err) {
    console.error('PDF load error:', err)
  }
}

async function renderPdfPage() {
  if (!pdfDoc.value || !pdfCanvas.value) return

  try {
    const page = await pdfDoc.value.getPage(pdfPage.value)
    const viewport = page.getViewport({ scale: pdfZoom.value })
    const canvas = pdfCanvas.value
    const context = canvas.getContext('2d')!

    canvas.height = viewport.height
    canvas.width = viewport.width

    await page.render({
      canvasContext: context,
      viewport: viewport,
    }).promise
  } catch (err) {
    console.error('PDF render error:', err)
  }
}

function pdfZoomIn() {
  pdfZoom.value = Math.min(pdfZoom.value + 0.25, 3.0)
  renderPdfPage()
}

function pdfZoomOut() {
  pdfZoom.value = Math.max(pdfZoom.value - 0.25, 0.5)
  renderPdfPage()
}

function pdfPrevPage() {
  if (pdfPage.value > 1) {
    pdfPage.value--
    renderPdfPage()
  }
}

function pdfNextPage() {
  if (pdfPage.value < pdfTotalPages.value) {
    pdfPage.value++
    renderPdfPage()
  }
}

// 方法
async function openFileByPath(path: string) {
  try {
    state.loading = true
    const fileData = await invoke<FileData>('open_local_file', { path })
    state.currentFile = fileData
  } catch (err) {
    state.error = `打开失败: ${err}`
  } finally {
    state.loading = false
  }
}

function formatBlock(command: string, value?: string) {
  document.execCommand(command, false, value)
}

function onEditorInput() {
  if (editorRef.value) {
    parsedContent.value.content = editorRef.value.innerHTML
  }
}

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

onMounted(async () => {})
</script>

<style scoped>
.docvault-app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: #f5f5f5;
}

.toolbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px 24px;
  background: #1a1a2e;
  color: white;
  flex-shrink: 0;
}

.toolbar-left {
  display: flex;
  align-items: baseline;
  gap: 12px;
}

.app-title {
  font-size: 20px;
  font-weight: 700;
  margin: 0;
}

.app-subtitle {
  font-size: 12px;
  opacity: 0.7;
}

.toolbar-actions {
  display: flex;
  gap: 10px;
}

.btn {
  padding: 8px 16px;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  font-weight: 500;
  transition: all 0.2s;
}

.btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.btn-primary { background: #4361ee; color: white; }
.btn-primary:hover:not(:disabled) { background: #3a56d4; }
.btn-success { background: #06d6a0; color: white; }
.btn-success:hover:not(:disabled) { background: #05c291; }
.btn-warning { background: #ff9e00; color: white; }
.btn-warning:hover:not(:disabled) { background: #e68f00; }
.btn-small {
  padding: 4px 10px;
  font-size: 12px;
  background: rgba(255, 255, 255, 0.2);
  color: white;
}

.alert {
  padding: 10px 24px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 14px;
  flex-shrink: 0;
}

.alert-error { background: #fee2e2; color: #dc2626; }
.alert-success { background: #dcfce7; color: #16a34a; }
.alert-close { background: none; border: none; font-size: 18px; cursor: pointer; opacity: 0.7; }

.progress-bar { height: 3px; background: #e5e7eb; flex-shrink: 0; }
.progress-fill { height: 100%; background: #4361ee; transition: width 0.3s ease; }

.main-content {
  display: flex;
  flex: 1;
  overflow: hidden;
}

.sidebar {
  width: 260px;
  background: white;
  border-right: 1px solid #e5e7eb;
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
}

.sidebar-header {
  padding: 12px 16px;
  border-bottom: 1px solid #e5e7eb;
}

.sidebar-header h3 { margin: 0; font-size: 14px; color: #374151; }

.file-list {
  flex: 1;
  overflow-y: auto;
  padding: 8px;
}

.file-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 12px;
  border-radius: 6px;
  cursor: pointer;
}

.file-item:hover { background: #f3f4f6; }
.file-item.active { background: #eef2ff; border: 1px solid #c7d2fe; }

.file-icon { font-size: 18px; }

.file-info { display: flex; flex-direction: column; overflow: hidden; }
.file-name { font-size: 13px; color: #1f2937; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.file-meta { font-size: 11px; color: #9ca3af; }

.empty-state { text-align: center; padding: 30px 20px; color: #9ca3af; font-size: 13px; }

.editor-area {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.document-viewer {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.doc-header {
  padding: 12px 24px;
  background: white;
  border-bottom: 1px solid #e5e7eb;
  flex-shrink: 0;
}

.doc-header h2 { margin: 0 0 4px 0; font-size: 16px; color: #1f2937; }
.doc-meta { display: flex; gap: 16px; font-size: 12px; color: #6b7280; }

/* 文本编辑器 */
.text-viewer {
  flex: 1;
  overflow: hidden;
  background: white;
}

.text-editor {
  width: 100%;
  height: 100%;
  border: none;
  outline: none;
  padding: 24px;
  font-family: 'Menlo', 'Monaco', monospace;
  font-size: 14px;
  line-height: 1.6;
  resize: none;
  background: #fafafa;
}

/* Word 编辑器 */
.word-editor {
  flex: 1;
  display: flex;
  flex-direction: column;
  background: white;
  overflow: hidden;
}

.editor-toolbar {
  display: flex;
  gap: 4px;
  padding: 8px 16px;
  border-bottom: 1px solid #e5e7eb;
  background: #f9fafb;
  flex-shrink: 0;
  flex-wrap: wrap;
}

.separator { color: #d1d5db; padding: 0 4px; }

.editor-scroll-area {
  flex: 1;
  overflow: auto;
  background: #e5e7eb;
  padding: 24px;
}

.editor-content {
  max-width: 800px;
  min-height: 100%;
  margin: 0 auto;
  padding: 40px;
  background: white;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  outline: none;
  line-height: 1.8;
  color: #374151;
}

/* Excel 查看器 */
.excel-viewer {
  flex: 1;
  overflow: auto;
  background: white;
}

.table-scroll-area {
  padding: 16px;
  overflow: auto;
}

.data-table {
  border-collapse: collapse;
  width: 100%;
  font-size: 13px;
}

.data-table td {
  border: 1px solid #e5e7eb;
  padding: 0;
  min-width: 80px;
}

.cell-input {
  width: 100%;
  height: 100%;
  border: none;
  outline: none;
  padding: 6px 10px;
  font-size: 13px;
}

.data-table tr:first-child td .cell-input {
  background: #f3f4f6;
  font-weight: 600;
}

/* PDF 查看器 */
.pdf-viewer {
  flex: 1;
  display: flex;
  flex-direction: column;
  background: #374151;
  overflow: hidden;
}

.pdf-toolbar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 16px;
  background: #1f2937;
  color: white;
  flex-shrink: 0;
}

.pdf-zoom, .pdf-page {
  font-size: 12px;
  min-width: 60px;
  text-align: center;
}

.pdf-scroll-area {
  flex: 1;
  overflow: auto;
  display: flex;
  justify-content: center;
  padding: 24px;
}

.pdf-canvas {
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
}

.unsupported {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  color: #6b7280;
}

.empty-editor {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
}

.empty-content { text-align: center; max-width: 400px; }
.empty-icon { font-size: 64px; }
.empty-content h3 { margin: 16px 0 8px; color: #1f2937; }
.empty-content p { color: #6b7280; margin-bottom: 24px; }

.btn-large { padding: 12px 24px; font-size: 16px; }

.status-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 6px 24px;
  background: #1a1a2e;
  color: white;
  font-size: 12px;
  flex-shrink: 0;
}

.status-loading { color: #fbbf24; }
.status-ready { color: #34d399; }
.status-info { opacity: 0.8; }
</style>
