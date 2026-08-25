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
            <!-- 视图模式：完整显示 docx-preview 渲染的 HTML -->
            <div v-if="!isEditing" class="view-mode">
              <div class="editor-toolbar">
                <button class="btn btn-primary btn-small" @click="startEditing">✏️ 编辑</button>
                <span class="toolbar-hint">Word 文档预览（完整格式）</span>
              </div>
              <div class="editor-scroll-area">
                <div
                  class="editor-content docx-preview-content"
                  v-html="parsedContent.content"
                ></div>
              </div>
            </div>
            <!-- 编辑模式：TipTap 富文本编辑 -->
            <div v-else class="edit-mode">
              <div class="editor-toolbar">
                <button class="btn btn-success btn-small" @click="stopEditing">✅ 完成</button>
                <span class="separator">|</span>
                <!-- 字体选择 -->
                <select class="toolbar-select" @change="setFontFamily(($event.target as HTMLSelectElement).value)" :value="currentFontFamily">
                  <option value="">字体</option>
                  <option value="SimSun">宋体</option>
                  <option value="SimHei">黑体</option>
                  <option value="KaiTi">楷体</option>
                  <option value="Microsoft YaHei">微软雅黑</option>
                  <option value="Arial">Arial</option>
                  <option value="Times New Roman">Times New Roman</option>
                </select>
                <!-- 字号选择 -->
                <select class="toolbar-select" @change="setFontSize(($event.target as HTMLSelectElement).value)" :value="currentFontSize">
                  <option value="">字号</option>
                  <option value="12">小五</option>
                  <option value="14">五号</option>
                  <option value="16">小四</option>
                  <option value="18">四号</option>
                  <option value="22">三号</option>
                  <option value="24">二号</option>
                  <option value="36">一号</option>
                </select>
                <span class="separator">|</span>
                <!-- 标题格式 -->
                <select class="toolbar-select" @change="setHeading(($event.target as HTMLSelectElement).value)" :value="currentHeading">
                  <option value="">正文</option>
                  <option value="1">标题 1</option>
                  <option value="2">标题 2</option>
                  <option value="3">标题 3</option>
                </select>
                <span class="separator">|</span>
                <!-- 文字格式 -->
                <button class="btn btn-small" @click="editor?.chain().focus().toggleBold().run()" :class="{ active: editor?.isActive('bold') }" title="加粗"><b>B</b></button>
                <button class="btn btn-small" @click="editor?.chain().focus().toggleItalic().run()" :class="{ active: editor?.isActive('italic') }" title="斜体"><i>I</i></button>
                <button class="btn btn-small" @click="editor?.chain().focus().toggleUnderline().run()" :class="{ active: editor?.isActive('underline') }" title="下划线"><u>U</u></button>
                <button class="btn btn-small" @click="editor?.chain().focus().toggleStrike().run()" :class="{ active: editor?.isActive('strike') }" title="删除线"><s>S</s></button>
                <span class="separator">|</span>
                <!-- 颜色 -->
                <input type="color" class="toolbar-color" @input="setTextColor($event)" title="文字颜色" />
                <input type="color" class="toolbar-color" @input="setHighlightColor($event)" title="高亮颜色" />
                <span class="separator">|</span>
                <!-- 对齐 -->
                <button class="btn btn-small" @click="editor?.chain().focus().setTextAlign('left').run()" :class="{ active: editor?.isActive({ textAlign: 'left' }) }" title="左对齐">⬅</button>
                <button class="btn btn-small" @click="editor?.chain().focus().setTextAlign('center').run()" :class="{ active: editor?.isActive({ textAlign: 'center' }) }" title="居中">↔</button>
                <button class="btn btn-small" @click="editor?.chain().focus().setTextAlign('right').run()" :class="{ active: editor?.isActive({ textAlign: 'right' }) }" title="右对齐">➡</button>
                <span class="separator">|</span>
                <!-- 列表 -->
                <button class="btn btn-small" @click="editor?.chain().focus().toggleBulletList().run()" :class="{ active: editor?.isActive('bulletList') }" title="无序列表">•</button>
                <button class="btn btn-small" @click="editor?.chain().focus().toggleOrderedList().run()" :class="{ active: editor?.isActive('orderedList') }" title="有序列表">1.</button>
                <span class="separator">|</span>
                <!-- 撤销重做 -->
                <button class="btn btn-small" @click="editor?.chain().focus().undo().run()" title="撤销">↩</button>
                <button class="btn btn-small" @click="editor?.chain().focus().redo().run()" title="重做">↪</button>
                <span class="separator">|</span>
                <!-- 插入 -->
                <button class="btn btn-small" @click="insertImage" title="插入图片">🖼</button>
                <button class="btn btn-small" @click="insertTable" title="插入表格">⊞</button>
              </div>
              <div class="editor-scroll-area">
                <editor-content :editor="editor" class="editor-content" />
              </div>
            </div>
          </div>

          <!-- Excel 表格 -->
          <div v-else-if="parsedContent.type === 'spreadsheet'" class="excel-viewer">
            <div class="table-scroll-area">
              <table class="data-table">
                <tbody>
                  <tr v-for="(row, rowIndex) in parsedContent.rows" :key="rowIndex">
                    <td v-for="(_cell, cellIndex) in row" :key="cellIndex" class="cell">
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
              <span class="pdf-info">PDF 文档</span>
            </div>
            <div class="pdf-scroll-area">
              <iframe
                :src="pdfBlobUrl"
                class="pdf-frame"
                title="PDF Viewer"
              ></iframe>
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
import { ref, onMounted, watch, onBeforeUnmount } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { useDocVault } from './composables/useDocVault'

// TipTap 编辑器
import { useEditor, EditorContent } from '@tiptap/vue-3'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import TextAlign from '@tiptap/extension-text-align'
import { TextStyle } from '@tiptap/extension-text-style'
import { Color } from '@tiptap/extension-color'
import Highlight from '@tiptap/extension-highlight'
import Image from '@tiptap/extension-image'
import { Table } from '@tiptap/extension-table'
import TableRow from '@tiptap/extension-table-row'
import TableCell from '@tiptap/extension-table-cell'
import TableHeader from '@tiptap/extension-table-header'
import FontFamily from '@tiptap/extension-font-family'

const { state, parsedContent, openFile, saveFile, convertToPdf, formatSize, getFileIcon } = useDocVault()

const textContent = ref('')
const pdfBlobUrl = ref<string>('')

// 当前格式状态
const currentFontFamily = ref('')
const currentFontSize = ref('')
const currentHeading = ref('')

// 编辑模式
const isEditing = ref(false)

// 防止循环更新的标志（必须在 useEditor 之前定义）
let isUpdatingFromWatch = false

// TipTap 编辑器实例
const editor = useEditor({
  extensions: [
    StarterKit,
    Underline,
    TextAlign.configure({ types: ['heading', 'paragraph'] }),
    TextStyle,
    Color,
    FontFamily,
    Highlight.configure({ multicolor: true }),
    Image,
    Table.configure({ resizable: true }),
    TableRow,
    TableCell,
    TableHeader,
  ],
  content: '',
  onUpdate: ({ editor }) => {
    // 同步更新 parsedContent，避免循环
    if (!isUpdatingFromWatch) {
      parsedContent.value = { type: 'html', content: editor.getHTML() }
    }
  },
  onSelectionUpdate: ({ editor }) => {
    // 更新当前格式状态
    const styleAttrs = editor.getAttributes('textStyle')
    currentFontFamily.value = styleAttrs.fontFamily || ''
    currentFontSize.value = styleAttrs.fontSize ? String(styleAttrs.fontSize).replace('px', '') : ''
    currentHeading.value = editor.isActive('heading') ? String(editor.getAttributes('heading').level) : ''
  },
})

// 工具栏方法
function setFontFamily(font: string) {
  currentFontFamily.value = font
  if (font) {
    editor.value?.chain().focus().setFontFamily(font).run()
  } else {
    editor.value?.chain().focus().unsetFontFamily().run()
  }
}

function setFontSize(size: string) {
  currentFontSize.value = size
  if (size) {
    const attrs = editor.value?.getAttributes('textStyle') || {}
    editor.value?.chain().focus().setMark('textStyle', { ...attrs, fontSize: size + 'px' }).run()
  } else {
    editor.value?.chain().focus().unsetMark('textStyle').run()
  }
}

function setHeading(level: string) {
  currentHeading.value = level
  if (level) {
    editor.value?.chain().focus().toggleHeading({ level: parseInt(level) as 1|2|3 }).run()
  } else {
    editor.value?.chain().focus().setParagraph().run()
  }
}

function setTextColor(event: Event) {
  const target = event.target as HTMLInputElement
  editor.value?.chain().focus().setColor(target.value).run()
}

function setHighlightColor(event: Event) {
  const target = event.target as HTMLInputElement
  editor.value?.chain().focus().toggleHighlight({ color: target.value }).run()
}

function insertImage() {
  const url = prompt('请输入图片URL:')
  if (url) {
    editor.value?.chain().focus().setImage({ src: url }).run()
  }
}

function insertTable() {
  editor.value?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
}

// 编辑模式控制
function startEditing() {
  isEditing.value = true
  if (editor.value) {
    editor.value.commands.setContent(parsedContent.value.content)
  }
}

function stopEditing() {
  isEditing.value = false
  if (editor.value) {
    parsedContent.value = { type: 'html', content: editor.value.getHTML() }
  }
}

// 监听 Word 内容变化，同步到编辑器
watch(
  () => parsedContent.value,
  (content) => {
    if (content.type === 'html' && editor.value && !isUpdatingFromWatch) {
      // 仅在内容不同时更新，避免循环
      if (content.content !== editor.value.getHTML()) {
        isUpdatingFromWatch = true
        editor.value.commands.setContent(content.content)
        isUpdatingFromWatch = false
      }
    }
  }
)

// PDF blob URL 监听
watch(
  () => parsedContent.value,
  (content) => {
    if (content.type === 'pdf' && content.content) {
      const binaryString = atob(content.content)
      const bytes = new Uint8Array(binaryString.length)
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i)
      }
      const blob = new Blob([bytes], { type: 'application/pdf' })
      pdfBlobUrl.value = URL.createObjectURL(blob)
    }
  }
)

// PDF blob URL 监听
watch(
  () => parsedContent.value,
  (content) => {
    if (content.type === 'pdf' && content.content) {
      const binaryString = atob(content.content)
      const bytes = new Uint8Array(binaryString.length)
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i)
      }
      const blob = new Blob([bytes], { type: 'application/pdf' })
      pdfBlobUrl.value = URL.createObjectURL(blob)
    }
  }
)

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

onBeforeUnmount(() => {
  editor.value?.destroy()
})
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
  align-items: center;
}

.toolbar-select {
  padding: 4px 8px;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  font-size: 12px;
  background: white;
  cursor: pointer;
}

.toolbar-select:hover {
  border-color: #9ca3af;
}

.toolbar-color {
  width: 28px;
  height: 28px;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  cursor: pointer;
  padding: 2px;
}

.toolbar-hint {
  font-size: 12px;
  color: #6b7280;
}

.btn.active {
  background: #4361ee;
  color: white;
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

/* TipTap 编辑器样式 */
.editor-content .ProseMirror {
  outline: none;
  min-height: 500px;
}

.editor-content .ProseMirror p {
  margin: 0 0 10px 0;
}

.editor-content .ProseMirror h1 {
  font-size: 2em;
  font-weight: bold;
  margin: 0.67em 0;
}

.editor-content .ProseMirror h2 {
  font-size: 1.5em;
  font-weight: bold;
  margin: 0.75em 0;
}

.editor-content .ProseMirror h3 {
  font-size: 1.17em;
  font-weight: bold;
  margin: 0.83em 0;
}

.editor-content .ProseMirror ul {
  list-style: disc;
  padding-left: 20px;
}

.editor-content .ProseMirror ol {
  list-style: decimal;
  padding-left: 20px;
}

.editor-content .ProseMirror table {
  border-collapse: collapse;
  width: 100%;
  margin: 10px 0;
}

.editor-content .ProseMirror table td,
.editor-content .ProseMirror table th {
  border: 1px solid #d1d5db;
  padding: 8px 12px;
}

.editor-content .ProseMirror table th {
  background: #f3f4f6;
  font-weight: 600;
}

.editor-content .ProseMirror img {
  max-width: 100%;
  height: auto;
}

/* docx-preview 渲染内容样式 */
.docx-preview-content {
  font-family: 'SimSun', '宋体', 'Times New Roman', serif;
  font-size: 12pt;
  line-height: 1.5;
  color: #000;
}

.docx-preview-content p {
  margin: 0 0 10px 0;
  text-indent: 2em;
}

.docx-preview-content h1,
.docx-preview-content h2,
.docx-preview-content h3,
.docx-preview-content h4,
.docx-preview-content h5,
.docx-preview-content h6 {
  margin: 1em 0 0.5em 0;
  font-weight: bold;
}

.docx-preview-content h1 { font-size: 2em; }
.docx-preview-content h2 { font-size: 1.5em; }
.docx-preview-content h3 { font-size: 1.17em; }

.docx-preview-content table {
  border-collapse: collapse;
  width: 100%;
  margin: 10px 0;
}

.docx-preview-content table td,
.docx-preview-content table th {
  border: 1px solid #000;
  padding: 5px 10px;
}

.docx-preview-content img {
  max-width: 100%;
  height: auto;
}

.docx-preview-content ul {
  list-style: disc;
  padding-left: 20px;
}

.docx-preview-content ol {
  list-style: decimal;
  padding-left: 20px;
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

.pdf-info {
  font-size: 12px;
  opacity: 0.8;
}

.pdf-scroll-area {
  flex: 1;
  overflow: auto;
  display: flex;
  justify-content: center;
  padding: 0;
}

.pdf-frame {
  width: 100%;
  height: 100%;
  border: none;
  background: white;
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
