<template>
  <div class="docvault-app">
    <!-- 顶部工具栏 -->
    <header class="toolbar">
      <div class="toolbar-left">
        <h1 class="app-title">📚 DocVault</h1>
        <span class="app-subtitle">完全离线的文档管理</span>
      </div>
      <div class="toolbar-actions">
        <button
          :disabled="state.loading"
          class="btn btn-primary"
          @click="openFile"
        >
          📂 打开文件
        </button>
        <button
          :disabled="!state.currentFile || state.loading"
          class="btn btn-success"
          @click="saveFile"
        >
          💾 保存
        </button>
        <button
          :disabled="!state.currentFile || state.loading"
          class="btn btn-warning"
          @click="convertToPdf"
        >
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
      <div
        class="progress-fill"
        :style="{ width: state.conversionProgress + '%' }"
      ></div>
    </div>

    <!-- 主内容区 -->
    <main class="main-content">
      <!-- 左侧文件列表 -->
      <aside class="sidebar">
        <div class="sidebar-header">
          <h3>最近文件</h3>
          <button class="btn btn-small" @click="scanDir(getDefaultDir)">
            🔄 刷新
          </button>
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
            暂无文件，点击"打开文件"开始
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
              <span v-if="state.currentFile.modified">
                修改: {{ formatDate(state.currentFile.modified) }}
              </span>
            </div>
          </div>

          <!-- Word 编辑区 -->
          <div v-if="isWordFile" class="word-editor">
            <div class="editor-toolbar">
              <button class="btn btn-small" @click="execCmd('bold')"><b>B</b></button>
              <button class="btn btn-small" @click="execCmd('italic')"><i>I</i></button>
              <button class="btn btn-small" @click="execCmd('underline')"><u>U</u></button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="execCmd('justifyLeft')">⬅</button>
              <button class="btn btn-small" @click="execCmd('justifyCenter')">↔</button>
              <button class="btn btn-small" @click="execCmd='justifyRight'">➡</button>
              <span class="separator">|</span>
              <button class="btn btn-small" @click="execCmd('insertUnorderedList')">• 列表</button>
            </div>
            <div
              ref="editorRef"
              class="editor-content"
              contenteditable="true"
              @input="onEditorInput"
              v-html="wordContent"
            ></div>
          </div>

          <!-- Excel 编辑区 -->
          <div v-else-if="isExcelFile" class="excel-editor">
            <div id="luckysheet-container" class="spreadsheet-container"></div>
            <div class="placeholder">
              <p>📊 Excel 编辑组件</p>
              <p>集成 Luckysheet/FortuneSheet 进行表格编辑</p>
              <pre>{{ state.currentFile.content_base64.substring(0, 100) }}...</pre>
            </div>
          </div>

          <!-- PDF 预览区 -->
          <div v-else-if="isPdfFile" class="pdf-viewer">
            <iframe
              :src="pdfUrl"
              class="pdf-frame"
              title="PDF Viewer"
            ></iframe>
          </div>

          <!-- 其他格式 -->
          <div v-else class="unsupported">
            <p>暂不支持该格式的在线编辑</p>
            <p>文件已加载，可尝试"Word转PDF"功能</p>
          </div>
        </div>

        <!-- 空状态 -->
        <div v-else class="empty-editor">
          <div class="empty-content">
            <span class="empty-icon">📄</span>
            <h3>欢迎使用 DocVault</h3>
            <p>完全离线、跨平台的文档管理解决方案</p>
            <ul class="feature-list">
              <li>📝 Word 文档编辑与转换</li>
              <li>📊 Excel 表格处理</li>
              <li>📕 PDF 查看与生成</li>
              <li>🔒 完全离线，数据安全</li>
            </ul>
            <button class="btn btn-primary btn-large" @click="openFile">
              打开文件开始
            </button>
          </div>
        </div>
      </section>
    </main>

    <!-- 底部状态栏 -->
    <footer class="status-bar">
      <span v-if="state.loading" class="status-loading">
        ⏳ 处理中...
      </span>
      <span v-else class="status-ready">
        ✅ 就绪
      </span>
      <span class="status-info">
        {{ state.currentFile ? state.currentFile.name : '未打开文件' }}
      </span>
      <span class="status-platform">
        {{ platform }}
      </span>
    </footer>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { platform as detectPlatform } from '@tauri-apps/plugin-os'
import { useDocVault } from '../composables/useDocVault'

const { state, openFile, saveFile, convertToPdf, scanDir, formatSize, getFileIcon } = useDocVault()

const editorRef = ref<HTMLDivElement | null>(null)
const wordContent = ref('<p>打开 Word 文档开始编辑...</p>')
const platform = ref('unknown')

// 计算属性
const isWordFile = computed(() =>
  ['docx', 'doc'].includes(state.currentFile?.extension || '')
)
const isExcelFile = computed(() =>
  ['xlsx', 'xls'].includes(state.currentFile?.extension || '')
)
const isPdfFile = computed(() =>
  state.currentFile?.extension === 'pdf'
)
const pdfUrl = computed(() => {
  if (!state.currentFile || !isPdfFile.value) return ''
  return `data:application/pdf;base64,${state.currentFile.content_base64}`
})

const getDefaultDir = ''

// 方法
async function openFileByPath(path: string) {
  try {
    state.loading = true
    const fileData = await invoke<FileData>('open_local_file', { path })
    state.currentFile = fileData

    if (isWordFile.value && fileData.extension === 'docx') {
      // 简单解析 docx 内容显示
      wordContent.value = '<p>Word 文档已加载，点击编辑...</p>'
    }
  } catch (err) {
    state.error = `打开失败: ${err}`
  } finally {
    state.loading = false
  }
}

function execCmd(command: string, value?: string) {
  document.execCommand(command, false, value)
}

function onEditorInput() {
  if (editorRef.value) {
    wordContent.value = editorRef.value.innerHTML
  }
}

function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleString('zh-CN')
}

onMounted(async () => {
  try {
    platform.value = await detectPlatform()
  } catch {
    platform.value = 'desktop'
  }
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
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
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

.btn-primary {
  background: #4361ee;
  color: white;
}
.btn-primary:hover:not(:disabled) {
  background: #3a56d4;
}

.btn-success {
  background: #06d6a0;
  color: white;
}
.btn-success:hover:not(:disabled) {
  background: #05c291;
}

.btn-warning {
  background: #ff9e00;
  color: white;
}
.btn-warning:hover:not(:disabled) {
  background: #e68f00;
}

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
}

.alert-error {
  background: #fee2e2;
  color: #dc2626;
}

.alert-success {
  background: #dcfce7;
  color: #16a34a;
}

.alert-close {
  background: none;
  border: none;
  font-size: 18px;
  cursor: pointer;
  opacity: 0.7;
}

.progress-bar {
  height: 3px;
  background: #e5e7eb;
}

.progress-fill {
  height: 100%;
  background: #4361ee;
  transition: width 0.3s ease;
}

.main-content {
  display: flex;
  flex: 1;
  overflow: hidden;
}

.sidebar {
  width: 280px;
  background: white;
  border-right: 1px solid #e5e7eb;
  display: flex;
  flex-direction: column;
}

.sidebar-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px;
  border-bottom: 1px solid #e5e7eb;
}

.sidebar-header h3 {
  margin: 0;
  font-size: 14px;
  color: #374151;
}

.file-list {
  flex: 1;
  overflow-y: auto;
  padding: 8px;
}

.file-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.15s;
}

.file-item:hover {
  background: #f3f4f6;
}

.file-item.active {
  background: #eef2ff;
  border: 1px solid #c7d2fe;
}

.file-icon {
  font-size: 20px;
}

.file-info {
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.file-name {
  font-size: 13px;
  color: #1f2937;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.file-meta {
  font-size: 11px;
  color: #9ca3af;
}

.empty-state {
  text-align: center;
  padding: 40px 20px;
  color: #9ca3af;
  font-size: 13px;
}

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
}

.doc-header {
  padding: 16px 24px;
  background: white;
  border-bottom: 1px solid #e5e7eb;
}

.doc-header h2 {
  margin: 0 0 8px 0;
  font-size: 18px;
  color: #1f2937;
}

.doc-meta {
  display: flex;
  gap: 16px;
  font-size: 12px;
  color: #6b7280;
}

.word-editor {
  flex: 1;
  display: flex;
  flex-direction: column;
  background: white;
}

.editor-toolbar {
  display: flex;
  gap: 4px;
  padding: 8px 16px;
  border-bottom: 1px solid #e5e7eb;
  background: #f9fafb;
}

.separator {
  color: #d1d5db;
  padding: 0 4px;
}

.editor-content {
  flex: 1;
  padding: 40px;
  max-width: 800px;
  margin: 0 auto;
  outline: none;
  line-height: 1.8;
  color: #374151;
}

.excel-editor,
.pdf-viewer {
  flex: 1;
  position: relative;
}

.spreadsheet-container {
  width: 100%;
  height: 100%;
}

.pdf-frame {
  width: 100%;
  height: 100%;
  border: none;
}

.placeholder {
  padding: 40px;
  text-align: center;
  color: #6b7280;
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

.empty-content {
  text-align: center;
  max-width: 400px;
}

.empty-icon {
  font-size: 64px;
}

.empty-content h3 {
  margin: 16px 0 8px;
  color: #1f2937;
}

.empty-content p {
  color: #6b7280;
  margin-bottom: 24px;
}

.feature-list {
  list-style: none;
  padding: 0;
  text-align: left;
  margin-bottom: 32px;
}

.feature-list li {
  padding: 8px 0;
  color: #4b5563;
  font-size: 14px;
}

.btn-large {
  padding: 12px 24px;
  font-size: 16px;
}

.status-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 6px 24px;
  background: #1a1a2e;
  color: white;
  font-size: 12px;
}

.status-loading {
  color: #fbbf24;
}

.status-ready {
  color: #34d399;
}

.status-info {
  opacity: 0.8;
}

.status-platform {
  opacity: 0.6;
  font-size: 11px;
}
</style>
