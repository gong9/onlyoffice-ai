<template>
  <div class="document-upload-page">
    <div class="page-header">
      <h1>文档管理</h1>
      <p>上传和管理您的 Office 文档</p>
    </div>

    <!-- 文件上传区域 -->
    <div class="upload-section">
      <h3>上传新文档</h3>
      <div class="upload-area" :class="{ 'drag-over': isDragOver }">
        <input
          ref="fileInput"
          type="file"
          accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.pdf"
          multiple
          @change="handleFileSelect"
          style="display: none"
        />
        
        <div 
          class="drop-zone"
          @click="triggerFileSelect"
          @dragover="handleDragOver"
          @dragleave="handleDragLeave"
          @drop="handleDrop"
        >
          <div class="upload-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
              <polyline points="10,9 9,9 8,9"/>
            </svg>
          </div>
          <div class="upload-text">
            <p><strong>点击上传</strong> 或拖拽文件到此区域</p>
            <p class="upload-hint">支持 .docx, .xlsx, .pptx, .doc, .xls, .ppt, .pdf 格式，单个文件不超过 50MB</p>
          </div>
        </div>
      </div>

      <!-- 上传进度 -->
      <div v-if="uploadingFiles.length > 0" class="upload-progress">
        <h4>上传进度</h4>
        <div v-for="file in uploadingFiles" :key="file.id" class="progress-item">
          <div class="file-info">
            <span class="file-name">{{ file.name }}</span>
            <span class="file-size">{{ formatFileSize(file.size) }}</span>
          </div>
          <div class="progress-bar">
            <div 
              class="progress-fill" 
              :style="{ width: file.progress + '%' }"
              :class="{ 'success': file.status === 'success', 'error': file.status === 'error' }"
            ></div>
          </div>
          <div class="progress-text">
            {{ file.status === 'uploading' ? file.progress + '%' : 
                file.status === 'success' ? '上传成功' : 
                file.status === 'error' ? '上传失败' : '等待上传' }}
          </div>
        </div>
      </div>
    </div>

    <!-- 文档列表 -->
    <div class="document-list-section">
      <div class="list-header">
        <h3>已上传文档</h3>
        <button class="refresh-btn" @click="refreshDocumentList" :disabled="loading">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="23 4 23 10 17 10"/>
            <polyline points="1 20 1 14 7 14"/>
            <path d="m3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
          </svg>
          刷新
        </button>
      </div>

      <div v-if="loading" class="loading">
        <div class="spinner"></div>
        <p>加载中...</p>
      </div>

      <div v-else-if="documents.length === 0" class="empty-state">
        <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14,2 14,8 20,8"/>
        </svg>
        <p>暂无文档</p>
        <p class="empty-hint">点击上方上传区域开始添加文档</p>
      </div>

      <div v-else class="document-grid">
        <div 
          v-for="doc in documents" 
          :key="doc.key" 
          class="document-card"
          @click="openDocument(doc)"
        >
          <div class="doc-icon">
            <svg v-if="doc.documentType === 'word'" width="32" height="32" viewBox="0 0 24 24" fill="#2B579A">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <path d="M10 12h4M10 16h4"/>
            </svg>
            <svg v-else-if="doc.documentType === 'cell'" width="32" height="32" viewBox="0 0 24 24" fill="#217346">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <rect x="8" y="11" width="8" height="6" fill="none" stroke="white" stroke-width="1"/>
              <line x1="8" y1="13" x2="16" y2="13" stroke="white" stroke-width="1"/>
              <line x1="12" y1="11" x2="12" y2="17" stroke="white" stroke-width="1"/>
            </svg>
            <svg v-else-if="doc.documentType === 'slide'" width="32" height="32" viewBox="0 0 24 24" fill="#D24726">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <rect x="8" y="12" width="8" height="4" fill="white"/>
            </svg>
            <svg v-else-if="doc.documentType === 'pdf'" width="32" height="32" viewBox="0 0 24 24" fill="#FF6B6B">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <text x="12" y="16" text-anchor="middle" fill="white" font-size="8" font-weight="bold">PDF</text>
            </svg>
            <svg v-else width="32" height="32" viewBox="0 0 24 24" fill="#999">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14,2 14,8 20,8"/>
              <rect x="8" y="12" width="8" height="4" fill="white"/>
            </svg>
          </div>
          <div class="doc-info">
            <h4 class="doc-name" :title="doc.fileName">{{ doc.fileName }}</h4>
            <p class="doc-meta">
              <span>{{ formatFileSize(doc.fileSize) }}</span>
              <span>{{ formatDate(doc.uploadTime) }}</span>
            </p>
          </div>
          <div class="doc-actions">
            <button class="action-btn" @click.stop="downloadDocument(doc)" title="下载">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
            </button>
            <button class="action-btn edit-btn" @click.stop="openDocument(doc)" title="编辑">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
              </svg>
            </button>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { uploadDocument, getDocumentList } from '../api/document'

export default {
  name: 'DocumentUpload',
  data() {
    return {
      isDragOver: false,
      uploadingFiles: [],
      documents: [],
      loading: false,
      fileIdCounter: 0
    }
  },
  mounted() {
    this.refreshDocumentList()
  },
  methods: {
    triggerFileSelect() {
      this.$refs.fileInput.click()
    },
    
    handleFileSelect(event) {
      const files = Array.from(event.target.files)
      this.uploadFiles(files)
      event.target.value = '' // 清空input，允许重复选择同一文件
    },
    
    handleDragOver(event) {
      event.preventDefault()
      this.isDragOver = true
    },
    
    handleDragLeave(event) {
      event.preventDefault()
      this.isDragOver = false
    },
    
    handleDrop(event) {
      event.preventDefault()
      this.isDragOver = false
      const files = Array.from(event.dataTransfer.files)
      this.uploadFiles(files)
    },
    
    async uploadFiles(files) {
      const validFiles = files.filter(file => {
        const validExtensions = ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.pdf']
        const extension = '.' + file.name.split('.').pop().toLowerCase()
        return validExtensions.includes(extension)
      })
      
      if (validFiles.length !== files.length) {
        this.$message.warning('部分文件格式不支持，已跳过')
      }
      
      for (const file of validFiles) {
        await this.uploadSingleFile(file)
      }
    },
    
    async uploadSingleFile(file) {
      const fileItem = {
        id: ++this.fileIdCounter,
        name: file.name,
        size: file.size,
        progress: 0,
        status: 'uploading'
      }
      
      this.uploadingFiles.push(fileItem)
      
      const formData = new FormData()
      formData.append('file', file)
      formData.append('fileName', file.name)
      
      try {
        const response = await uploadDocument(formData)
        
        if (response.data.code === 200) {
          fileItem.status = 'success'
          fileItem.progress = 100
          this.$message.success(`${file.name} 上传成功`)
          // 上传成功后刷新文档列表
          setTimeout(() => {
            this.refreshDocumentList()
          }, 1000)
        } else {
          throw new Error(response.data.message)
        }
      } catch (error) {
        fileItem.status = 'error'
        this.$message.error(`${file.name} 上传失败: ${error.response?.data?.message || error.message}`)
      }
      
      // 3秒后移除上传进度项
      setTimeout(() => {
        const index = this.uploadingFiles.findIndex(item => item.id === fileItem.id)
        if (index !== -1) {
          this.uploadingFiles.splice(index, 1)
        }
      }, 3000)
    },
    
    async refreshDocumentList() {
      this.loading = true
      try {
        const response = await getDocumentList()
        if (response.data.code === 200) {
          this.documents = response.data.data.files || []
        } else {
          this.$message.error('获取文档列表失败')
        }
      } catch (error) {
        this.$message.error('获取文档列表失败: ' + (error.response?.data?.message || error.message))
      } finally {
        this.loading = false
      }
    },
    
    openDocument(doc) {
      // 跳转到文档编辑页面，使用JWT加密
      const editorUrl = `/onlyoffice?key=${doc.key}&useJwtEncrypt=y`
      window.open(editorUrl, '_blank')
    },
    
    downloadDocument(doc) {
      // 下载文档
      const link = document.createElement('a')
      link.href = doc.url
      link.download = doc.fileName
      link.click()
    },
    
    formatFileSize(bytes) {
      if (bytes === 0) return '0 B'
      const k = 1024
      const sizes = ['B', 'KB', 'MB', 'GB']
      const i = Math.floor(Math.log(bytes) / Math.log(k))
      return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i]
    },
    
    formatDate(dateString) {
      const date = new Date(dateString)
      return date.toLocaleDateString('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      })
    }
  }
}
</script>

<style scoped>
.document-upload-page {
  width: 100%;
  max-width: none;
  margin: 0;
  padding: 24px;
  background: #f5f5f5;
  height: 100vh;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

/* 在大屏幕上保留适当的边距 */
@media (min-width: 1600px) {
  .document-upload-page {
    max-width: 1400px;
    margin: 0 auto;
  }
}

.page-header {
  text-align: center;
  margin-bottom: 32px;
}

.page-header h1 {
  font-size: 28px;
  color: #333;
  margin-bottom: 8px;
}

.page-header p {
  color: #666;
  font-size: 16px;
}

.upload-section {
  background: white;
  border-radius: 8px;
  padding: 24px;
  margin-bottom: 24px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}

.upload-section h3 {
  margin-bottom: 16px;
  color: #333;
}

.upload-area {
  border: 2px dashed #d9d9d9;
  border-radius: 8px;
  transition: all 0.3s;
}

.upload-area.drag-over {
  border-color: #1890ff;
  background-color: #f0f9ff;
}

.drop-zone {
  padding: 48px 24px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s;
}

.drop-zone:hover {
  background-color: #fafafa;
}

.upload-icon {
  color: #8c8c8c;
  margin-bottom: 16px;
}

.upload-text p {
  margin: 8px 0;
  color: #666;
}

.upload-text .upload-hint {
  font-size: 14px;
  color: #999;
}

.upload-progress {
  margin-top: 24px;
  padding: 16px;
  background: #f9f9f9;
  border-radius: 6px;
}

.upload-progress h4 {
  margin-bottom: 12px;
  color: #333;
}

.progress-item {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 12px;
}

.file-info {
  min-width: 200px;
}

.file-name {
  font-weight: 500;
  color: #333;
}

.file-size {
  font-size: 12px;
  color: #999;
  margin-left: 8px;
}

.progress-bar {
  flex: 1;
  height: 8px;
  background: #f0f0f0;
  border-radius: 4px;
  overflow: hidden;
}

.progress-fill {
  height: 100%;
  background: #1890ff;
  transition: width 0.3s;
}

.progress-fill.success {
  background: #52c41a;
}

.progress-fill.error {
  background: #ff4d4f;
}

.progress-text {
  min-width: 80px;
  text-align: right;
  font-size: 12px;
  color: #666;
}

.document-list-section {
  background: white;
  border-radius: 8px;
  padding: 24px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.list-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.list-header h3 {
  color: #333;
  margin: 0;
}

.refresh-btn {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 16px;
  background: #1890ff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.3s;
}

.refresh-btn:hover:not(:disabled) {
  background: #40a9ff;
}

.refresh-btn:disabled {
  background: #d9d9d9;
  cursor: not-allowed;
}

.loading {
  text-align: center;
  padding: 48px;
  color: #666;
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
}

.spinner {
  width: 32px;
  height: 32px;
  border: 3px solid #f0f0f0;
  border-top: 3px solid #1890ff;
  border-radius: 50%;
  animation: spin 1s linear infinite;
  margin: 0 auto 16px;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

.empty-state {
  text-align: center;
  padding: 48px;
  color: #999;
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
}

.empty-state svg {
  margin-bottom: 16px;
  opacity: 0.5;
}

.empty-hint {
  font-size: 14px;
  margin-top: 8px;
}

.document-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 16px;
  overflow-y: auto;
  flex: 1;
  padding: 0 4px; /* 为滚动条留出空间 */
}

/* 针对不同屏幕尺寸优化网格布局 */
@media (min-width: 768px) {
  .document-grid {
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  }
}

@media (min-width: 1200px) {
  .document-grid {
    grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  }
}

@media (min-width: 1600px) {
  .document-grid {
    grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  }
}

.document-card {
  border: 1px solid #f0f0f0;
  border-radius: 8px;
  padding: 16px;
  cursor: pointer;
  transition: all 0.3s;
  background: white;
}

.document-card:hover {
  border-color: #1890ff;
  box-shadow: 0 4px 12px rgba(24,144,255,0.15);
}

.doc-icon {
  margin-bottom: 12px;
}

.doc-info {
  margin-bottom: 12px;
}

.doc-name {
  font-size: 16px;
  font-weight: 500;
  color: #333;
  margin: 0 0 8px 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.doc-meta {
  font-size: 12px;
  color: #999;
  margin: 0;
  display: flex;
  justify-content: space-between;
}

.doc-actions {
  display: flex;
  gap: 8px;
}

.action-btn {
  padding: 8px;
  border: 1px solid #d9d9d9;
  background: white;
  border-radius: 4px;
  cursor: pointer;
  color: #666;
  transition: all 0.3s;
}

.action-btn:hover {
  border-color: #1890ff;
  color: #1890ff;
}

.edit-btn:hover {
  background: #1890ff;
  color: white;
}

/* 自定义滚动条样式 */
.document-grid::-webkit-scrollbar {
  width: 8px;
}

.document-grid::-webkit-scrollbar-track {
  background: #f1f1f1;
  border-radius: 4px;
}

.document-grid::-webkit-scrollbar-thumb {
  background: #c1c1c1;
  border-radius: 4px;
}

.document-grid::-webkit-scrollbar-thumb:hover {
  background: #a8a8a8;
}

/* 为 Firefox 添加滚动条样式 */
.document-grid {
  scrollbar-width: thin;
  scrollbar-color: #c1c1c1 #f1f1f1;
}
</style> 