<template>
  <div class="document-upload-page">
    <div class="page-header">
      <h1>文档上传</h1>
      <p>上传您的 Office 文档并开始编辑</p>
    </div>

    <!-- 文件上传区域 -->
    <div class="upload-section">
      <div class="upload-area" :class="{ 'drag-over': isDragOver }">
        <input
          ref="fileInput"
          type="file"
          accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.pdf"
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
          <div class="upload-content">
            <div class="upload-icon">
              <svg width="32" height="24" viewBox="0 0 32 24" fill="none">
                <path d="M29.055 14.6413C28.7261 16.692 26.8905 18.2041 24.7346 18.2002C23.9314 18.2002 23.2802 18.827 23.2802 19.6001C23.2802 20.3732 23.9314 21 24.7346 21C28.3201 20.9951 31.3672 18.4766 31.9142 15.066C32.4612 11.6554 30.3449 8.36972 26.9279 7.32439C25.2403 2.9231 20.8822 0 16.008 0C11.1337 0 6.77563 2.9231 5.08799 7.32439C1.66275 8.36286 -0.462096 11.6523 0.0857665 15.0683C0.633629 18.4844 3.69014 21.004 7.2813 21C8.08456 21 8.73574 20.3732 8.73574 19.6001C8.73574 18.827 8.08456 18.2002 7.2813 18.2002C5.13116 18.1951 3.3049 16.6842 2.97684 14.6389C2.64878 12.5936 3.91658 10.6229 5.96502 9.994L7.32348 9.58103L7.81653 8.29312C9.08264 4.99214 12.3514 2.79991 16.0072 2.79991C19.6631 2.79991 22.9318 4.99214 24.1979 8.29312L24.6924 9.58103L26.0509 9.994C28.1076 10.6161 29.3839 12.5906 29.055 14.6413ZM17.4286 23.2906L17.4286 15.4868L19.7143 15.4868C19.8225 15.4868 19.9214 15.426 19.9698 15.3299C20.0182 15.2338 20.0078 15.1187 19.9429 15.0327L16.2286 10.1135C16.1746 10.0421 16.0899 10 16 10C15.9101 10 15.8254 10.0421 15.7714 10.1135L12.0571 15.0327C11.9922 15.1187 11.9818 15.2338 12.0302 15.3299C12.0786 15.426 12.1775 15.4868 12.2857 15.4868L14.5714 15.4868L14.5714 23.2906C14.5714 23.4787 14.6467 23.6592 14.7806 23.7922C14.9146 23.9253 15.0963 24 15.2857 24L16.7143 24C16.9037 24 17.0854 23.9253 17.2194 23.7922C17.3533 23.6592 17.4286 23.4787 17.4286 23.2906Z" fill="#5983FF"/>
              </svg>
            </div>
            <div class="upload-text">
              <p class="upload-hint">点击或将文件拖拽到这里上传</p>
            </div>
            <button class="upload-btn" @click.stop="triggerFileSelect">
              上传文档
            </button>
            <p class="format-hint">支持格式：docx、xlsx、pptx、doc、xls、ppt、pdf，单个文件不能超过50MB</p>
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
  </div>
</template>

<script>
import { uploadDocument } from '../api/document'

export default {
  name: 'DocumentUpload',
  data() {
    return {
      isDragOver: false,
      uploadingFiles: [],
      fileIdCounter: 0
    }
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
      
      // 只处理第一个文件，上传成功后直接跳转
      if (validFiles.length > 0) {
        await this.uploadSingleFile(validFiles[0])
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
          this.$message.success(`${file.name} 上传成功，正在跳转到编辑页面...`)
          
          // 上传成功后直接跳转到编辑页面
          setTimeout(() => {
            const key = response.data.data.key || response.data.data.fileName
            const editorUrl = `/onlyoffice?key=${key}&useJwtEncrypt=y`
            window.location.href = editorUrl
          }, 1500) // 延迟1.5秒让用户看到成功消息
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
    
    formatFileSize(bytes) {
      if (bytes === 0) return '0 B'
      const k = 1024
      const sizes = ['B', 'KB', 'MB', 'GB']
      const i = Math.floor(Math.log(bytes) / Math.log(k))
      return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i]
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
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  align-items: center;
}

/* 在大屏幕上保留适当的边距 */
@media (min-width: 1600px) {
  .document-upload-page {
    /* max-width: 1400px; */
    margin: 0 auto;
  }
}

.page-header {
  text-align: center;
  margin-bottom: 48px;
}

.page-header h1 {
  font-size: 28px;
  color: #101219;
  margin-bottom: 8px;
  font-weight: 600;
}

.page-header p {
  color: #686F82;
  font-size: 16px;
}

.upload-section {
  background: white;
  border-radius: 10px;
  padding: 40px;
  margin-bottom: 40px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.1);
  width: 100%;
  max-width: 800px;
}

.upload-section h3 {
  margin-bottom: 16px;
  color: #333;
}

.upload-area {
  border: 1px dashed rgba(89, 131, 255, 0.75);
  border-radius: 8px;
  transition: all 0.3s;
  background: white;
}

.upload-area.drag-over {
  border-color: #5983FF;
  background-color: rgba(89, 131, 255, 0.05);
}

.drop-zone {
  padding: 48px 80px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s;
  min-height: 300px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.drop-zone:hover {
  background-color: rgba(89, 131, 255, 0.02);
}

.upload-content {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 20px;
  width: 100%;
}

.upload-icon {
  margin-bottom: 20px;
}

.upload-text {
  margin-bottom: 20px;
}

.upload-hint {
  font-size: 14px;
  color: #686F82;
  margin: 0;
  font-family: "Source Han Sans", "PingFang SC", sans-serif;
}

.upload-btn {
  background: #5983FF;
  color: white;
  border: none;
  border-radius: 6px;
  padding: 8px 21px;
  height: 40px;
  font-size: 16px;
  font-weight: 500;
  cursor: pointer;
  transition: background 0.3s;
  font-family: "PingFangSC", "PingFang SC", sans-serif;
}

.upload-btn:hover {
  background: #4A6FE6;
}

.format-hint {
  font-size: 14px;
  color: #979DAD;
  margin: 0;
  font-family: "Source Han Sans", "PingFang SC", sans-serif;
}

.upload-progress {
  margin-top: 24px;
  padding: 16px;
  background: #f9f9f9;
  border-radius: 6px;
}

.upload-progress h4 {
  margin-bottom: 12px;
  color: #101219;
  font-weight: 600;
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
  color: #101219;
}

.file-size {
  font-size: 12px;
  color: #979DAD;
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
  background: #5983FF;
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
  color: #686F82;
}



/* 响应式设计 */
@media (max-width: 768px) {
  .document-upload-page {
    padding: 16px;
  }
  
  .upload-section {
    padding: 24px;
  }
  
  .drop-zone {
    padding: 32px 24px;
    min-height: 250px;
  }
  
  .upload-content {
    gap: 16px;
  }
}

@media (max-width: 480px) {
  .drop-zone {
    padding: 24px 16px;
    min-height: 200px;
  }
  
  .upload-section {
    padding: 20px;
  }
}
</style> 