<template>
  <div id="app">
    <!-- 文档上传页面 -->
    <document-upload v-if="currentPage === 'upload'" />
    
    <!-- OnlyOffice 编辑器页面 -->
    <document-editor-jwt v-else />
  </div>
</template>

<script>
import locale from 'ant-design-vue/lib/locale-provider/zh_CN'
import DocumentEditorJwt from './views/onlyoffice/document-editor-jwt.vue'
import DocumentUpload from './views/document-upload.vue'

export default {
  components: {
    DocumentEditorJwt,
    DocumentUpload
  },
  data() {
    return {
      locale,
      currentPage: 'upload'
    }
  },
  mounted() {
    this.determineCurrentPage()
    // 监听URL变化
    window.addEventListener('popstate', this.determineCurrentPage)
  },
  beforeDestroy() {
    window.removeEventListener('popstate', this.determineCurrentPage)
  },
  methods: {
    determineCurrentPage() {
      const path = window.location.pathname
      const urlParams = new URLSearchParams(window.location.search)
      
      // 如果URL包含key参数，显示编辑器页面
      if (urlParams.has('key') || path.includes('/onlyoffice')) {
        this.currentPage = 'editor'
      } else {
        this.currentPage = 'upload'
      }
    }
  }
}
</script>

<style lang="less">
body {
  padding: 0;
  margin: 0;
}
html,
body {
  font-size: 14px;
  color: #333;
  font-family: "Helvetica Neue", Helvetica, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "微软雅黑", Arial, sans-serif;
}

#app {
  min-height: 100vh;
}
</style>
