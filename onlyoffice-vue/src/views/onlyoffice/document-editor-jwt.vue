<template>
  <onlyoffice-editor :loading="loading.editor" :config="editorConfig">
  </onlyoffice-editor>
</template>

<script>
import OnlyofficeEditor from './modules/onlyoffice-editor'
import { queryDocumentInfo } from '@/api/onlyoffice'

export default {
  data () {
    return {
      loading: {
        editor: false,
        save: false,
        forceSave: false
      },
      detail: {},
      editorConfig: {}
    }
  },
  components: {
    OnlyofficeEditor
  },
  created () {
    this.queryDocumentInfo()
  },
  methods: {
    // 获取文档配置信息
    queryDocumentInfo () {
      this.loading.editor = true
      
      // 从URL参数中获取key和加密设置
      const urlParams = new URLSearchParams(window.location.search)
      const key = urlParams.get('key') || 'test11.docx'
      const useJwtEncrypt = urlParams.get('useJwtEncrypt') || 'y'
      
      queryDocumentInfo({ key, useJwtEncrypt })
        .then(res => {
          const data = res.data || {}
          const { id, remarks } = data
          this.detail = { id, remarks }
          this.editorConfig = data.editorConfig
        })
        .finally(() => {
          this.loading.editor = false
        })
    },
  }
}
</script>
