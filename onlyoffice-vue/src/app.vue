<template>
    <document-editor-jwt />
</template>

<script>
import locale from 'ant-design-vue/lib/locale-provider/zh_CN'
import DocumentEditorJwt from './views/onlyoffice/document-editor-jwt.vue'

export default {
  components: {
    DocumentEditorJwt
  },
  data () {
    return {
      locale
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

</style>
