import request from './request'

/**
 * 上传文档文件
 * @param {FormData} formData 文件数据
 * @returns {Promise}
 */
export function uploadDocument(formData) {
  return request({
    url: '/api/v1/document/upload',
    method: 'post',
    data: formData,
    headers: {
      'Content-Type': 'multipart/form-data'
    }
  })
}

/**
 * 获取文档列表
 * @returns {Promise}
 */
export function getDocumentList() {
  return request({
    url: '/api/v1/document/list',
    method: 'get'
  })
}

/**
 * 获取文档信息
 * @param {Object} params 查询参数
 * @returns {Promise}
 */
export function getDocumentInfo(params) {
  return request({
    url: '/api/v1/document/info',
    method: 'get',
    params
  })
}

/**
 * 强制保存文档
 * @param {Object} data 保存数据
 * @returns {Promise}
 */
export function forceSaveDocument(data) {
  return request({
    url: '/api/v1/document/forceSave',
    method: 'post',
    data
  })
} 