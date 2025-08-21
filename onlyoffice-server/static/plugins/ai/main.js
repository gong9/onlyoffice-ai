/* eslint-disable */

(function (window, undefined) {
  window.Asc.plugin.init = function (initData) {
    var me = this;

    // 添加工具栏菜单项
    this.executeMethod('AddToolbarMenuItem', [getToolbarItems()]);

    function getDocumentText(callback) {
      me.callCommand(
        function () {
          try {
            var doc = Api.GetDocument();
            var fullText = '';

            // 方法1: 尝试使用GetRange获取整个文档的文本
            try {
              var range = doc.GetRange();
              if (range && range.GetText) {
                fullText = range.GetText();
              }
            } catch (e) {
              console.log('GetRange方法失败，尝试其他方法');
            }
            console.log('文档文本长度:', fullText.length);
            return fullText;
          } catch (error) {
            console.error('获取文档文本失败:', error);
            return '';
          }
        },
        false,
        true,
        function (result) {
          if (callback) callback(result);
        },
      );
    }

    // 清除所有批注
    function clearAllComments(callback) {
      me.callCommand(
        function () {
          try {
            var doc = Api.GetDocument();

            // 获取文档中的所有批注
            var allComments = doc.GetAllComments();

            if (allComments && allComments.length > 0) {
              // 从后往前删除批注，避免索引问题
              for (var i = allComments.length - 1; i >= 0; i--) {
                var comment = allComments[i];
                if (comment && comment.Delete) {
                  comment.Delete();
                }
              }
              console.log('已清除 ' + allComments.length + ' 个批注');
            } else {
              console.log('文档中没有批注需要清除');
            }
          } catch (error) {
            console.error('清除批注失败:', error);
          }
        },
        false,
        true,
        function () {
          if (callback) callback();
        },
      );
    }

    // API请求方法
    function callCheckAPI(documentText, callback) {
      var xhr = new XMLHttpRequest();
      var url =
        'http://211.90.218.31:10000/zodiac-ym9aZ/prod-api/hzm/common_check/check';

      xhr.open('POST', url, true);
      xhr.setRequestHeader('Accept', 'application/json, text/plain, */*');
      xhr.setRequestHeader('Accept-Language', 'zh-CN,zh;q=0.9');
      xhr.setRequestHeader(
        'Authorization',
        'Bearer eyJhbGciOiJIUzUxMiJ9.eyJsb2dpbl91c2VyX2tleSI6ImQzN2I1Mzg3LWRiYzctNDc2Yi04MmYwLTFhNDllYjhjZTc5ZCJ9.RORlkskPERXvlZAV_4OmWbYupq87QSB4__WLoEbIuxi59H1AFcWQhQOIfDE7522KlrDD1TWouUakDyBbmprZrg',
      );
      xhr.setRequestHeader('Cache-Control', 'no-cache');
      xhr.setRequestHeader('Content-Type', 'application/json;charset=UTF-8');
      xhr.setRequestHeader('Pragma', 'no-cache');

      xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
          if (xhr.status === 200) {
            try {
              var response = JSON.parse(xhr.responseText);
              if (callback) callback(null, response);
            } catch (e) {
              if (callback) callback('解析响应数据失败: ' + e.message, null);
            }
          } else {
            if (callback) callback('API请求失败，状态码: ' + xhr.status, null);
          }
        }
      };

      xhr.onerror = function () {
        if (callback) callback('网络请求失败', null);
      };

      // 发送文档文本数据
      var requestData = {
        doc_parsing_content: documentText,
        edit_html_content: '',
        html: '',
        checkFunctions: 719,
        checkMode: 2,
      };

      try {
        xhr.send(JSON.stringify(requestData));
      } catch (e) {
        if (callback) callback('发送请求失败: ' + e.message, null);
      }
    }

    function addCommentToDocument() {
      clearAllComments(function () {
        me.callCommand(
          function () {
            var addedCount = 0;
            var targetRanges = [
              {
                startIndex: 98,
                endIndex: 106,
                comment: '标点符号错误',
                author: 'AI批注',
                id: 1,
              },
            ];

            try {
              var doc = Api.GetDocument();

              var sortedRanges = targetRanges.slice().sort(function (a, b) {
                return a.startIndex - b.startIndex;
              });

              function addCommentByCharacterIndexBatch() {
                var doc = Api.GetDocument();
                var count = doc.GetElementsCount();
                var globalCharIndex = 0;
                var processedCount = 0;

                const runOperationsGroupById = {};
                // 找到所有与当前Run重叠
                var overlappingComments = [];

                for (var i = 0; i < count; i++) {
                  var element = doc.GetElement(i);
                  if (!element || element.GetClassType() !== 'paragraph')
                    continue;

                  try {
                    var runsCount = element.GetElementsCount();

                    for (var j = 0; j < runsCount; j++) {
                      var run = element.GetElement(j);
                      if (run && run.GetClassType() === 'run') {
                        var text = run.GetText();
                        if (!text) continue;

                        var runStartIndex = globalCharIndex;
                        var visibleCharCount = 0;
                        for (var k = 0; k < text.length; k++) {
                          var char = text[k];
                          if (char === '\n' || char === '\r' || char === '\t')
                            continue;
                          visibleCharCount++;
                        }
                        var runEndIndex = globalCharIndex + visibleCharCount;

                        for (var r = 0; r < sortedRanges.length; r++) {
                          var range = sortedRanges[r];

                          // 当前run与目标range重叠
                          if (
                            runStartIndex < range.endIndex &&
                            runEndIndex > range.startIndex
                          ) {
                            var commentStartInRun = Math.max(
                              0,
                              range.startIndex - runStartIndex,
                            );

                            var commentEndInRun = Math.min(
                              visibleCharCount,
                              range.endIndex - runStartIndex,
                            );

                            overlappingComments.push({
                              startInRun: commentStartInRun,
                              endInRun: commentEndInRun,
                              comment: range.comment,
                              author: range.author,
                              originalRange: range,
                              id: range.id,
                              globalStartIndex: runStartIndex,
                              run: run,
                              text: text,
                              runIndex: j,
                            });
                          }
                        }

                        // console.log(JSON.stringify(text), '可见字符数:', visibleCharCount, '总长度:', text.length);

                        globalCharIndex += visibleCharCount;
                      }
                    }

                    console.log(globalCharIndex, 'globalCharIndex');

                    // // 批量执行这个段落的所有Run操作（从后往前，避免索引问题）
                    // runOperations.reverse();

                    // for (var op = 0; op < runOperations.length; op++) {
                    //   var operation = runOperations[op];
                    //   try {
                    //     processCommentOperation(element, operation);
                    //     console.log(operation, 'targetRanges');
                    //     processedCount += operation.comments.length;
                    //   } catch (e) {
                    //     console.error('批注操作失败:', e);
                    //   }
                    // }

                    // Object.keys(runOperationsGroupById).forEach(id => {
                    //   const operations = runOperationsGroupById[id]
                    //   processCommentOperation(element, operations)
                    // })
                  } catch (error) {
                    console.error('段落 ' + i + ' 处理失败:', error);
                  }
                }

                console.log(
                  '批量批注处理完成，处理了 ' + processedCount + ' 个批注',
                );
                mergeRunByCommentId(overlappingComments);

                return processedCount;
              }

              function mergeRunByCommentId(overlappingComments) {
                const mergedComments = [];

                overlappingComments.forEach((comment) => {
                  if (comment.id === comment.id) {
                    mergedComments.push(comment);
                  }
                });
              }

              function mergeRun(runOperationsGroupById) {}

              /**
               * handle run operation
               * @param  element
               * @param  operation
               * @returns
               */
              function processCommentOperation(element, operation) {
                var run = operation.run;
                var text = operation.text;
                var comments = operation.comments;
                var runIndex = operation.runIndex;

                if (
                  comments.length === 1 &&
                  comments[0].startInRun === 0 &&
                  comments[0].endInRun === text.length
                ) {
                  run.AddComment(
                    comments[0].comment +
                      `[${comments[0].originalRange.startIndex}-${comments[0].originalRange.endIndex}]`,
                    comments[0].author,
                  );

                  return;
                }

                try {
                  var mergedComments = mergeCommentRanges(comments);

                  var originalFormat = {};
                  try {
                    originalFormat.bold = run.GetBold ? run.GetBold() : false;
                    originalFormat.italic = run.GetItalic
                      ? run.GetItalic()
                      : false;
                    originalFormat.underline = run.GetUnderline
                      ? run.GetUnderline()
                      : false;
                    originalFormat.fontSize = run.GetFontSize
                      ? run.GetFontSize()
                      : null;
                    originalFormat.fontFamily = run.GetFontFamily
                      ? run.GetFontFamily()
                      : null;
                  } catch (e) {
                    console.log('格式获取失败:', e);
                  }

                  // 按照批注区间拆分文本
                  var segments = [];
                  var lastEnd = 0;

                  for (var c = 0; c < mergedComments.length; c++) {
                    var comment = mergedComments[c];

                    // 前段（无批注）
                    if (comment.startInRun > lastEnd) {
                      segments.push({
                        text: text.substring(lastEnd, comment.startInRun),
                        hasComment: false,
                      });
                    }

                    // 批注段
                    segments.push({
                      text: text.substring(
                        comment.startInRun,
                        comment.endInRun,
                      ),
                      hasComment: true,
                      comment: comment.comment,
                      author: comment.author,
                    });

                    lastEnd = comment.endInRun;
                  }

                  // 最后的无批注段
                  if (lastEnd < text.length) {
                    segments.push({
                      text: text.substring(lastEnd),
                      hasComment: false,
                    });
                  }

                  run.ClearContent();
                  if (segments[0] && segments[0].text) {
                    run.AddText(segments[0].text);
                  }

                  for (var s = 1; s < segments.length; s++) {
                    var segment = segments[s];
                    if (!segment.text) continue;

                    var newRun = Api.CreateRun();
                    newRun.AddText(segment.text);

                    try {
                      if (originalFormat.bold)
                        newRun.SetBold(originalFormat.bold);
                      if (originalFormat.italic)
                        newRun.SetItalic(originalFormat.italic);
                      if (originalFormat.underline)
                        newRun.SetUnderline(originalFormat.underline);
                      if (originalFormat.fontSize)
                        newRun.SetFontSize(originalFormat.fontSize);
                      if (originalFormat.fontFamily)
                        newRun.SetFontFamily(originalFormat.fontFamily);
                    } catch (e) {
                      console.error('格式应用失败:', e);
                    }

                    if (segment.hasComment) {
                      try {
                        newRun.AddComment(
                          segment.comment +
                            `[${segment.originalRange.startIndex}-${segment.originalRange.endIndex}]`,
                          segment.author,
                        );
                      } catch (e) {
                        console.error('精确批注失败:', e);
                      }
                    }

                    element.AddElement(newRun, runIndex + s);
                  }
                } catch (e) {
                  console.error('批注操作失败:', e);
                }
              }

              function mergeCommentRanges(comments) {
                if (comments.length <= 1) return comments;

                comments.sort(function (a, b) {
                  return a.startInRun - b.startInRun;
                });

                // var merged = [comments[0]];
                // for (var i = 1; i < comments.length; i++) {
                //   var current = comments[i];
                //   var last = merged[merged.length - 1];
                //
                //   if (current.startInRun <= last.endInRun) {
                //     // 重叠，合并
                //     last.endInRun = Math.max(last.endInRun, current.endInRun);
                //     last.comment += '; ' + current.comment; // 合并批注内容
                //   } else {
                //     // 不重叠，添加新区间
                //     merged.push(current);
                //   }
                // }

                return merged;
              }

              // 执行批量批注处理
              addedCount = addCommentByCharacterIndexBatch();
            } catch (error) {
              console.error('主要批注添加过程失败:', error);
            }
          },
          false,
          true,
          function () {
            alert('添加批注完成');
          },
        );
      });
    }

    // 定义工具栏项目
    function getToolbarItems() {
      return {
        guid: window.Asc.plugin.info.guid,
        tabs: [
          {
            id: 'proofread_tab',
            text: '智能校对',
            items: [
              {
                id: 'addComment',
                type: 'button',
                text: 'AI校对',
                hint: '对文档进行AI智能校对并添加批注',
                icons: 'icon.png',
                lockInViewMode: true,
                enableToggle: false,
                separator: false,
              },
              {
                id: 'getDocumentText',
                type: 'button',
                text: '获取文档文本',
                hint: '获取整个文档的文本内容',
                icons: 'icon.png',
                lockInViewMode: true,
                enableToggle: false,
                separator: false,
              },
              {
                id: 'clearAllComments',
                type: 'button',
                text: '清除所有批注',
                hint: '清除文档中的所有批注',
                icons: 'icon.png',
                lockInViewMode: true,
                enableToggle: false,
                separator: false,
              },
              {
                id: 'checkDocument',
                type: 'button',
                text: 'API校对',
                hint: '使用API接口对文档进行校对',
                icons: 'icon.png',
                lockInViewMode: true,
                enableToggle: false,
                separator: false,
              },
            ],
          },
        ],
      };
    }

    // 处理工具栏按钮点击事件
    this.attachToolbarMenuClickEvent('addComment', function (data) {
      // 执行批注添加逻辑
      addCommentToDocument();
    });

    // 处理获取文档文本按钮点击事件
    this.attachToolbarMenuClickEvent('getDocumentText', function (data) {
      // 获取文档文本
      getDocumentText(function (text) {
        if (text) {
          console.log('完整文档文本:', text);
          console.log('文档文本长度:', text.replace(/[\r\n]+/g, '').length);
        } else {
          alert('获取文档文本失败');
        }
      });
    });

    // 处理清除所有批注按钮点击事件
    this.attachToolbarMenuClickEvent('clearAllComments', function (data) {
      // 清除所有批注
      clearAllComments(function () {
        alert('所有批注已清除完成');
      });
    });

    // 处理API校对按钮点击事件
    this.attachToolbarMenuClickEvent('checkDocument', function (data) {
      // 先获取文档文本
      getDocumentText(function (documentText) {
        if (documentText) {
          // 调用API进行校对
          callCheckAPI(``, function (error, response) {
            if (error) {
              alert('API校对失败: ' + error);
            } else {
              console.log('API校对结果:', response);
              alert('API校对完成，请查看控制台日志获取详细结果');
            }
          });
        } else {
          alert('获取文档文本失败，无法进行API校对');
        }
      });
    });

    // ===========================================
    // 业务逻辑函数（保留原有的addCommentAction函数作为备用）
    // ===========================================

    // 添加批注
    function addCommentAction() {
      addCommentToDocument();
    }

    // ===========================================
    // 事件绑定（保留原有的点击事件绑定）
    // ===========================================

    // 开始添加批注
    $('#addText').click(addCommentAction);

    // ===========================================
    // 插件生命周期
    // ===========================================

    // 插件事件处理
    window.Asc.plugin.onExternalMouseUp = function () {
      var event = document.createEvent('MouseEvents');
      event.initMouseEvent(
        'mouseup',
        true,
        true,
        window,
        1,
        0,
        0,
        0,
        0,
        false,
        false,
        false,
        false,
        0,
        null,
      );
      document.dispatchEvent(event);
    };

    window.Asc.plugin.button = function (id) {
      if (id === -1) {
        this.executeCommand('close', '');
      }
    };
  };
})(window, undefined);
