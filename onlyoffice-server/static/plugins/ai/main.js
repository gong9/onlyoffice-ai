/* eslint-disable */

(function (window, undefined) {
  window.Asc.plugin.init = function (initData) {
    var me = this;

    // 添加工具栏菜单项
    this.executeMethod('AddToolbarMenuItem', [getToolbarItems()]);

    function getDocumentText(callback) {
      me.callCommand(
        function () {

          function showEditorLoading() {
            const container = parent.parent.parent.window.parent.window.parent.window[0].document
            let loader = container.getElementById('editorLoader');
            if (!loader) {
              loader = container.createElement('div');
              loader.id = 'editorLoader';
              
              // 创建简单的loading内容
              loader.innerHTML = `
                <div style="
                  position: fixed;
                  top: 0;
                  left: 0;
                  width: 100vw;
                  height: 100vh;
                  background: rgba(0,0,0,0.3);
                  display: flex;
                  justify-content: center;
                  align-items: center;
                  z-index: 999999;
                  font-family: Arial, sans-serif;
                  transition: opacity 0.5s ease-out, visibility 0.5s ease-out;
                  opacity: 1;
                  visibility: visible;
                ">
                  <div style="
                    text-align: center;
                    color: #333;
                    padding: 40px;
                    background: rgba(255,255,255,0.95);
                    border-radius: 15px;
                    border: 1px solid rgba(0,0,0,0.1);
                    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
                    transition: transform 0.5s ease-out;
                    transform: scale(1);
                  ">
                    <div style="
                      width: 60px;
                      height: 60px;
                      border: 4px solid rgba(74,111,230,0.3);
                      border-top: 4px solid #4A6FE6;
                      border-radius: 50%;
                      margin: 0 auto 20px;
                      animation: spin 1s linear infinite;
                    "></div>
                    <h3 style="margin: 0 0 15px 0; font-size: 20px; color: #333;">AI智能校对中...</h3>
                    <p style="margin: 0; font-size: 14px; color: #666;">正在对文档进行智能审查，请稍候</p>
                  </div>
                </div>
              `;
              
              // 添加简单的CSS动画
              const style = container.createElement('style');
              style.textContent = `
                @keyframes spin {
                  0% { transform: rotate(0deg); }
                  100% { transform: rotate(360deg); }
                }
              `;
              
              container.head.appendChild(style);
              container.body.appendChild(loader);
            } else {
              loader.style.opacity = '1';
              loader.style.visibility = 'visible';
              loader.style.display = 'flex';
              
              const titleElement = loader.querySelector('h3');
              const descElement = loader.querySelector('p');
              if (titleElement) {
                titleElement.textContent = 'AI智能校对中...';
              }
              if (descElement) {
                descElement.textContent = '正在对文档进行智能审查，请稍候';
              }
              
              // 重新显示旋转动画
              const spinner = loader.querySelector('div[style*="border-radius: 50%"]');
              if (spinner) {
                spinner.style.display = 'block';
                spinner.style.animation = 'spin 1s linear infinite';
                spinner.style.borderTop = '4px solid #4A6FE6';
              }
            }
            loader.style.display = 'flex';
          }

        showEditorLoading()

          try {
            var doc = Api.GetDocument();
            var fullText = '';

            try {
              var range = doc.GetRange();
              if (range && range.GetText) {
                fullText = range.GetText();
              }
            } catch (e) {
              console.log('GetRange方法失败，尝试其他方法');
            }
            console.log('文档文本长度:', fullText.replace(/\r/g, "").length);
            console.log(fullText.replace(/\r/g, ""));
            
            
            return fullText.replace(/\r/g, "");
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
        'http://211.90.218.31:10000/zodiac-ym9aZ/prod-api/hzm/audit/check';

      xhr.open('POST', url, true);
      xhr.setRequestHeader('Accept', 'application/json, text/plain, */*');
      xhr.setRequestHeader('Accept-Language', 'zh-CN,zh;q=0.9');
      xhr.setRequestHeader(
        'Authorization',
        'Bearer eyJhbGciOiJIUzUxMiJ9.eyJsb2dpbl91c2VyX2tleSI6ImY1NThjNDRkLWMzMjItNDQ5YS1hNDA5LWYxMjYxOTMyMTc5NyJ9.QKnOOSvPI2LDfiMlQJJgzAHEXsZLxI_x3OsTX7b1aFGXvophMqT5qE0YrMp228O_uTluJs96HLnwjh4Bgko3cQ',
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
        editText: documentText,
        tid: '20250827194717_c23fc8c35e5249b9ae34c2769ebbff0f',
      };

      try {
        xhr.send(JSON.stringify(requestData));
      } catch (e) {
        if (callback) callback('发送请求失败: ' + e.message, null);
      }
    }

    // SSE请求方法
    function callSSEAPI(tid, callback) {
      var url = 'http://211.90.218.31:10000/zodiac-ym9aZ/prod-api/hzm/sse/audit/result';
      
      var lastData = null;
      
      fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'text/event-stream',
          'Content-Type': 'application/json',
          'Authorization': 'Bearer eyJhbGciOiJIUzUxMiJ9.eyJsb2dpbl91c2VyX2tleSI6ImY1NThjNDRkLWMzMjItNDQ5YS1hNDA5LWYxMjYxOTMyMTc5NyJ9.QKnOOSvPI2LDfiMlQJJgzAHEXsZLxI_x3OsTX7b1aFGXvophMqT5qE0YrMp228O_uTluJs96HLnwjh4Bgko3cQ'
        },
        body: JSON.stringify({ tid: tid })
      })
      .then(function(response) {
        if (!response.ok) {
          throw new Error('HTTP ' + response.status);
        }
        
        var reader = response.body.getReader();
        var decoder = new TextDecoder();
        var buffer = '';
        
        async function readStream() {
          while (true) {
            const result = await reader.read();
            if (result.done) {
              console.log('SSE流结束，最后数据:', lastData);
              if (callback) {
                callback(null, lastData ? [lastData] : []);
              }
              break;
            }

            const chunk = decoder.decode(result.value, { stream: true });
            buffer += chunk;
            console.log('收到数据块:', chunk);
            console.log('当前buffer:', buffer);

            const lines = buffer.split('\n');
            buffer = lines.pop();

            let eventBuffer = [];
            for (let i = 0; i < lines.length; i++) {
              const line = lines[i].trim();

              if (line === '') {
                if (eventBuffer.length > 0) {
                  const dataStr = eventBuffer.join('');
                  eventBuffer = [];

                  if (dataStr === '[DONE]') {
                    if (callback) {
                      callback(null, lastData ? [lastData] : []);
                    }
                    return;
                  }

                  try {
                    const data = JSON.parse(dataStr);
                    lastData = data;
                    if (callback) {
                      callback(null, data, true);
                    }
                  } catch (e) {
                    console.error('解析失败:', e, '原始数据:', dataStr);
                  }
                }
              } else if (line.startsWith('data:')) {
                eventBuffer.push(line.substring(5).trim());
              } else if (line.startsWith('{') && line.endsWith('}')) {
                try {
                  const data = JSON.parse(line);
                  lastData = data;
                  if (callback) {
                    callback(null, data, true);
                  }
                } catch (e) {
                  console.error('直接解析JSON失败:', e, '原始数据:', line);
                }
              }
            }
          }
        }
        
        return readStream();
      })
      .catch(function(error) {
        console.error('SSE请求失败:', error);
        if (callback) {
          callback('SSE请求失败: ' + error.message, null);
        }
      });
      
      // 返回关闭函数
      return {
        close: function() {
          if (callback) {
            callback(null, lastData ? [lastData] : []);
          }
        }
      };
    }

    function addCommentToDocument(range) {
      
      const result = Array.isArray(range)
        ? range.map((item, index) => ({
            startIndex: item.globalOffset[0],
            endIndex: item.globalOffset[1],
            comment: item.category.level3,
            author: "AI批注",
            id: index + 1,
            desc:(item.suggestionList||[])[0]?.desc1||''
          }))
        : [];

      Asc.scope.targetRanges = result;

      clearAllComments(function () {
        me.callCommand(
          function () {

            var addedCount = 0;
            var targetRanges = Asc.scope.targetRanges

            try {
              var doc = Api.GetDocument();

              var sortedRanges = targetRanges.slice().sort(function (a, b) {
                return a.startIndex - b.startIndex;
              });

              // 过滤重叠区间
              function filterOverlappingRanges(ranges) {
                if (!ranges || ranges.length <= 1) {
                  return ranges;
                }

                var filteredRanges = [];
                var currentRange = ranges[0];

                for (var i = 1; i < ranges.length; i++) {
                  var nextRange = ranges[i];
                  
                  if (currentRange.endIndex >= nextRange.startIndex) {
                    if (currentRange.endIndex - currentRange.startIndex >= nextRange.endIndex - nextRange.startIndex) {
                      continue;
                    } else {
                      currentRange = nextRange;
                    }
                  } else {
                    filteredRanges.push(currentRange);
                    currentRange = nextRange;
                  }
                }
                filteredRanges.push(currentRange);
                
                return filteredRanges;
              }

          
              sortedRanges = filterOverlappingRanges(sortedRanges);

              function addCommentByCharacterIndexBatch() {
                var doc = Api.GetDocument();
                var count = doc.GetElementsCount();
                var globalCharIndex = 0;
                var processedCount = 0;

                var overlappingComments = [];

                for (var i = 0; i < count; i++) {
                  const element = doc.GetElement(i);
                  if (!element || element.GetClassType() !== 'paragraph')
                    continue;

                  try {
                    const runsCount = element.GetElementsCount();

                    for (var j = 0; j < runsCount; j++) {
                      const run = element.GetElement(j);
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
                              paragraph: element,
                              paragraphIndex: i,
                            });
                          }
                        }

                        // console.log(JSON.stringify(text), '可见字符数:', visibleCharCount, '总长度:', text.length);

                        globalCharIndex += visibleCharCount;
                      }
                    }

                    console.log(globalCharIndex, 'globalCharIndex');
                  } catch (error) {
                    console.error('段落 ' + i + ' 处理失败:', error);
                  }
                }

                // 

                console.log(
                  '批量批注处理完成，处理了 ' + processedCount + ' 个批注',
                );

                const runOperationsGroupByIdObject =
                  mergeRunByCommentId(overlappingComments);

                Object.keys(runOperationsGroupByIdObject).forEach((id) => {
                  const operations = runOperationsGroupByIdObject[id];
                  processRun(operations);
                });

                 function hideEditorLoading() {
                    const container = parent.parent.parent.window.parent.window.parent.window[0].document;
                    const loader = container.getElementById('editorLoader');
                    if (loader) {
                      const titleElement = loader.querySelector('h3');
                      const descElement = loader.querySelector('p');
                      if (titleElement) {
                        titleElement.textContent = '校对完成';
                      }
                      if (descElement) {
                        descElement.textContent = 'AI智能校对已完成，批注已添加到文档中';
                      }
                      
                      // 隐藏旋转动画
                      const spinner = loader.querySelector('div[style*="border-radius: 50%"]');
                      if (spinner) {
                        spinner.style.display = 'none';
                      }
                      
                      setTimeout(function() {
                        loader.style.opacity = '0';
                        loader.style.visibility = 'hidden';
                      }, 2000);
                    }
                  }

                hideEditorLoading()
                return processedCount;

         
              }

              /**
               * 根据comment id 合并run
               * @param  overlappingComments
               * @returns
               */
              function mergeRunByCommentId(overlappingComments) {
                const runOperationsGroupByIdObject = {};

                overlappingComments.forEach((comment) => {
                  if (!runOperationsGroupByIdObject[comment.id]) {
                    runOperationsGroupByIdObject[comment.id] = [];
                  }
                  runOperationsGroupByIdObject[comment.id].push(comment);

                  runOperationsGroupByIdObject[comment.id].sort((a, b) => {
                    return a.globalStartIndex - b.globalStartIndex;
                  });
                });

                return runOperationsGroupByIdObject;
              }

              function processRun(operations) {
                if (!operations || operations.length === 0) return;
                mergeAdjacentRuns(operations);
              }

              /**
               * 合并相邻的run
               * @param operations
               * @returns
               */
              function mergeAdjacentRuns(operations) {
                if (!operations || operations.length === 0) return;

                if (operations.length === 1) {
                  processSingleRun(operations[0]);
                } else {
                  processMultipleRuns(operations);
                }
              }

              /**
               * 处理单个run
               * @param operation
               * @returns
               */
              function processSingleRun(operation) {
                const run = operation.run;
                const text = operation.text;
                const startInRun = operation.startInRun;
                const endInRun = operation.endInRun;
                const paragraph = operation.paragraph;

                // 如果批注覆盖整个run，直接添加批注
                if (startInRun === 0 && endInRun === text.length) {
                  run.AddComment(
                      `${operation.originalRange.desc}`,
                    operation.author,
                  );
                  return;
                }

                const runFormat = getRunFormat(run);
                const beforeText = text.substring(0, startInRun);
                const commentText = text.substring(startInRun, endInRun);
                const afterText = text.substring(endInRun);

                run.ClearContent();

                if (beforeText) {
                  run.AddText(beforeText);
                }

                // 批注run
                const commentRun = Api.CreateRun();
                commentRun.AddText(commentText);
                applyRunFormat(commentRun, runFormat);

                const runIndex = operation.runIndex;
                paragraph.AddElement(
                  commentRun,
                  runIndex + (beforeText ? 1 : 0),
                );

                if (afterText) {
                  const afterRun = Api.CreateRun();
                  afterRun.AddText(afterText);
                  applyRunFormat(afterRun, runFormat);
                  paragraph.AddElement(
                    afterRun,
                    runIndex + (beforeText ? 2 : 1),
                  );
                }

                commentRun.AddComment(
                  `${operation.originalRange.desc}`,
                  operation.author,
                );
              }

              /**
               * 处理一个range 包含多个run的情况
               * @param operations
               * @returns
               */
              function processMultipleRuns(operations) {
                if (isRunsCrossParagraph(operations)) {
                  console.error('range 跨段落， 请检查是否存在误报');
                  return;
                }

                const firstOp = operations[0];
                const lastOp = operations[operations.length - 1];
                const firstRun = firstOp.run;

                const firstRunFormat = getRunFormat(firstRun);
                const lastRunFormat = getRunFormat(lastOp.run);

                const beforeText = firstOp.text.substring(
                  0,
                  firstOp.startInRun,
                );

                let commentText = firstOp.text.substring(firstOp.startInRun);

                for (let i = 1; i < operations.length - 1; i++) {
                  commentText += operations[i].text;
                }

                if (operations.length > 1) {
                  commentText += lastOp.text.substring(0, lastOp.endInRun);
                }

                const afterText = lastOp.text.substring(lastOp.endInRun);

                firstRun.ClearContent();

                for (let i = 1; i < operations.length; i++) {
                  const op = operations[i];
                  op.run.ClearContent();
                }

                if (beforeText) {
                  firstRun.AddText(beforeText);
                }

                const commentRun = Api.CreateRun();
                commentRun.AddText(commentText);
                applyRunFormat(commentRun, firstRunFormat);

                // commentRun.SetHighlight('yellow');

                const paragraph = firstOp.paragraph;
                const runIndex = firstOp.runIndex;

                paragraph.AddElement(
                  commentRun,
                  runIndex + (beforeText ? 1 : 0),
                );

                if (afterText) {
                  const afterRun = Api.CreateRun();
                  afterRun.AddText(afterText);
                  applyRunFormat(afterRun, lastRunFormat);
                  paragraph.AddElement(
                    afterRun,
                    runIndex + (beforeText ? 2 : 1),
                  );
                }

                commentRun.AddComment(
                  `${firstOp.originalRange.desc}`,
                  firstOp.author,
                );
              }

              /**
               * 检查operations中的run是否跨段落
               * @param  operations
               */
              function isRunsCrossParagraph(operations) {
                if (!operations || operations.length <= 1) {
                  return false;
                }

                const firstParagraphIndex = operations[0].paragraphIndex;

                for (let i = 1; i < operations.length; i++) {
                  if (operations[i].paragraphIndex !== firstParagraphIndex) {
                    return true;
                  }
                }

                return false;
              }

              /**
               * 获取run的格式信息
               * @param  run
               * @returns 格式对象
               */
              function getRunFormat(run) {
                const format = {};

                try {
                  format.bold = run.GetBold ? run.GetBold() : false;
                  format.italic = run.GetItalic ? run.GetItalic() : false;
                  format.underline = run.GetUnderline
                    ? run.GetUnderline()
                    : false;
                  format.strikeout = run.GetStrikeout
                    ? run.GetStrikeout()
                    : false;
                  format.fontSize = run.GetFontSize ? run.GetFontSize() : null;
                  format.fontFamily = run.GetFontFamily
                    ? run.GetFontFamily()
                    : null;
                  // format.color = run.GetColor ? run.GetColor() : null;
                  format.vertAlign = run.GetVertAlign
                    ? run.GetVertAlign()
                    : null;
                  format.spacing = run.GetSpacing ? run.GetSpacing() : null;
                  format.caps = run.GetCaps ? run.GetCaps() : null;
                  format.smallCaps = run.GetSmallCaps
                    ? run.GetSmallCaps()
                    : null;
                } catch (e) {
                  console.log('获取格式失败:', e);
                }

                return format;
              }

              /**
               * 应用格式到run
               * @param  run
               * @param  format
               */
              function applyRunFormat(run, format) {
                try {
                  if (format.bold !== undefined && run.SetBold) {
                    run.SetBold(format.bold);
                  }
                  if (format.italic !== undefined && run.SetItalic) {
                    run.SetItalic(format.italic);
                  }
                  if (format.underline !== undefined && run.SetUnderline) {
                    run.SetUnderline(format.underline);
                  }
                  if (format.strikeout !== undefined && run.SetStrikeout) {
                    run.SetStrikeout(format.strikeout);
                  }
                  if (format.fontSize && run.SetFontSize) {
                    run.SetFontSize(format.fontSize);
                  }
                  if (format.fontFamily && run.SetFontFamily) {
                    run.SetFontFamily(format.fontFamily);
                  }
                  if (format.color && run.SetColor) {
                    run.SetColor(format.color);
                  }
                  if (format.vertAlign && run.SetVertAlign) {
                    run.SetVertAlign(format.vertAlign);
                  }
                  if (format.spacing !== undefined && run.SetSpacing) {
                    run.SetSpacing(format.spacing);
                  }
                  if (format.caps !== undefined && run.SetCaps) {
                    run.SetCaps(format.caps);
                  }
                  if (format.smallCaps !== undefined && run.SetSmallCaps) {
                    run.SetSmallCaps(format.smallCaps);
                  }
                } catch (e) {
                  console.log('应用格式失败:', e);
                }
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
              // {
              //   id: 'addComment',
              //   type: 'button',
              //   text: 'AI校对',
              //   hint: '对文档进行AI智能校对并添加批注',
              //   icons: 'icon.png',
              //   lockInViewMode: true,
              //   enableToggle: false,
              //   separator: false,
              // },
              // {
              //   id: 'getDocumentText',
              //   type: 'button',
              //   text: '获取文档文本',
              //   hint: '获取整个文档的文本内容',
              //   icons: 'icon.png',
              //   lockInViewMode: true,
              //   enableToggle: false,
              //   separator: false,
              // },
              // {
              //   id: 'clearAllComments',
              //   type: 'button',
              //   text: '清除所有批注',
              //   hint: '清除文档中的所有批注',
              //   icons: 'icon.png',
              //   lockInViewMode: true,
              //   enableToggle: false,
              //   separator: false,
              // },
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
      addCommentToDocument();
    });

    // 处理获取文档文本按钮点击事件
    this.attachToolbarMenuClickEvent('getDocumentText', function (data) {
      getDocumentText(function (text) {
        if (text) {
          console.log('文档文本长度:', text.replace(/[\r\n]+/g, '').length);
          // console.log('文档文本:', text.replace(/[\r\n]+/g, ''));
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
      getDocumentText(function (documentText) {
        if (documentText) {
          callCheckAPI(documentText, function (error, response) {
            if (error) {
              console.error('API校对失败:', error);
              alert('API校对失败: ' + error);
            } else {
              console.log('API校对结果:', response);
              
              var tid = '20250827194717_c23fc8c35e5249b9ae34c2769ebbff0f';
              var sseConnection = callSSEAPI(tid, function (sseError, sseResponse, isRealtime) {
                if (sseError) {
                  console.log('SSE请求失败:', sseError);
                  if (!isRealtime) {
                    alert('SSE请求失败: ' + sseError);
                  }
                } else {
                  if (isRealtime) {
                    console.log('收到实时SSE数据:', sseResponse);
                  } else {

                    const range = JSON.parse(sseResponse[0].data.result.editing_check_result)
                  

                    addCommentToDocument(range)

                    // 调用添加
                    // alert('API校对和SSE请求都已完成，请查看控制台日志获取详细结果');
                  }
                }
              });
            }
          });
        } else {
          alert('获取文档文本失败，无法进行API校对');
        }
      });
    });


    // 添加批注
    function addCommentAction() {
      addCommentToDocument();
    }

    // 开始添加批注
    $('#addText').click(addCommentAction);

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
