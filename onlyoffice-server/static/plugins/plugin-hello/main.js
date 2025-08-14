/* eslint-disable */

(function(window, undefined) {
  window.Asc.plugin.init = function(initData) {
    console.log('插件开始初始化')
    console.log(window)

    var me = this
    $('#addText').click(function() {
      me.callCommand(function() {
        try {
          // 获取文档对象
          var oDocument = Api.GetDocument()

          // 示例：要高亮的文本范围（你可以根据实际需求修改这些值）
          const targetRanges = [
            { startIndex: 0, endIndex: 20},  // 高亮第0-7个字符
            // 可以添加更多范围...
          ];

          // 简化版：直接在现有Run上分段高亮
          function highlightByCharacterIndexSimple(startIndex, endIndex) {
            let doc = Api.GetDocument();
            let count = doc.GetElementsCount();
            let globalCharIndex = 0;
            
            console.log(`开始简化高亮: 索引 ${startIndex} 到 ${endIndex}`);
            
            for (let i = 0; i < count; i++) {
              let element = doc.GetElement(i);
              if (!element || element.GetClassType() !== "paragraph") continue;
              
              try {
                let runsCount = element.GetElementsCount();
                
                for (let j = 0; j < runsCount; j++) {
                  let run = element.GetElement(j);
                  if (run && run.GetClassType() === "run") {
                    let text = run.GetText();
                    if (!text) continue;
                    
                    let runStartIndex = globalCharIndex;
                    let runEndIndex = globalCharIndex + text.length;
                    
                    console.log(`检查Run: "${text}" (索引 ${runStartIndex}-${runEndIndex-1})`);
                    
                                         // 检查这个Run是否与目标范围重叠
                     if (runStartIndex < endIndex && runEndIndex > startIndex) {
                       // 计算需要高亮的精确范围
                       let highlightStartInRun = Math.max(0, startIndex - runStartIndex);
                       let highlightEndInRun = Math.min(text.length, endIndex - runStartIndex);
                       
                       console.log(`Run内高亮范围: ${highlightStartInRun}-${highlightEndInRun} (总长度${text.length})`);
                       
                       // 如果整个Run都需要高亮
                       if (highlightStartInRun === 0 && highlightEndInRun === text.length) {
                         try {
                           run.SetHighlight("yellow");
                           console.log(`✅ 高亮整个Run: "${text}"`);
                         } catch (e) {
                           console.log(`❌ 整个Run高亮失败: "${text}"`, e);
                         }
                       } 
                                              // 如果只需要部分高亮，在段落级别拆分Run
                        else {
                          try {
                            let beforeText = text.substring(0, highlightStartInRun);
                            let highlightText = text.substring(highlightStartInRun, highlightEndInRun);
                            let afterText = text.substring(highlightEndInRun);
                            
                            console.log(`分段: 前"${beforeText}" + 高亮"${highlightText}" + 后"${afterText}"`);
                            
                                                        // 保存原始Run的所有格式属性
                             console.log("正在保存原始格式...");
                             let originalFormat = {};
                             try {
                               originalFormat.bold = run.GetBold ? run.GetBold() : false;
                               originalFormat.italic = run.GetItalic ? run.GetItalic() : false;
                               originalFormat.underline = run.GetUnderline ? run.GetUnderline() : false;
                               originalFormat.fontSize = run.GetFontSize ? run.GetFontSize() : null;
                               originalFormat.fontFamily = run.GetFontFamily ? run.GetFontFamily() : null;
                               
                               // 完全跳过颜色处理，避免undefined问题
                               originalFormat.color = null;
                               console.log("跳过颜色处理，使用默认颜色");
                               
                               console.log("原始格式:", originalFormat);
                             } catch (e) {
                               console.log("格式获取失败:", e);
                             }
                             
                             // 修改当前Run为前段文本
                             run.ClearContent();
                             if (beforeText) {
                               run.AddText(beforeText);
                             }
                             
                             // 创建高亮Run（完全继承原格式）
                             if (highlightText) {
                               let highlightRun = Api.CreateRun();
                               highlightRun.AddText(highlightText);
                               
                                                                // 先应用原始格式，再添加高亮背景
                                 try {
                                   if (originalFormat.bold) highlightRun.SetBold(originalFormat.bold);
                                   if (originalFormat.italic) highlightRun.SetItalic(originalFormat.italic);
                                   if (originalFormat.underline) highlightRun.SetUnderline(originalFormat.underline);
                                   if (originalFormat.fontSize) highlightRun.SetFontSize(originalFormat.fontSize);
                                   if (originalFormat.fontFamily) highlightRun.SetFontFamily(originalFormat.fontFamily);
                                   
                                   // 跳过颜色设置，使用默认黑色
                                   console.log("使用默认文字颜色");
                                   
                                   // 最后添加高亮背景（不影响其他格式）
                                   highlightRun.SetHighlight("yellow");
                                   console.log("✅ 高亮Run格式应用成功");
                                 } catch (e) {
                                   console.log("❌ 格式应用失败:", e);
                                 }
                               
                               // 在段落中插入高亮Run
                               element.AddElement(highlightRun, j + 1);
                             }
                             
                             // 创建后段Run（完全继承原格式）
                             if (afterText) {
                               let afterRun = Api.CreateRun();
                               afterRun.AddText(afterText);
                               
                                                                // 应用原始格式（不添加高亮）
                                 try {
                                   if (originalFormat.bold) afterRun.SetBold(originalFormat.bold);
                                   if (originalFormat.italic) afterRun.SetItalic(originalFormat.italic);
                                   if (originalFormat.underline) afterRun.SetUnderline(originalFormat.underline);
                                   if (originalFormat.fontSize) afterRun.SetFontSize(originalFormat.fontSize);
                                   if (originalFormat.fontFamily) afterRun.SetFontFamily(originalFormat.fontFamily);
                                   
                                   // 跳过颜色设置，使用默认黑色
                                   console.log("后段使用默认文字颜色");
                                   
                                   console.log("✅ 后段Run格式应用成功");
                                 } catch (e) {
                                   console.log("❌ 后段格式应用失败:", e);
                                 }
                               
                               // 在段落中插入后段Run
                               element.AddElement(afterRun, j + 2);
                             }
                            
                            console.log(`✅ 精确高亮成功: "${highlightText}"`);
                            
                            // 跳过新插入的Run，避免重复处理
                            if (highlightText) j++;
                            if (afterText) j++;
                            
                          } catch (e) {
                            console.log(`❌ 精确高亮失败，退回整体高亮: "${text}"`, e);
                            try {
                              run.SetHighlight("yellow");
                            } catch (e2) {
                              console.log(`❌ 退回方案也失败: "${text}"`, e2);
                            }
                          }
                        }
                     } else {
                       console.log(`跳过Run: "${text}" (不在目标范围)`);
                     }
                    
                    globalCharIndex += text.length;
                  }
                }
              } catch (error) {
                console.log(`❌ 段落 ${i} 处理失败:`, error);
              }
            }
            
            console.log(`简化高亮完成: ${startIndex}-${endIndex}`);
          }

          // 先做一个简单测试：测试各种高亮方法
          console.log("=== 开始API测试 ===");
          let testDoc = Api.GetDocument();
          let firstElement = testDoc.GetElement(0);
          if (firstElement && firstElement.GetClassType() === "paragraph") {
            let firstRun = firstElement.GetElement(0);
            if (firstRun && firstRun.GetClassType() === "run") {
              console.log("找到第一个Run，测试高亮方法");
              
              // 测试1：SetHighlight
              try {
                firstRun.SetHighlight("yellow");
                console.log("✅ SetHighlight 调用成功");
              } catch (e) {
                console.log("❌ SetHighlight 失败:", e);
              }
              
              // 测试2：SetHighlight 数字参数
              try {
                firstRun.SetHighlight(255, 255, 0);
                console.log("✅ SetHighlight RGB 调用成功");
              } catch (e) {
                console.log("❌ SetHighlight RGB 失败:", e);
              }
              
              // 测试3：查看Run的方法
              console.log("Run 可用方法:", Object.getOwnPropertyNames(firstRun));
              
              // 测试4：尝试其他格式方法
              try {
                firstRun.SetBold(true);
                console.log("✅ SetBold 成功");
              } catch (e) {
                console.log("❌ SetBold 失败:", e);
              }
            } else {
              console.log("❌ 没找到第一个Run");
            }
          } else {
            console.log("❌ 没找到第一个段落");
          }
          console.log("=== API测试结束 ===");

          // 执行简化版高亮
          targetRanges.forEach((range, index) => {
            console.log(`处理范围 ${index + 1}: ${range.startIndex} - ${range.endIndex}`);
            highlightByCharacterIndexSimple(range.startIndex, range.endIndex);
          });

        } catch (error) {
          console.error(error)
        }
      }, false, true, function () {
        console.log('操作成功')
      })
    })

    // 在插件 iframe 之外释放鼠标按钮时调用的函数
    window.Asc.plugin.onExternalMouseUp = function() {
      var event = document.createEvent('MouseEvents')
      event.initMouseEvent('mouseup', true, true, window, 1, 0, 0, 0, 0, false, false, false, false, 0, null)
      document.dispatchEvent(event)
    }

    window.Asc.plugin.button = function(id) {
      // 被中断或关闭窗口
      if (id === -1) {
        this.executeCommand('close', '')
      }
	  }
  }
})(window, undefined)
