/**
 * 删除标签 - 完整版
 * 删除文档中所有形如 ${XXXXX} 的内容，包括页眉页脚
 */
function 删除标签() {
    var doc = Application.ActiveDocument;
    Application.ScreenUpdating = false;
    
    var log = [];
    var totalCount = 0;
    
    try {
        log.push("===== 开始删除标签 =====");
        log.push("文档名称: " + doc.Name);
        log.push("节数量: " + doc.Sections.Count);
        
        // ===== 正文处理 =====
        var result = processRangeForDeletion(doc.Content, "正文", log);
        totalCount += result;
        
        // ===== 页眉/页脚处理 =====
        for (var i = 1; i <= doc.Sections.Count; i++) {
            var sec = doc.Sections(i);
            
            // 页眉
            for (var h = 1; h <= 3; h++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Headers(h).Exists) {
                        var headerResult = processRangeForDeletion(
                            sec.Headers(h).Range, 
                            "页眉-节" + i + "-类型" + h, 
                            log
                        );
                        totalCount += headerResult;
                    }
                } catch (e) {
                    // 忽略错误，某些页眉可能不存在
                }
            }
            
            // 页脚
            for (var f = 1; f <= 3; f++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Footers(f).Exists) {
                        var footerResult = processRangeForDeletion(
                            sec.Footers(f).Range, 
                            "页脚-节" + i + "-类型" + f, 
                            log
                        );
                        totalCount += footerResult;
                    }
                } catch (e) {
                    // 忽略错误，某些页脚可能不存在
                }
            }
        }
        
        log.push("===== 删除完成 =====");
        log.push("总共删除: " + totalCount + " 个标签");
        
        // 输出日志
        outputToDebugWindow(log);
        
        alert("✅ 标签删除完成!\n" + 
              "共删除 " + totalCount + " 个标签\n\n" +
              "详细日志请在 VBA 编辑器中按 Ctrl+G 查看");
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        alert("❌ 处理文档时出错: " + errorMsg);
        log.push("错误: " + errorMsg);
        outputToDebugWindow(log);
    } finally {
        Application.ScreenUpdating = true;
    }
}

/**
 * 处理指定范围内的标签删除
 * @param {Range} rng 要处理的范围
 * @param {string} sectionName 区域名称(用于日志)
 * @param {Array} log 日志数组
 * @returns {number} 删除的标签数量
 */
function processRangeForDeletion(rng, sectionName, log) {
    var count = 0;
    var originalRange = rng.Duplicate;
    
    try {
        log.push("处理区域: " + sectionName);
        
        // 配置查找
        rng.Find.ClearFormatting();
        rng.Find.Text = "\\$\\{*\\}"; // 通配符模式：${任何内容}
        rng.Find.MatchWildcards = true; // 启用通配符
        rng.Find.Wrap = 0; // wdFindStop - 不循环
        
        // 执行查找并删除
        while (rng.Find.Execute()) {
            var tagText = rng.Text;
            log.push("[" + sectionName + "] 删除标签: " + tagText);
            
            // 删除标签内容
            rng.Text = "";
            count++;
            
            // 重要：重新设置查找范围，避免无限循环
            rng = originalRange.Duplicate;
            rng.Find.ClearFormatting();
            rng.Find.Text = "\\$\\{*\\}";
            rng.Find.MatchWildcards = true;
            rng.Find.Wrap = 0;
        }
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("[" + sectionName + "] 处理区域时出错: " + errorMsg);
    }
    
    return count;
}

/**
 * 将日志输出到 VBA 调试窗口
 */
function outputToDebugWindow(log) {
    try {
        if (typeof Debug !== "undefined" && Debug.Print) {
            Debug.Print("===== 删除标签日志开始 =====");
            for (var i = 0; i < log.length; i++) {
                Debug.Print(log[i]);
            }
            Debug.Print("===== 删除标签日志结束 =====");
        }
    } catch (e) {
        // 忽略错误，某些环境可能不支持Debug
    }
}