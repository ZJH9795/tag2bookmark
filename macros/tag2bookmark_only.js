/**
 * 转书签 - 优化版 (支持正文、页眉、页脚)
 * 在所有 ${xxx} 占位符后添加书签
 */
function 转书签() {
    var doc = Application.ActiveDocument;
    Application.ScreenUpdating = false;
    
    var log = [];
    var successCount = 0;
    var totalCount = 0;
    var bookmarkCounts = {};
    
    try {
        log.push("===== 开始处理书签转换 =====");
        log.push("文档名称: " + doc.Name);
        log.push("节数量: " + doc.Sections.Count);
        
        // ===== 正文处理 =====
        var result = processRangeForBookmarks(doc.Content, "正文", bookmarkCounts, log);
        successCount += result.successCount;
        totalCount += result.totalCount;
        
        // ===== 页眉/页脚处理 =====
        for (var i = 1; i <= doc.Sections.Count; i++) {
            var sec = doc.Sections(i);
            
            // 页眉
            for (var h = 1; h <= 3; h++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Headers(h).Exists) {
                        var headerResult = processRangeForBookmarks(
                            sec.Headers(h).Range, 
                            "页眉-节" + i + "-类型" + h, 
                            bookmarkCounts, 
                            log
                        );
                        successCount += headerResult.successCount;
                        totalCount += headerResult.totalCount;
                    }
                } catch (e) {
                    // 忽略错误，某些页眉可能不存在
                }
            }
            
            // 页脚
            for (var f = 1; f <= 3; f++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Footers(f).Exists) {
                        var footerResult = processRangeForBookmarks(
                            sec.Footers(f).Range, 
                            "页脚-节" + i + "-类型" + f, 
                            bookmarkCounts, 
                            log
                        );
                        successCount += footerResult.successCount;
                        totalCount += footerResult.totalCount;
                    }
                } catch (e) {
                    // 忽略错误，某些页脚可能不存在
                }
            }
        }
        
        log.push("===== 处理完成 =====");
        log.push("总占位符数量: " + totalCount);
        log.push("成功添加书签: " + successCount);
        
        // 输出日志
        outputToDebugWindow(log);
        
        alert("✅ 书签转换完成!\n" + 
              "共处理 " + totalCount + " 个占位符\n" +
              "成功添加 " + successCount + " 个书签\n\n" +
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
 * 处理指定范围内的占位符，为它们添加书签
 * @param {Range} rng 要处理的范围
 * @param {string} sectionName 区域名称(用于日志)
 * @param {Object} bookmarkCounts 书签计数器(引用传递)
 * @param {Array} log 日志数组
 * @returns {Object} {successCount, totalCount}
 */
function processRangeForBookmarks(rng, sectionName, bookmarkCounts, log) {
    var successCount = 0;
    var totalCount = 0;
    
    try {
        log.push("处理区域: " + sectionName);
        
        // 先扫描获取所有唯一占位符名称
        var placeholderNames = {};
        var scanRange = rng.Duplicate;
        
        scanRange.Find.ClearFormatting();
        scanRange.Find.Text = "\\$\\{*\\}";
        scanRange.Find.MatchWildcards = true;
        scanRange.Find.Wrap = 0; // wdFindStop
        
        while (scanRange.Find.Execute()) {
            var matchResult = scanRange.Text.match(/\$\{([^\}]+)\}/);
            if (matchResult) {
                var rawName = matchResult[1];
                placeholderNames[rawName] = true;
            }
            scanRange.Collapse(0); // wdCollapseEnd
        }
        
        // 为每种占位符名称创建书签
        for (var rawName in placeholderNames) {
            var cleanName = cleanBookmarkName(rawName);
            var counter = 0;
            var searchRange = rng.Duplicate;
            
            searchRange.Find.ClearFormatting();
            searchRange.Find.Text = "${" + rawName + "}";
            searchRange.Find.MatchCase = true;
            searchRange.Find.Wrap = 0; // wdFindStop
            
            while (searchRange.Find.Execute()) {
                totalCount++;
                
                // 生成书签名
                var bookmarkName = cleanName;
                if (bookmarkCounts.hasOwnProperty(cleanName)) {
                    if (counter > 0) {
                        bookmarkName = cleanName + "_" + counter;
                    }
                } else {
                    bookmarkCounts[cleanName] = 0;
                }
                
                try {
                    // 复制范围并移动到结束位置(在占位符后添加书签)
                    var bookmarkRange = searchRange.Duplicate;
                    bookmarkRange.Collapse(0); // wdCollapseEnd - 折叠到结束位置
                    
                    // 如果书签已存在则先删除
                    if (Application.ActiveDocument.Bookmarks.Exists(bookmarkName)) {
                        Application.ActiveDocument.Bookmarks(bookmarkName).Delete();
                    }
                    
                    // 添加书签
                    Application.ActiveDocument.Bookmarks.Add(bookmarkName, bookmarkRange);
                    successCount++;
                    
                    log.push("[" + sectionName + "] 添加书签: " + bookmarkName + 
                             " (原始占位符: ${" + rawName + "})");
                    
                    // 更新计数器
                    if (bookmarkCounts.hasOwnProperty(cleanName)) {
                        bookmarkCounts[cleanName]++;
                    }
                    counter++;
                    
                } catch (e) {
                    var errorMsg = e.description || e.message || String(e);
                    log.push("[" + sectionName + "] 书签创建失败: " + bookmarkName + 
                             " - 错误: " + errorMsg);
                }
                
                // 重要: 折叠范围到末尾，继续查找
                searchRange.Collapse(0); // wdCollapseEnd
            }
        }
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("[" + sectionName + "] 处理区域时出错: " + errorMsg);
    }
    
    return {
        successCount: successCount,
        totalCount: totalCount
    };
}

/**
 * 将日志输出到 VBA 调试窗口
 */
function outputToDebugWindow(log) {
    try {
        if (typeof Debug !== "undefined" && Debug.Print) {
            Debug.Print("===== 书签转换日志开始 =====");
            for (var i = 0; i < log.length; i++) {
                Debug.Print(log[i]);
            }
            Debug.Print("===== 书签转换日志结束 =====");
        }
    } catch (e) {
        // 忽略错误，某些环境可能不支持Debug
    }
}

/**
 * 辅助函数：清理书签名
 * 确保书签名符合Word要求
 */
function cleanBookmarkName(s) {
    if (!s || s.trim() === "") {
        return "Placeholder";
    }
    
    // 允许字母、数字、下划线和中文，其他字符替换为下划线
    var cleaned = s.replace(/[^\w\u4e00-\u9fa5]/g, "_");
    
    // 不能以数字开头
    if (/^[0-9]/.test(cleaned)) {
        cleaned = "B_" + cleaned;
    }
    
    // 合并多个下划线
    cleaned = cleaned.replace(/_+/g, "_");
    // 移除首尾下划线
    cleaned = cleaned.replace(/^_+|_+$/g, "");
    
    // 确保不为空
    if (cleaned === "") {
        cleaned = "Placeholder";
    }
    
    // Word书签名最大长度为40个字符
    if (cleaned.length > 40) {
        cleaned = cleaned.substring(0, 40);
    }
    
    return cleaned;
}