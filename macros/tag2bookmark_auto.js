/**
 * 标签转书签自动化
 * 依次执行：1. 使占位符唯一化 2. 转书签 3. 删除标签
 * 处理范围：正文、页眉、页脚
 */
function 标签转书签自动化() {
    var doc = Application.ActiveDocument;
    Application.ScreenUpdating = false;
    
    var log = [];
    var summary = {
        uniqueCount: 0,
        bookmarkSuccess: 0,
        bookmarkTotal: 0,
        deletedCount: 0
    };
    
    var startTime = new Date().getTime();
    
    try {
        log.push("===== 完整处理流程开始 =====");
        log.push("文档名称: " + doc.Name);
        log.push("处理时间: " + new Date().toLocaleString());
        log.push("节数量: " + doc.Sections.Count);
        
        // ===== 1. 使占位符唯一化 =====
        log.push("\n===== 步骤1: 使占位符唯一化 =====");
        summary.uniqueCount = processUniquePlaceholders(doc, log);
        
        // ===== 2. 转书签 =====
        log.push("\n===== 步骤2: 转书签 =====");
        var bookmarkResult = processBookmarks(doc, log);
        summary.bookmarkSuccess = bookmarkResult.successCount;
        summary.bookmarkTotal = bookmarkResult.totalCount;
        
        // ===== 3. 删除标签 =====
        log.push("\n===== 步骤3: 删除标签 =====");
        summary.deletedCount = processDeleteTags(doc, log);
        
        // ===== 汇总结果 =====
        var endTime = new Date().getTime();
        var duration = Math.round((endTime - startTime) / 1000);
        
        log.push("\n===== 处理完成 =====");
        log.push("总耗时: " + duration + " 秒");
        log.push("===== 最终统计 =====");
        log.push("• 唯一化占位符数量: " + summary.uniqueCount + " 个");
        log.push("• 书签转换成功: " + summary.bookmarkSuccess + "/" + summary.bookmarkTotal + " 个");
        log.push("• 已删除标签数量: " + summary.deletedCount + " 个");
        
        // 输出日志
        outputToDebugWindow(log);
        
        // 显示汇总结果
        alert("✅ 完整处理流程完成!\n\n" +
              "处理统计:\n" +
              "• 唯一化占位符: " + summary.uniqueCount + " 个\n" +
              "• 书签转换成功: " + summary.bookmarkSuccess + "/" + summary.bookmarkTotal + " 个\n" +
              "• 已删除标签: " + summary.deletedCount + " 个\n\n" +
              "总耗时: " + duration + " 秒\n\n" +
              "详细日志请在 VBA 编辑器中按 Ctrl+G 查看");
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("❌ 处理过程中发生错误: " + errorMsg);
        outputToDebugWindow(log);
        alert("❌ 处理过程中发生错误:\n" + errorMsg + "\n\n详细信息请查看VBA调试窗口(按Ctrl+G)");
    } finally {
        Application.ScreenUpdating = true;
    }
}

// =============== 辅助函数 ===============

/**
 * 步骤1: 使占位符唯一化
 */
function processUniquePlaceholders(doc, log) {
    var totalReplaced = 0;
    var nameCounters = {}; // 用于跟踪每个占位符名称的出现次数
    
    try {
        log.push("[处理开始] 开始唯一化处理，初始化计数器");
        
        // ===== 正文处理 =====
        log.push("\n[正文] 开始处理正文中的占位符");
        var bodyResult = processRangeForUnique(doc.Content, log, "正文", nameCounters);
        totalReplaced += bodyResult.replacedCount;
        log.push("[正文] 处理完成，唯一化数量: " + bodyResult.replacedCount);
        
        // ===== 页眉/页脚处理 =====
        var headerFooterCount = 0;
        for (var i = 1; i <= doc.Sections.Count; i++) {
            var sec = doc.Sections(i);
            
            // 页眉
            for (var h = 1; h <= 3; h++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Headers(h).Exists) {
                        log.push("\n[页眉-节" + i + "-类型" + h + "] 开始处理");
                        var headerResult = processRangeForUnique(
                            sec.Headers(h).Range,
                            log,
                            "页眉-节" + i + "-类型" + h,
                            nameCounters
                        );
                        totalReplaced += headerResult.replacedCount;
                        headerFooterCount += headerResult.placeholderCount;
                        log.push("[页眉-节" + i + "-类型" + h + "] 处理完成，唯一化数量: " + headerResult.replacedCount);
                    }
                } catch (e) {
                    log.push("[页眉-节" + i + "-类型" + h + "] 处理出错: " + (e.description || e.message || String(e)));
                }
            }
            
            // 页脚
            for (var f = 1; f <= 3; f++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Footers(f).Exists) {
                        log.push("\n[页脚-节" + i + "-类型" + f + "] 开始处理");
                        var footerResult = processRangeForUnique(
                            sec.Footers(f).Range,
                            log,
                            "页脚-节" + i + "-类型" + f,
                            nameCounters
                        );
                        totalReplaced += footerResult.replacedCount;
                        headerFooterCount += footerResult.placeholderCount;
                        log.push("[页脚-节" + i + "-类型" + f + "] 处理完成，唯一化数量: " + footerResult.replacedCount);
                    }
                } catch (e) {
                    log.push("[页脚-节" + i + "-类型" + f + "] 处理出错: " + (e.description || e.message || String(e)));
                }
            }
        }
        
        log.push("\n[处理总结] 唯一化处理完成");
        log.push("[处理总结] 总唯一化数量: " + totalReplaced);
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("❌ 唯一化处理出错: " + errorMsg);
    }
    
    return totalReplaced;
}

/**
 * 处理指定范围的占位符唯一化 - 修复版
 */
function processRangeForUnique(rng, log, sectionName, nameCounters) {
    var replacedCount = 0;
    var placeholderCount = 0;
    
    try {
        // 直接在范围内查找并处理，避免范围失效问题
        var findRange = rng.Duplicate;
        findRange.Find.ClearFormatting();
        findRange.Find.Text = "\\$\\{[!\\}]{1,}\\}"; // 更精确的通配符：${后跟至少一个非}字符，后跟}
        findRange.Find.MatchWildcards = true;
        findRange.Find.Wrap = 0; // wdFindStop
        
        while (findRange.Find.Execute()) {
            var originalText = findRange.Text;
            var matchResult = originalText.match(/\$\{([^\}]+)\}/);
            
            if (matchResult) {
                placeholderCount++;
                var baseName = matchResult[1].trim();
                var currentText = originalText;
                
                // 初始化计数器
                if (!nameCounters.hasOwnProperty(baseName)) {
                    nameCounters[baseName] = 0;
                    log.push("[" + sectionName + "] 首次遇到占位符: " + baseName);
                }
                
                // 如果不是第一次出现，需要重命名
                if (nameCounters[baseName] > 0) {
                    var newName = baseName + "_" + nameCounters[baseName];
                    var newText = "${" + newName + "}";
                    
                    // 应用更改
                    findRange.Text = newText;
                    replacedCount++;
                    
                    log.push("[" + sectionName + "] 唯一化 #" + nameCounters[baseName] + ": " + 
                             originalText + " → " + newText);
                    
                    // 更新当前文本用于后续处理
                    currentText = newText;
                }
                
                // 重要：增加计数器
                nameCounters[baseName]++;
                
                // 重要：在替换后，将范围折叠到末尾继续搜索
                findRange.Collapse(0); // wdCollapseEnd
                
            } else {
                // 没有匹配，继续前进
                log.push("[" + sectionName + "] 警告: 未匹配到占位符格式: " + findRange.Text);
                findRange.Collapse(0);
            }
        }
        
        log.push("[" + sectionName + "] 扫描完成: 共找到 " + placeholderCount + " 个占位符，其中 " + replacedCount + " 个被唯一化");
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("[" + sectionName + "] 唯一化处理出错: " + errorMsg);
    }
    
    return {
        replacedCount: replacedCount,
        placeholderCount: placeholderCount
    };
}

/**
 * 步骤2: 转书签
 */
function processBookmarks(doc, log) {
    var successCount = 0;
    var totalCount = 0;
    var bookmarkCounts = {};
    
    try {
        // ===== 正文处理 =====
        var result = processRangeForBookmarks(doc, doc.Content, "正文", bookmarkCounts, log);
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
                            doc,
                            sec.Headers(h).Range, 
                            "页眉-节" + i + "-类型" + h, 
                            bookmarkCounts, 
                            log
                        );
                        successCount += headerResult.successCount;
                        totalCount += headerResult.totalCount;
                    }
                } catch (e) {
                    log.push("[页眉-节" + i + "-类型" + h + "] 书签处理出错: " + (e.description || e.message || String(e)));
                }
            }
            
            // 页脚
            for (var f = 1; f <= 3; f++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Footers(f).Exists) {
                        var footerResult = processRangeForBookmarks(
                            doc,
                            sec.Footers(f).Range, 
                            "页脚-节" + i + "-类型" + f, 
                            bookmarkCounts, 
                            log
                        );
                        successCount += footerResult.successCount;
                        totalCount += footerResult.totalCount;
                    }
                } catch (e) {
                    log.push("[页脚-节" + i + "-类型" + f + "] 书签处理出错: " + (e.description || e.message || String(e)));
                }
            }
        }
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("❌ 书签转换出错: " + errorMsg);
    }
    
    return {
        successCount: successCount,
        totalCount: totalCount
    };
}

/**
 * 处理指定范围内的占位符，为它们添加书签
 */
function processRangeForBookmarks(doc, rng, sectionName, bookmarkCounts, log) {
    var successCount = 0;
    var totalCount = 0;
    
    try {
        // 获取所有占位符
        var placeholders = [];
        var scanRange = rng.Duplicate;
        
        scanRange.Find.ClearFormatting();
        scanRange.Find.Text = "\\$\\{[!\\}]{1,}\\}";
        scanRange.Find.MatchWildcards = true;
        scanRange.Find.Wrap = 0; // wdFindStop
        
        while (scanRange.Find.Execute()) {
            var matchResult = scanRange.Text.match(/\$\{([^\}]+)\}/);
            if (matchResult) {
                placeholders.push({
                    text: scanRange.Text,
                    name: matchResult[1],
                    range: scanRange.Duplicate
                });
            }
            scanRange.Collapse(0); // wdCollapseEnd
        }
        
        // 为每个占位符创建书签
        for (var i = 0; i < placeholders.length; i++) {
            var placeholder = placeholders[i];
            var rawName = placeholder.name;
            var cleanName = cleanBookmarkName(rawName);
            
            totalCount++;
            
            // 生成书签名
            var bookmarkName = cleanName;
            var counter = 0;
            
            // 检查书签名是否已存在，如果存在则添加后缀
            while (doc.Bookmarks.Exists(bookmarkName)) {
                counter++;
                bookmarkName = cleanName + "_" + counter;
            }
            
            // 记录使用的书签名
            if (!bookmarkCounts.hasOwnProperty(cleanName)) {
                bookmarkCounts[cleanName] = 0;
            }
            
            try {
                // 在占位符起始位置创建书签
                var bookmarkRange = placeholder.range.Duplicate;
                bookmarkRange.Collapse(1); // 1=起始位置
                
                // 添加书签
                doc.Bookmarks.Add(bookmarkName, bookmarkRange);
                successCount++;
                
                log.push("[" + sectionName + "] 书签: " + 
                         bookmarkName + " ← " + placeholder.text);
                
                // 更新计数器
                bookmarkCounts[cleanName]++;
                
            } catch (e) {
                var errorMsg = e.description || e.message || String(e);
                log.push("[" + sectionName + "] 书签创建失败: " + 
                         bookmarkName + " - " + errorMsg);
            }
        }
        
        log.push("[" + sectionName + "] 书签处理完成: " + successCount + "/" + totalCount + " 个成功");
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("[" + sectionName + "] 书签处理出错: " + errorMsg);
    }
    
    return {
        successCount: successCount,
        totalCount: totalCount
    };
}

/**
 * 步骤3: 删除标签
 */
function processDeleteTags(doc, log) {
    var totalCount = 0;
    
    try {
        // ===== 正文处理 =====
        log.push("\n[正文] 开始删除标签");
        totalCount += processRangeForDeletion(doc.Content, "正文", log);
        
        // ===== 页眉/页脚处理 =====
        for (var i = 1; i <= doc.Sections.Count; i++) {
            var sec = doc.Sections(i);
            
            // 页眉
            for (var h = 1; h <= 3; h++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Headers(h).Exists) {
                        log.push("\n[页眉-节" + i + "-类型" + h + "] 开始删除标签");
                        totalCount += processRangeForDeletion(
                            sec.Headers(h).Range, 
                            "页眉-节" + i + "-类型" + h, 
                            log
                        );
                    }
                } catch (e) {
                    log.push("[页眉-节" + i + "-类型" + h + "] 删除处理出错: " + (e.description || e.message || String(e)));
                }
            }
            
            // 页脚
            for (var f = 1; f <= 3; f++) { // 1=首页, 2=偶数页, 3=奇数页
                try {
                    if (sec.Footers(f).Exists) {
                        log.push("\n[页脚-节" + i + "-类型" + f + "] 开始删除标签");
                        totalCount += processRangeForDeletion(
                            sec.Footers(f).Range, 
                            "页脚-节" + i + "-类型" + f, 
                            log
                        );
                    }
                } catch (e) {
                    log.push("[页脚-节" + i + "-类型" + f + "] 删除处理出错: " + (e.description || e.message || String(e)));
                }
            }
        }
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("❌ 标签删除出错: " + errorMsg);
    }
    
    return totalCount;
}

/**
 * 处理指定范围内的标签删除
 */
function processRangeForDeletion(rng, sectionName, log) {
    var count = 0;
    
    try {
        // 配置查找
        var findRange = rng.Duplicate;
        findRange.Find.ClearFormatting();
        findRange.Find.Text = "\\$\\{[!\\}]{1,}\\}"; // 匹配 ${...} 格式
        findRange.Find.MatchWildcards = true;
        findRange.Find.Wrap = 0; // wdFindStop
        
        // 执行查找并删除
        while (findRange.Find.Execute()) {
            var tagText = findRange.Text;
            log.push("[" + sectionName + "] 删除: " + tagText);
            
            // 删除标签内容
            findRange.Text = "";
            count++;
            
            // 继续查找
            findRange.Collapse(0); // wdCollapseEnd
        }
        
        log.push("[" + sectionName + "] 删除完成: 共删除 " + count + " 个标签");
        
    } catch (e) {
        var errorMsg = e.description || e.message || String(e);
        log.push("[" + sectionName + "] 删除处理出错: " + errorMsg);
    }
    
    return count;
}

/**
 * 将日志输出到 VBA 调试窗口
 */
function outputToDebugWindow(log) {
    try {
        var debugLog = "===== 完整处理流程日志 =====\n";
        for (var i = 0; i < log.length; i++) {
            debugLog += log[i] + "\n";
        }
        
        if (typeof Debug !== "undefined" && Debug.Print) {
            Debug.Print(debugLog);
        }
    } catch (e) {
        // 忽略错误
    }
}

/**
 * 清理书签名 - 增强版
 */
function cleanBookmarkName(s) {
    if (!s || s.trim() === "") {
        return "Placeholder";
    }
    
    // 允许字母、数字、下划线和中文，其他字符替换为下划线
    var cleaned = s.replace(/[^\w\u4e00-\u9fa5]/g, "_");
    
    // 移除Word禁止的特殊字符
    cleaned = cleaned.replace(/[\[\]\<\>\\:;\*\?\"\/\|]/g, "_");
    
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