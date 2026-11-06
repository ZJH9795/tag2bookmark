/**
 * WPS Word JSA 宏
 * 占位符唯一化（顺序严格按阅读顺序）
 *
 * 规则：
 *   第 1 次出现：${name}
 *   第 2 次出现：${name_1}
 *   第 3 次出现：${name_2}
 */

function 使占位符唯一化() {
    var doc = Application.ActiveDocument;
    Application.ScreenUpdating = false;

    var log = [];
    var totalReplaced = 0;

    try {
        log.push("===== 开始处理占位符唯一化 =====");
        log.push("文档名称: " + doc.Name);
        log.push("节数量: " + doc.Sections.Count);

        // ===== 正文 =====
        totalReplaced += processRange(doc.Content, log, "正文");

        // ===== 页眉 / 页脚 =====
        for (var i = 1; i <= doc.Sections.Count; i++) {
            var sec = doc.Sections(i);

            // 页眉
            for (var h = 1; h <= 3; h++) {
                try {
                    if (sec.Headers(h)) {
                        totalReplaced += processRange(
                            sec.Headers(h).Range,
                            log,
                            "页眉-节" + i + "-类型" + h
                        );
                    }
                } catch (e) {}
            }

            // 页脚
            for (var f = 1; f <= 3; f++) {
                try {
                    if (sec.Footers(f)) {
                        totalReplaced += processRange(
                            sec.Footers(f).Range,
                            log,
                            "页脚-节" + i + "-类型" + f
                        );
                    }
                } catch (e) {}
            }
        }

        log.push("===== 处理完成 =====");
        log.push("总替换次数: " + totalReplaced);

        outputToDebugWindow(log);

        alert(
            "✅ 占位符唯一化完成\n" +
            "共替换 " + totalReplaced + " 个重复占位符\n\n" +
            "详细日志请在 VBA 编辑器中按 Ctrl+G 查看"
        );

    } catch (e) {
        alert("❌ 执行出错: " + (e.message || e));
    } finally {
        Application.ScreenUpdating = true;
    }
}


/**
 * 对指定 Range 进行占位符唯一化
 * 顺序以 Word/WPS Find 的阅读顺序为准
 */
function processRange(rng, log, sectionName) {
    var replaced = 0;

    // === 第一步：扫描本范围内所有占位符名称 ===
    var names = {};
    var scan = rng.Duplicate;

    scan.Find.ClearFormatting();
    scan.Find.Text = "\\$\\{*\\}";
    scan.Find.MatchWildcards = true;
    scan.Find.Wrap = 0; // wdFindStop

    while (scan.Find.Execute()) {
        var m = scan.Text.match(/\$\{([^\}]+)\}/);
        if (m) {
            names[m[1]] = true;
        }
        scan.Collapse(0); // wdCollapseEnd
    }

    // === 第二步：对每一种占位符，按 Find 顺序编号 ===
    for (var name in names) {
        var counter = 0;
        var r = rng.Duplicate;

        r.Find.ClearFormatting();
        r.Find.Text = "${" + name + "}";
        r.Find.MatchCase = true;
        r.Find.Wrap = 0; // wdFindStop

        while (r.Find.Execute()) {
            if (counter > 0) {
                var newText = "${" + name + "_" + counter + "}";
                r.Text = newText;
                replaced++;

                log.push(
                    "[" + sectionName + "] " +
                    "${" + name + "} → " + newText
                );
            }
            counter++;
            r.Collapse(0); // wdCollapseEnd
        }
    }

    return replaced;
}


/**
 * 将日志输出到 VBA 调试窗口
 */
function outputToDebugWindow(log) {
    try {
        Debug.Print("===== 日志开始 =====");
        for (var i = 0; i < log.length; i++) {
            Debug.Print(log[i]);
        }
        Debug.Print("===== 日志结束 =====");
    } catch (e) {
        // 忽略
    }
}
