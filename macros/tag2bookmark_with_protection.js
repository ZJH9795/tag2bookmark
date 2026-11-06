/**
 * 标签转书签_自动加解密
 * 功能流程：
 * 1. 检查文档保护状态
 * 2. 如有保护，解除文档保护（密码由用户在运行时输入）
 * 3. 执行标签转书签自动化
 * 4. 只有当原始文档受保护时，才重新应用文档保护
 */
function 标签转书签_自动加解密() {
    var doc = Application.ActiveDocument;
    Application.ScreenUpdating = false;
    
    // WPS中保护类型常量 (实测值)
    var wdNoProtection_WPS = -1; // WPS中无保护文档返回-1
    var wdAllowOnlyFormFields = 3; // 窗体保护类型
    
    // ===== 密码由用户在运行时输入（不写死在代码里） =====
    var PASSWORD = "";
    try {
        var pt0 = doc.ProtectionType;
        if (pt0 !== -1 && pt0 !== 0) {
            PASSWORD = InputBox("文档受保护，请输入文档保护密码：", "需要密码");
            if (!PASSWORD) {
                alert("已取消：未输入文档保护密码");
                return;
            }
        }
    } catch (e0) { /* ProtectionType 不可读时忽略，后续流程会自行检测 */ }

    var log = [];
    var wasProtectedInitially = false;
    var protectionRemoved = false;
    var operationSuccess = true;
    var finalMessage = "";
    var startTime = new Date().getTime();
    
    // 临时覆盖alert函数
    var originalAlert = alert;
    alert = function() {};
    
    try {
        log.push("===== 解锁_转书签_加锁流程开始 =====");
        log.push("文档名称: " + doc.Name);
        log.push("处理时间: " + new Date().toLocaleString());
        
        // ===== 1. 检查文档初始保护状态 (WPS专用逻辑) =====
        log.push("\n===== 步骤1: 检查文档初始保护状态 (WPS专用) =====");
        
        try {
            var currentProtectionType = doc.ProtectionType;
            log.push("[保护] 检测到ProtectionType值: " + currentProtectionType);
            
            // WPS专用逻辑: 无保护文档返回-1，受保护返回其他正值
            if (currentProtectionType === wdNoProtection_WPS) {
                wasProtectedInitially = false;
                log.push("[保护] ✅ WPS识别: ProtectionType=-1 表示文档未受保护");
            } else if (currentProtectionType === 0) {
                wasProtectedInitially = false;
                log.push("[保护] ✅ 标准识别: ProtectionType=0 表示文档未受保护");
            } else if (currentProtectionType > 0) {
                wasProtectedInitially = true;
                log.push("[保护] ✅ 识别到文档受保护，类型: " + currentProtectionType);
                
                // 尝试解除保护
                doc.Unprotect(PASSWORD);
                protectionRemoved = true;
                log.push("[保护] ✅ 文档保护已成功解除");
            } else {
                // 对于其他未知值，进行编辑能力测试
                log.push("[保护] ⚠️ 未知ProtectionType值，进行编辑能力测试...");
                
                try {
                    // 尝试插入一个临时段落来测试编辑权限
                    var tempRange = doc.Content;
                    tempRange.Collapse(1); // wdCollapseStart = 1
                    var originalText = tempRange.Text;
                    
                    // 尝试插入测试文本
                    tempRange.Text = "WPS_PROTECTION_TEST";
                    tempRange.Text = originalText; // 立即恢复
                    
                    wasProtectedInitially = false;
                    log.push("[保护] ✅ 编辑测试成功，文档未受保护");
                } catch (e) {
                    wasProtectedInitially = true;
                    log.push("[保护] ❌ 编辑测试失败，文档受保护: " + e.message);
                    
                    // 尝试解除保护
                    doc.Unprotect(PASSWORD);
                    protectionRemoved = true;
                    log.push("[保护] ✅ 文档保护已成功解除");
                }
            }
        } catch (e) {
            var errorMsg = e.description || e.message || String(e);
            log.push("[保护] ⚠️ 检查/解除保护时出错: " + errorMsg);
            
            // 特殊处理WPS的密码错误信息
            if (errorMsg.toLowerCase().includes("password") || 
                errorMsg.toLowerCase().includes("密码") ||
                errorMsg.includes("-2147467259") || // WPS常见错误码
                errorMsg.includes("70")) { // 权限错误
                throw new Error("文档保护密码错误！请确认密码是否正确");
            }
            
            // 其他错误，进行编辑能力测试
            log.push("[保护] ⚠️ 进行编辑能力测试以确定保护状态...");
            
            try {
                var tempRange = doc.Content;
                tempRange.Collapse(1);
                var originalText = tempRange.Text;
                
                tempRange.Text = "WPS_PROTECTION_TEST";
                tempRange.Text = originalText;
                
                wasProtectedInitially = false;
                log.push("[保护] ✅ 编辑测试成功，文档未受保护");
            } catch (editError) {
                wasProtectedInitially = true;
                log.push("[保护] ❌ 编辑测试失败，文档受保护: " + editError.message);
                
                // 尝试解除保护
                doc.Unprotect(PASSWORD);
                protectionRemoved = true;
                log.push("[保护] ✅ 文档保护已成功解除");
            }
        }
        
        // ===== 2. 执行标签转书签自动化 =====
        log.push("\n===== 步骤2: 执行标签转书签自动化 =====");
        
        try {
            标签转书签自动化();
            log.push("[转换] ✅ 标签转书签自动化执行成功");
        } catch (e) {
            var convertError = e.description || e.message || String(e);
            log.push("[转换] ❌ 标签转书签自动化执行失败: " + convertError);
            throw new Error("标签转书签自动化执行失败: " + convertError);
        }
        
        // ===== 3. 仅当文档初始受保护时，才重新应用保护 =====
        if (wasProtectedInitially) {
            log.push("\n===== 步骤3: 重新应用文档保护 =====");
            
            try {
                doc.Protect(wdAllowOnlyFormFields, true, PASSWORD, true, true);
                log.push("[保护] ✅ 文档保护已成功重新应用");
                
                // 验证保护状态
                var newProtectionType = doc.ProtectionType;
                log.push("[保护] 重新保护后的ProtectionType: " + newProtectionType);
            } catch (e) {
                var protectError = e.description || e.message || String(e);
                log.push("[保护] ❌ 重新应用保护时出错: " + protectError);
                throw new Error("文档保护设置失败: " + protectError);
            }
        } else {
            log.push("\n===== 步骤3: 跳过保护设置 =====");
            log.push("[保护] ⚠️ 原始文档初始状态为未受保护，保持无保护状态");
            log.push("[保护] ⚠️ 不会为文档添加任何保护");
        }
        
        // ===== 处理完成 =====
        var endTime = new Date().getTime();
        var duration = Math.round((endTime - startTime) / 1000);
        
        log.push("\n===== 全流程完成 =====");
        log.push("总耗时: " + duration + " 秒");
        log.push("初始保护状态: " + (wasProtectedInitially ? "已保护" : "未保护"));
        log.push("最终保护状态: " + (wasProtectedInitially ? "已保护" : "未保护"));
        
        finalMessage = "✅ 操作成功完成!\n\n" +
                      "总耗时: " + duration + " 秒\n" +
                      (wasProtectedInitially ? 
                       "• 文档保护已重新启用\n" : 
                       "• 文档保持无保护状态\n");
        
    } catch (e) {
        operationSuccess = false;
        var errorMsg = e.description || e.message || String(e);
        log.push("\n===== 流程异常终止 =====");
        log.push("❌ 全局错误: " + errorMsg);
        
        // 仅当文档初始受保护且成功解除了保护时，尝试恢复保护
        if (wasProtectedInitially && protectionRemoved) {
            try {
                log.push("[安全] 尝试恢复文档保护状态...");
                doc.Protect(wdAllowOnlyFormFields, true, PASSWORD, true, true);
                log.push("[安全] ✅ 文档保护状态已恢复");
            } catch (protectError) {
                var protectErrorMsg = protectError.description || protectError.message || String(protectError);
                log.push("[安全] ❌ 恢复保护失败: " + protectErrorMsg);
            }
        }
        
        finalMessage = "❌ 操作失败:\n" + errorMsg + "\n\n" +
                      (wasProtectedInitially ? 
                       "⚠️ 文档保护状态已尝试恢复\n" : 
                       "ℹ️ 文档保持无保护状态\n") +
                      "详细日志请查看控制台";
    } finally {
        // 恢复原始alert函数
        alert = originalAlert;
        
        // 输出日志
        outputToDebugWindow(log);
        
        // 显示最终结果
        alert(finalMessage + 
              (operationSuccess ? "\n\n详细日志请在 JS 控制台中查看" : ""));
        
        Application.ScreenUpdating = true;
    }
}

function outputToDebugWindow(log) {
    try {
        var debugLog = "===== 解锁_转书签_加锁 日志 =====\n";
        debugLog += log.join("\n");
        console.log(debugLog);
    } catch (e) {
        // 忽略错误
    }
}