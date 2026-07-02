// V20260702.01 — 逾期催办提醒（扫描 PM/点检 Tasklist_history，发送逾期邮件提醒）
// 入口：scanAndRemindOverdue（周一~五 10:33 定时 or 手动）
// 逻辑：扫描 Database_MasterData（PM）和点检数据库的 Tasklist_history，
//       对逾期项按审批链路匹配责任人 → 发送 HTML 邮件（遵循 docs/邮件UI规范.md 模式 C）
// 迁移自：Database_MasterData code.gs（scriptId: 168P7bXzcwBS9b9Fx9u6e-m52zAusUwMED4aOPeCGZLbsuHVIlKT9rpT1）
// 修复：邮箱重复拼接 @colpal.com 的 bug（Database for Web 存储完整邮箱，不再追加域名后缀）

// ========== 配置常量 ==========
const OR_CONFIG = {
  PM_DB_ID: "1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U",
  INSPECTION_DB_ID: "1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY",
  WEB_APP_URL: "https://script.google.com/a/macros/colpal.com/s/AKfycbxpDYL02i5FaFzUDcIoW3siG2U94cvWUUnz_F5x2BO1jnrXoMGzFQH-jw9C4nvZ7FE/exec",
  OVERDUE_THRESHOLD_APPROVAL: 3,     // 待审批逾期阈值（天）
  OVERDUE_THRESHOLD_DISSEMINATE: 5,  // 待发放逾期阈值（天）
  SENDER_NAME: "逾期催办提醒系统",
};

const OR_FUNC_NAME = "scanAndRemindOverdue";

// ========== 主入口 ==========
function scanAndRemindOverdue(e) {
  const trigger = e ? "定时" : "手动";
  const startTime = new Date();

  try {
    writeLog(OR_FUNC_NAME, "成功", "开始执行逾期催办扫描，时间: " + formatVariableAsDateHms(startTime), trigger, "");

    const pmConfig = { dbId: OR_CONFIG.PM_DB_ID, dbName: "保养/ PM" };
    const inConfig = { dbId: OR_CONFIG.INSPECTION_DB_ID, dbName: "点检/ Inspection" };

    let totalSent = 0;
    try { totalSent += _or_scanSingleDatabase(pmConfig, trigger); } catch (e2) {
      writeLog(OR_FUNC_NAME, "失败", "PM 催办扫描异常: " + e2.message, trigger, "");
    }
    try { totalSent += _or_scanSingleDatabase(inConfig, trigger); } catch (e2) {
      writeLog(OR_FUNC_NAME, "失败", "Inspection 催办扫描异常: " + e2.message, trigger, "");
    }

    const endTime = new Date();
    const duration = ((endTime.getTime() - startTime.getTime()) / 1000).toFixed(1);
    writeLog(OR_FUNC_NAME, "成功", "逾期催办执行完成，耗时: " + duration + "秒，共发送 " + totalSent + " 封催办邮件", trigger, "");

  } catch (err) {
    writeLog(OR_FUNC_NAME, "失败", err.message, trigger, err.stack || "");
  }
}

// ========== 单数据库扫描 ==========
function _or_scanSingleDatabase(config, trigger) {
  const ss = SpreadsheetApp.openById(config.dbId);
  const wsHistory = ss.getSheetByName("Tasklist_history");
  const wsMail = ss.getSheetByName("Database for Web");

  if (!wsHistory || !wsMail) {
    writeLog(OR_FUNC_NAME, "失败", config.dbName + ": 找不到 Tasklist_history 或 Database for Web 工作表", trigger, "");
    return 0;
  }

  const historyData = wsHistory.getRange(2, 1, wsHistory.getLastRow() - 1, 15).getValues();
  const mailData = wsMail.getRange(2, 1, wsMail.getLastRow() - 1, 7).getValues();
  writeLog(OR_FUNC_NAME, "成功", config.dbName + ": 读取到 " + historyData.length + " 条历史记录, " + mailData.length + " 条邮件配置", trigger, "");

  // 构建 MachineType → 邮箱配置 的查找表
  const emailLookup = {};
  mailData.forEach(function (r) {
    const mt = (r[0] || "").toString().trim();
    if (!mt) return;
    emailLookup[mt] = {
      Mail_Approve1: (r[2] || "").toString().trim(),
      Mail_Approve2: (r[3] || "").toString().trim(),
      Mail_Disseninate: (r[4] || "").toString().trim(),
      Mail_CC: (r[5] || "").toString().trim(),
      Mail_Production: (r[6] || "").toString().trim(),
    };
  });

  // 扫描逾期项，按收件人聚合
  const now = new Date();
  const remindersByRecipient = {};

  historyData.forEach(function (row) {
    const status = (row[12] || "").toString().trim();
    if (status !== "待审批/ Pending" && status !== "待发放/ Wait for Dissminater") return;

    const submitDate = _or_extractTimestamp(row[2]);
    if (!submitDate) return;

    const daysElapsed = Math.floor((now.getTime() - submitDate.getTime()) / 86400000);
    const threshold = status === "待发放/ Wait for Dissminater"
      ? OR_CONFIG.OVERDUE_THRESHOLD_DISSEMINATE
      : OR_CONFIG.OVERDUE_THRESHOLD_APPROVAL;
    if (daysElapsed < threshold) return;

    const responsibleEmail = _or_determineEmail(row, emailLookup);
    if (!responsibleEmail) return;

    if (!remindersByRecipient[responsibleEmail]) remindersByRecipient[responsibleEmail] = [];
    remindersByRecipient[responsibleEmail].push({
      machineType: (row[0] || "").toString().trim(),
      status: status,
      daysElapsed: daysElapsed,
      reason: (row[3] || "").toString().trim(),
      submitMail: (row[13] || "").toString().trim(),
    });
  });

  // 逐收件人发送邮件
  var sentCount = 0;
  for (var emailAddr in remindersByRecipient) {
    var items = remindersByRecipient[emailAddr];
    // 修复：Database for Web 已存储完整邮箱，不再拼接 @colpal.com
    var recipientEmail = emailAddr.indexOf("@") !== -1 ? emailAddr : emailAddr + "@colpal.com";
    var htmlBody = _or_buildEmailBody(config.dbName, items);
    var subject = "【逾期提醒】" + config.dbName + " — " + items.length + " 项待处理 / Overdue " + config.dbName + " Items — " + items.length + " Pending";

    try {
      GmailApp.sendEmail(recipientEmail, subject, "请使用支持 HTML 的邮件客户端查看此邮件。", {
        htmlBody: htmlBody,
        name: OR_CONFIG.SENDER_NAME,
      });
      sentCount++;
      writeLog(OR_FUNC_NAME, "成功", config.dbName + ": 已发送催办邮件至 " + recipientEmail + " (" + items.length + " 项)", trigger, "");
    } catch (mailErr) {
      writeLog(OR_FUNC_NAME, "失败", config.dbName + ": 发送邮件至 " + recipientEmail + " 失败: " + mailErr.message, trigger, "");
    }
  }

  writeLog(OR_FUNC_NAME, "成功", config.dbName + ": 催办扫描完成，发送 " + sentCount + " 封邮件，涉及 " + Object.keys(remindersByRecipient).length + " 位审批人", trigger, "");
  return sentCount;
}

// ========== 邮件正文生成（follow 邮件UI规范 模式 C + 卡片表格） ==========
function _or_buildEmailBody(dbName, items) {
  // 数据行
  var tableRows = "";
  items.forEach(function (item, i) {
    var rowBg, daysCell;

    if (item.daysElapsed > 7) {
      rowBg = "#fff5f5";
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;' +
        'background:linear-gradient(135deg,#e74c3c,#c0392b);color:white;font-weight:600;font-size:12px;">' +
        '<span style="display:block;">[逾期] ' + item.daysElapsed + '天</span>' +
        '<span style="display:block;font-size:10px;opacity:0.9;">Days</span></span>';
    } else if (item.daysElapsed > 3) {
      rowBg = "#fffbf0";
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;' +
        'background:linear-gradient(135deg,#f39c12,#e67e22);color:white;font-weight:600;font-size:12px;">' +
        '<span style="display:block;">' + item.daysElapsed + '天</span>' +
        '<span style="display:block;font-size:10px;opacity:0.9;">Days</span></span>';
    } else {
      rowBg = "#ffffff";
      daysCell = item.daysElapsed + " 天";
    }

    tableRows += '<tr style="background-color:' + rowBg + ';">' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + item.machineType + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + item.status + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + daysCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;max-width:200px;word-wrap:break-word;">' + (item.reason || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + (item.submitMail || "—") + '</td>' +
      '</tr>';
  });

  var htmlBody =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">' +

    // 模式 C：白色卡片 + 左侧红色色条
    '<div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;border-left:5px solid #f44336;">' +
    '<h2 style="color:#d32f2f;margin:0 0 8px 0;">【逾期提醒】' + dbName + ' 待审批/待发放项</h2>' +
    '<p style="color:#c62828;margin:0;font-size:14px;">Overdue Approval / Dissemination Items — ' + items.length + ' 项待处理</p>' +
    '</div>' +

    // 正文卡片
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;">' +
    '<p style="font-size:15px;color:#34495e;line-height:1.8;">' +
    '您好！以下 ' + dbName + ' 待审批/待发放项已逾期，请尽快处理：<br>' +
    '<span style="font-size:13px;opacity:0.8;">Hello! The following ' + dbName + ' items are overdue. Please process as soon as possible.</span>' +
    '</p>' +

    // 渐变表头卡片表格
    '<div style="overflow-x:auto;margin:24px 0;">' +
    '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.1);font-size:13px;">' +
    '<thead><tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">' +
    '<th style="padding:10px;text-align:center;font-weight:600;">机型<br><span style="font-size:11px;opacity:0.85;">Machine</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">状态<br><span style="font-size:11px;opacity:0.85;">Status</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">等待天数<br><span style="font-size:11px;opacity:0.85;">Days</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">变更原因<br><span style="font-size:11px;opacity:0.85;">Reason</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">申请人<br><span style="font-size:11px;opacity:0.85;">Applier</span></th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table></div>' +

    '<p style="margin-top:24px;font-size:14px;color:#34495e;">' +
    '请点击下方链接登录系统处理：<br>' +
    '<span style="font-size:12px;opacity:0.8;">Please click the link below to log in and process:</span></p>' +
    '<a href="' + OR_CONFIG.WEB_APP_URL + '" ' +
    'style="display:inline-block;padding:10px 20px;background:#E60012;color:white;' +
    'text-decoration:none;border-radius:4px;font-size:14px;margin-top:8px;">' +
    '任务清单变更管理 / Tasklist MoC</a></div>' +

    // 尾部 模式 B：双语
    '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;">' +
    '<p style="margin-bottom:10px;">请及时查看并处理相关逾期事项。<br>' +
    '<span style="font-size:0.9em;opacity:0.8;">Please review and handle related overdue items promptly.</span></p>' +
    '<p style="margin:0;font-style:italic;">此邮件由逾期催办提醒系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.8;">This email is automatically sent by the Overdue Reminder System, please do not reply.</span></p>' +
    '</div></div></body></html>';

  return htmlBody;
}

// ========== 辅助函数 ==========

/** 从操作者字符串提取时间戳 */
function _or_extractTimestamp(operatorStr) {
  if (!operatorStr || typeof operatorStr !== "string") return null;
  var m = operatorStr.match(/(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})$/);
  if (m) return new Date(m[1]);
  m = operatorStr.match(/(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2})$/);
  if (m) return new Date(m[1]);
  return null;
}

/** 根据审批状态确定责任人邮箱 */
function _or_determineEmail(row, emailLookup) {
  var machineType = (row[0] || "").toString().trim();
  var emailConfig = emailLookup[machineType];
  if (!emailConfig) return null;

  var status = (row[12] || "").toString().trim();
  var productionApproval = (row[10] || "").toString().trim();
  var productionApprover = (row[4] || "").toString().trim();
  var approver1 = (row[6] || "").toString().trim();
  var approver2 = (row[8] || "").toString().trim();

  if (status === "待发放/ Wait for Dissminater") return emailConfig.Mail_Disseninate;
  if (status === "待审批/ Pending") {
    if (productionApproval === "Y" && !productionApprover) return emailConfig.Mail_Production;
    if (!approver1) return emailConfig.Mail_Approve1;
    if (approver1 && !approver2) return emailConfig.Mail_Approve2;
  }
  return null;
}
