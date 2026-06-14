// V20260606.1 — 质量巡场未解决问题每日提醒
// 入口：sendQualityTourIssues（每日 08:40 定时 or 手动）
// 逻辑：扫描 2026注塑月度质量巡场记录 中所有 2026 年 sheet（排除"巡场模板"），
//       找出当前状态 ≠ Done 的问题，汇总后邮件发送给 S&C 团队

// ========== 数据源配置 ==========
const _qt_ID_TOUR = "1O-cwRTfW7jB-TSAeDDNvi_LyEM-Dhsi_Y2AqIwbku4c";
const _qt_EXCLUDE_SHEET = "巡场模板";
const _qt_YEAR_FILTER = "2026";

const _qt_ID_PERMISSION = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const _qt_SHEET_USERID = "userID";
const _qt_PERM_PROCESS_COL = 14;           // O列(0-indexed): 工序
const _qt_PERM_ROLE_COL = 15;              // P列: 职位
const _qt_PERM_EMAIL_COL = 9;              // J列: GMail
const _qt_PERM_PROCESS_VAL = "INJ";
const _qt_PERM_ROLE_VAL = "S&C";

// 表头列（第 3 行）
const _qt_HEADERS = ["NO.", "机台号", "品种", "问题描述", "优先级", "跟进措施", "预计完成日期", "责任人", "当前状态"];

// ========== 主入口 ==========
function sendQualityTourIssues(e) {
  const trigger = e ? "定时" : "手动";
  try {
    console.log("开始执行质量巡场问题提醒...");

    // 1. 获取所有符合条件的数据 sheet
    const ssTour = SpreadsheetApp.openById(_qt_ID_TOUR);
    const allSheets = ssTour.getSheets();
    const dataSheets = [];

    for (var i = 0; i < allSheets.length; i++) {
      var sheetName = allSheets[i].getName();
      if (sheetName === _qt_EXCLUDE_SHEET) continue;
      if (sheetName.indexOf(_qt_YEAR_FILTER) === -1) continue;
      dataSheets.push(allSheets[i]);
    }

    console.log("符合条件的 sheet 数量: " + dataSheets.length);

    if (dataSheets.length === 0) {
      writeLog("sendQualityTourIssues", "跳过", "无 2026 年巡场记录 sheet", trigger, "");
      return;
    }

    // 2. 遍历每个 sheet，提取未解决问题
    // sheetResults: [{ sheetName, issues: [{no, machine, product, problem, priority, action, dueDate, owner, status}] }]
    var sheetResults = [];
    var totalUnresolved = 0;

    for (var s = 0; s < dataSheets.length; s++) {
      var sheet = dataSheets[s];
      var data = sheet.getDataRange().getValues();

      if (data.length < 4) continue; // 至少要有表头+数据行

      var issues = [];
      for (var r = 3; r < data.length; r++) {
        var row = data[r];
        var no = String(row[0] || "").trim();

        // 跳过无效行（NO. 为空）
        if (!no) continue;

        var status = String(row[8] || "").trim();

        // 跳过已解决问题（Done，大小写不敏感）
        if (status.toUpperCase() === "DONE") continue;

        // 未解决问题
        issues.push({
          no: no,
          machine: String(row[1] || "").trim(),
          product: String(row[2] || "").trim(),
          problem: String(row[3] || "").trim(),
          priority: String(row[4] || "").trim(),
          action: String(row[5] || "").trim(),
          dueDate: row[6] instanceof Date ? formatVariableAsDate(row[6]) : String(row[6] || "").trim(),
          owner: String(row[7] || "").trim(),
          status: status || "未填写"
        });
      }

      if (issues.length > 0) {
        sheetResults.push({ sheetName: sheet.getName(), issues: issues });
        totalUnresolved += issues.length;
      }

      console.log("Sheet [" + sheet.getName() + "] 未解决问题: " + issues.length);
    }

    console.log("总计未解决问题: " + totalUnresolved);

    // 3. 获取收件人
    var recipients = _qt_getRecipients();

    // 4. 发送邮件
    if (recipients.length > 0) {
      var today = formatVariableAsDate(new Date());
      var subject = "[质量巡场] 未解决问题汇总 " + today;
      var html = _qt_buildEmailHtml(sheetResults, totalUnresolved, today);

      try {
        _qt_sendMail(recipients.join(","), subject, html);
        var summary = "未解决=" + totalUnresolved + "件, 覆盖" + sheetResults.length + "个sheet, TO=" + recipients.length + "人";
        writeLog("sendQualityTourIssues", "成功", summary, trigger, "TO: " + recipients.join(","));
        console.log("邮件发送成功: " + summary);
      } catch (err) {
        writeLog("sendQualityTourIssues", "失败", err.message, trigger, "TO: " + recipients.join(","));
        console.error("发送失败: " + err.message);
      }
    } else {
      writeLog("sendQualityTourIssues", "跳过", "无匹配收件人(O=INJ,P=S&C)", trigger, "");
      console.warn("未找到匹配收件人");
    }

    console.log("质量巡场问题提醒执行完毕");

  } catch (err) {
    console.error(err.stack || err.message);
    try { writeLog("sendQualityTourIssues", "失败", err.message, trigger, err.stack || ""); } catch (e2) { }
  }
}

// ========== 收件人 ==========
/** 从 userID 表读取 O列=INJ 且 P列=S&C 的邮箱 */
function _qt_getRecipients() {
  try {
    var sheet = SpreadsheetApp.openById(_qt_ID_PERMISSION).getSheetByName(_qt_SHEET_USERID);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    var maxCol = Math.max(_qt_PERM_ROLE_COL + 1, _qt_PERM_EMAIL_COL + 1);
    var data = sheet.getRange(1, 1, lastRow, maxCol).getValues();
    var recipients = [];

    for (var i = 2; i < data.length; i++) {
      var process = String(data[i][_qt_PERM_PROCESS_COL] || "").trim();
      var role = String(data[i][_qt_PERM_ROLE_COL] || "").trim();
      var email = String(data[i][_qt_PERM_EMAIL_COL] || "").trim();

      if (process === _qt_PERM_PROCESS_VAL && role === _qt_PERM_ROLE_VAL && email) {
        recipients.push(email.toLowerCase());
      }
    }

    console.log("匹配收件人: " + recipients.length + " 人");
    return recipients;
  } catch (err) {
    console.error("获取收件人失败: " + err.message);
    return [];
  }
}

// ========== Gmail 发送 ==========
function _qt_sendMail(to, subject, htmlBody) {
  GmailApp.sendEmail(to, subject, "", {
    htmlBody: htmlBody,
    name: "Quality Tour Alert"
  });
}

// ========== 邮件 HTML ==========
function _qt_buildEmailHtml(sheetResults, totalUnresolved, today) {
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:960px;margin:0 auto">';

  // 红色标题栏
  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">质量巡场未解决问题汇总</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">扫描范围：2026年巡场记录（TB1/TB2），排除"Done"状态</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + today + ' | 共 ' + totalUnresolved + ' 件未解决</p>';
  html += '</div>';

  html += '<div style="padding:24px">';

  // ===== 无问题 =====
  if (totalUnresolved === 0) {
    html += '<p style="color:#27ae60;font-weight:bold;font-size:16px">★ 所有巡场记录的问题均已解决（Done），无未解决问题。</p>';
  }

  // ===== 按 Sheet 分组展示 =====
  for (var s = 0; s < sheetResults.length; s++) {
    var result = sheetResults[s];
    html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:' + (s > 0 ? '32px' : '0') + '">' + result.sheetName + ' (' + result.issues.length + '件)</h3>';

    var rows = result.issues.map(function(item) {
      return [item.no, item.machine, item.product, item.problem, item.priority, item.action, item.dueDate, item.owner, item.status];
    });
    html += buildHtmlTable(_qt_HEADERS, rows, "#E60012");
  }

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由 Quality Tour Alert 系统自动发送</p>';
  html += '</div></div></body></html>';
  return html;
}
