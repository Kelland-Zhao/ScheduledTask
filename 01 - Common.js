// V20260524.1 — 通用工具

// ========== 日期格式化 ==========
function formatVariableAsDate(variable) {
  if (!variable) return "";
  return Utilities.formatDate(new Date(variable), currentTimeZone, "yyyy-MM-dd");
}

function formatVariableAsDateHm(variable) {
  if (!variable) return "";
  return Utilities.formatDate(new Date(variable), currentTimeZone, "yyyy-MM-dd HH:mm");
}

function formatVariableAsDateHms(variable) {
  if (!variable) return "";
  return Utilities.formatDate(new Date(variable), currentTimeZone, "yyyy-MM-dd HH:mm:ss");
}

/** 当月第几周 */
function getWeekInMonth(variable) {
  let result = 1;
  if (variable) {
    const date = new Date(variable);
    const w = date.getDay();
    const d = date.getDate();
    result = Math.ceil((d + 6 - w) / 7);
  }
  return result;
}

/** 今日 00:00:00 */
function todayDate() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/** 明日 00:00:00 */
function tomorrowDate() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  d.setHours(0, 0, 0, 0);
  return d;
}

/** 目标日期距今日天数：负数=逾期，0=今天，正数=未来 */
function diffDays(targetDate) {
  if (!targetDate) return null;
  return Math.floor((targetDate.getTime() - todayDate().getTime()) / 86400000);
}

// ========== Log 写入 ==========
function writeLog(funcName, status, detail, trigger, remark) {
  if (!sbnLog) return;
  sbnLog.appendRow([
    formatVariableAsDateHms(new Date()),
    funcName,
    status,
    detail || "",
    trigger || "",
    remark || ""
  ]);
}

// ========== 企微 Webhook 推送 ==========
function postWechat(content) {
  const payload = { msgtype: "markdown", markdown: { content: content } };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(WECHAT_WEBHOOK_URL, options);
  return JSON.parse(resp.getContentText());
}

// ========== 权限表读取 ==========
/** 返回 {email(lowercase):true} 白名单 map；列不存在时返回 null */
function getAuthorizedEmails() {
  try {
    const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3) return {};
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headerRow = data[1]; // 第2行字段名
    const alertColIdx = headerRow.indexOf(PERMISSION_ALERT_HEADER);
    if (alertColIdx < 0) {
      console.warn("权限表未找到列: " + PERMISSION_ALERT_HEADER);
      return null;
    }
    const result = {};
    for (let i = 2; i < data.length; i++) {
      const email = String(data[i][PERMISSION_EMAIL_COL_IDX] || "").trim().toLowerCase();
      const flag = String(data[i][alertColIdx] || "").trim();
      if (email && flag === "Y") result[email] = true;
    }
    return result;
  } catch (e) {
    console.error("getAuthorizedEmails 失败: " + e.message);
    return null;
  }
}

// ========== SmartMeeting 数据读取 ==========
/**
 * 读取 INJ SDM 跟进项
 * @returns {Array} [{date, topic, followUp, assignee, dueDate, dueDateObj, status, isDone}]
 */
function getFollowUpRows() {
  const sheet = SpreadsheetApp.openById(SMARTMEETING_SPREADSHEET_ID).getSheetByName(SMARTMEETING_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - DATA_MONTHS);
  cutoff.setHours(0, 0, 0, 0);

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;
    const meetingDate = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (isNaN(meetingDate.getTime()) || meetingDate < cutoff) continue;

    const topic = String(row[3] || "").trim();
    if (topic !== MEETING_TOPIC_FILTER) continue;

    // 跟进人可能多人，逗号分隔，统一转小写
    const assigneesRaw = String(row[7] || "").trim();
    if (!assigneesRaw || assigneesRaw === "-") continue;
    const assignees = assigneesRaw.split(",")
      .map(function(s) { return s.trim().toLowerCase(); })
      .filter(function(s) { return s.length > 0; });
    if (assignees.length === 0) continue;

    let dueDateObj = null;
    let dueDateStr = "";
    const dueVal = row[8];
    if (dueVal) {
      const tmp = dueVal instanceof Date ? new Date(dueVal) : new Date(dueVal);
      if (!isNaN(tmp.getTime())) {
        tmp.setHours(0, 0, 0, 0);
        dueDateObj = tmp;
        dueDateStr = formatVariableAsDate(tmp);
      }
    }

    // E列 记录时间（跟进项创建时间）
    const recordVal = row[4];
    let recordDateObj = null;
    let recordDateStr = "";
    if (recordVal) {
      const tmp = recordVal instanceof Date ? new Date(recordVal) : new Date(recordVal);
      if (!isNaN(tmp.getTime())) {
        tmp.setHours(0, 0, 0, 0);
        recordDateObj = tmp;
        recordDateStr = formatVariableAsDate(tmp);
      }
    }

    const status = String(row[9] || "").trim();
    rows.push({
      date: formatVariableAsDate(meetingDate),  // A列 会议日期
      recordDate: recordDateStr,                  // E列 记录时间（字符串）
      recordDateObj: recordDateObj,               // E列 记录时间（Date 对象）
      topic: topic,
      followUp: String(row[5] || "").trim(),
      assignees: assignees,         // 数组（已小写）
      assigneesRaw: assigneesRaw,   // 原始字符串（用于调试）
      dueDate: dueDateStr,
      dueDateObj: dueDateObj,
      status: status,
      isDone: status === DONE_STATUS
    });
  }
  return rows;
}

// ========== HTML 表格（用于 Gmail）==========
function buildHtmlTable(headers, rows, headerBg) {
  let html = '<table border="1" style="border-collapse:collapse;width:100%;font-size:13px">';
  html += '<tr style="background:' + headerBg + ';color:white">';
  headers.forEach(function(h) {
    html += '<th style="padding:10px;text-align:left">' + escapeHtml(h) + '</th>';
  });
  html += '</tr>';
  rows.forEach(function(row, idx) {
    const bg = idx % 2 === 0 ? "#f8f9fa" : "#ffffff";
    html += '<tr style="background:' + bg + '">';
    row.forEach(function(cell) {
      html += '<td style="padding:10px">' + escapeHtml(String(cell || "-")) + '</td>';
    });
    html += '</tr>';
  });
  html += '</table>';
  return html;
}

function escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ========== Gmail 发送 ==========
function sendMail(to, subject, htmlBody, cc) {
  const options = { htmlBody: htmlBody, name: "SmartMeeting Alert" };
  if (cc) options.cc = cc;
  GmailApp.sendEmail(to, subject, "", options);
}

// ========== 本月跟进项完成率统计 ==========
/**
 * 按 E列记录时间过滤当月，统计各授权跟进人 total/done
 * @returns { email: { total, done } }
 */
function getCurrentMonthStats(rows, authorized) {
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  monthStart.setHours(0, 0, 0, 0);
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  monthEnd.setHours(23, 59, 59, 999);

  const stats = {};
  rows.forEach(function(r) {
    if (!r.recordDateObj || r.recordDateObj < monthStart || r.recordDateObj > monthEnd) return;
    r.assignees.forEach(function(email) {
      if (!authorized[email]) return;
      if (!stats[email]) stats[email] = { total: 0, done: 0 };
      stats[email].total++;
      if (r.isDone) stats[email].done++;
    });
  });
  return stats;
}
