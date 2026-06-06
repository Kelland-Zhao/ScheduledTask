// V20260606.3 — Milestone邮件提醒（自定义milestone + userID查上级/管理员）

const PROJECT_TRACKING_SPREADSHEET_ID = "1aoQDjeWU9Xa9clloyTwiXL6WS62tYVbB0-VOavpAgAM";
const PROJECT_SHEET_NAME = "项目总表";

// userID 列索引（0-based）
const UID_GMail_COL = 9;        // J列：人员邮箱
const UID_PROCESS_COL = 14;     // O列：工序
const UID_BH_COL = 59;          // BH列：项目跟进权限管理（管理员标记）
const UID_BI_COL = 60;          // BI列：直线上级邮箱

/**
 * 从 userID 表读取：
 * - supervisorMap: Leader邮箱 → 上级邮箱（BI列）
 * - injAdminEmails: INJ工序管理员（O列=INJ 且 BH列="INJ工序管理员"）
 * - newProductAdminEmails: 新品自动化项目管理员（BH列="新品自动化项目管理员"）
 */
function _ms_buildUserLookup() {
  var sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var result = { supervisorMap: {}, injAdminEmails: [], newProductAdminEmails: [] };
  if (lastRow < 3) return result;

  var data = sheet.getRange(3, 1, lastRow - 2, 61).getValues(); // A~BI 共61列
  var injAdminSet = {};
  var newProductAdminSet = {};

  data.forEach(function(row) {
    var email = String(row[UID_GMail_COL] || "").trim().toLowerCase();
    if (!email) return;
    var supervisorEmail = String(row[UID_BI_COL] || "").trim().toLowerCase();
    if (supervisorEmail) {
      result.supervisorMap[email] = supervisorEmail;
    }
    // BH列角色值
    var bhVal = String(row[UID_BH_COL] || "").trim();
    if (bhVal === "INJ工序管理员" && !injAdminSet[email]) {
      injAdminSet[email] = true;
      result.injAdminEmails.push(email);
    }
    if (bhVal === "新品自动化项目管理员" && !newProductAdminSet[email]) {
      newProductAdminSet[email] = true;
      result.newProductAdminEmails.push(email);
    }
  });

  return result;
}

/** 从 JSON name 提取中文显示名（取 / 前部分） */
function _ms_extractDisplayName(fullName) {
  if (!fullName) return "";
  var idx = fullName.indexOf("/");
  if (idx > -1) {
    return fullName.substring(0, idx).trim();
  }
  return fullName.trim();
}

function milestoneReminder() {
  var projectSpreadsheet = SpreadsheetApp.openById(PROJECT_TRACKING_SPREADSHEET_ID);
  var projectSheet = projectSpreadsheet.getSheetByName(PROJECT_SHEET_NAME);
  var logSheet = saas.getSheetByName("Log");

  // 读取已发送记录（去重）
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var todayStr = Utilities.formatDate(today, currentTimeZone, "yyyy-MM-dd");

  var existingKeys = new Set();
  var logLr = logSheet.getLastRow();
  if (logLr > 1) {
    logSheet.getRange(2, 4, logLr - 1, 1).getValues().forEach(function(r) { if (r[0]) existingKeys.add(r[0]); });
  }

  var lastRow = projectSheet.getLastRow();
  if (lastRow <= 1) {
    logSummary(logSheet, 0);
    return;
  }

  // 预加载 userID 查找表
  var userLookup = _ms_buildUserLookup();

  // 新格式：A-I 共 9 列
  var data = projectSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var leaderMap = {};
  var newLogRows = [];

  data.forEach(function(row) {
    var projectName = String(row[0] || "").trim();
    var leaderStr = String(row[1] || "").trim();
    var status = String(row[3] || "").trim();             // D列：状态
    if (!projectName || status !== "Ongoing") return;

    var toEmail = extractEmail(leaderStr).toLowerCase();
    var leaderName = extractName(leaderStr);

    // 解析 E列 Milestones_JSON
    var msJsonRaw = row[4];
    var milestones = [];
    if (msJsonRaw) {
      try {
        milestones = JSON.parse(String(msJsonRaw));
      } catch (e) {
        console.warn("JSON 解析失败: " + projectName + " - " + e.message);
        writeLog("milestoneReminder", "失败", "JSON解析失败: " + projectName, "定时", e.message);
        return;
      }
    }

    milestones.forEach(function(ms) {
      var planned = String(ms.planned || "").trim();
      var msStatus = String(ms.status || "").trim();
      if (!planned || planned === "NA" || msStatus === "已完成") return;

      // 直接用 JSON name 提取显示名（不再关键词匹配）
      var displayName = _ms_extractDisplayName(String(ms.name || ""));
      if (!displayName) return;

      // 解析日期（JSON 中是字符串 "yyyy-MM-dd"）
      var plannedDate = new Date(planned + "T00:00:00");
      if (isNaN(plannedDate.getTime())) return;
      plannedDate.setHours(0, 0, 0, 0);

      var days = Math.round((plannedDate - today) / 86400000);

      var triggerType = "";
      if (days === 0) triggerType = "到期当天";
      else if (days === 1) triggerType = "提前1天";
      else if (days === 2) triggerType = "提前2天";
      else if (days < 0) triggerType = "超期" + Math.abs(days) + "天";
      if (!triggerType) return;

      var dedupKey = projectName + "|" + displayName + "|" + todayStr;
      if (existingKeys.has(dedupKey)) return;

      var plannedStr = Utilities.formatDate(plannedDate, currentTimeZone, "yyyy-MM-dd");

      // 事项责任人
      var ownerName = String(ms.owner || "").trim();
      var ownerEmail = String(ms.ownerEmail || "").trim().toLowerCase();

      var item = {
        projectName: projectName,
        milestoneName: displayName,
        plannedStr: plannedStr,
        triggerType: triggerType,
        days: days,
        ownerName: ownerName,
        ownerEmail: ownerEmail
      };

      // 按 Leader 邮箱分组（用于合并邮件）
      var groupKey = toEmail || "no-leader";
      if (!leaderMap[groupKey]) {
        leaderMap[groupKey] = {
          leaderName: leaderName,
          leaderEmail: toEmail,
          overdue: [],
          upcoming: [],
          ownerEmails: {}
        };
      }
      if (days < 0) {
        leaderMap[groupKey].overdue.push(item);
      } else {
        leaderMap[groupKey].upcoming.push(item);
      }
      // 收集所有 ownerEmail（用于 TO）
      if (ownerEmail && ownerEmail !== toEmail) {
        leaderMap[groupKey].ownerEmails[ownerEmail] = true;
      }

      newLogRows.push([todayStr, projectName, displayName, dedupKey, triggerType]);
      existingKeys.add(dedupKey);
    });
  });

  var itemCount = 0;
  for (var key in leaderMap) {
    if (key === "no-leader" && Object.keys(leaderMap[key].overdue).length === 0 && Object.keys(leaderMap[key].upcoming).length === 0) continue;
    var entry = leaderMap[key];
    sendMilestoneEmail(entry, todayStr, userLookup);
    itemCount += entry.overdue.length + entry.upcoming.length;
  }

  if (newLogRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogRows.length, 5).setValues(newLogRows);
  }
  logSummary(logSheet, itemCount);
}

function extractEmail(str) {
  var m = str.match(/\(([^)]+@[^)]+)\)/);
  return m ? m[1] : "";
}

function extractName(str) {
  return str.replace(/\s*\([^)]*\)\s*/, "").trim();
}

/**
 * 发送邮件
 * @param {Object} entry - { leaderName, leaderEmail, overdue[], upcoming[], ownerEmails{} }
 */
function sendMilestoneEmail(entry, todayStr, userLookup) {
  var leaderEmail = entry.leaderEmail;
  var ownerEmails = Object.keys(entry.ownerEmails);
  var overdueItems = entry.overdue;
  var upcomingItems = entry.upcoming;

  // 构建 TO 列表：新品自动化项目管理员 + 各事项责任人
  var toList = [];
  userLookup.newProductAdminEmails.forEach(function(e) {
    if (toList.indexOf(e) === -1) toList.push(e);
  });
  ownerEmails.forEach(function(e) {
    if (toList.indexOf(e) === -1) toList.push(e);
  });
  if (toList.length === 0) return;

  // 构建 CC 列表：Leader + Leader上级 + INJ工序管理员 + kelland
  var ccList = [];

  // Leader
  if (leaderEmail && ccList.indexOf(leaderEmail) === -1 && toList.indexOf(leaderEmail) === -1) {
    ccList.push(leaderEmail);
  }

  // Leader 上级（BI列）
  if (leaderEmail) {
    var supervisorEmail = userLookup.supervisorMap[leaderEmail];
    if (supervisorEmail && ccList.indexOf(supervisorEmail) === -1 && toList.indexOf(supervisorEmail) === -1) {
      ccList.push(supervisorEmail);
    }
  }

  // INJ 工序管理员（BH列="INJ工序管理员"）
  userLookup.injAdminEmails.forEach(function(e) {
    if (ccList.indexOf(e) === -1 && toList.indexOf(e) === -1) {
      ccList.push(e);
    }
  });

  // 保留 Kelland 在 CC
  var kellandEmail = "kelland_zhao@colpal.com";
  if (ccList.indexOf(kellandEmail) === -1 && toList.indexOf(kellandEmail) === -1) {
    ccList.push(kellandEmail);
  }

  var hasOverdue = overdueItems.length > 0;
  var subject = hasOverdue
    ? "【项目逾期】有 " + overdueItems.length + " 个Milestone已逾期"
    : "【项目临期】有 " + upcomingItems.length + " 个Milestone即将到期";

  GmailApp.sendEmail(toList.join(","), subject, "请使用支持HTML的邮件客户端查看此邮件。", {
    htmlBody: generateMilestoneEmailContent(entry.leaderName, overdueItems, upcomingItems, todayStr),
    cc: ccList.join(","),
    name: "项目Milestone提醒系统"
  });
}

function generateMilestoneEmailContent(leaderName, overdueItems, upcomingItems, todayStr) {
  var hasOverdue = overdueItems.length > 0;
  var accentColor = hasOverdue ? "#f44336" : "#f39c12";
  var darkColor   = hasOverdue ? "#d32f2f" : "#e65100";
  var bgColor     = hasOverdue ? "#ffebee" : "#fff8e1";

  var buildTable = function(items, isOverdue) {
    var headerGrad = isOverdue
      ? "linear-gradient(135deg,#f44336,#d32f2f)"
      : "linear-gradient(135deg,#f39c12,#e67e22)";
    var rowBgAlt = isOverdue ? "#fff5f5" : "#fffbf0";
    var rows = "";
    items.forEach(function(item, i) {
      var badgeLabel;
      if (item.days === 0) {
        badgeLabel = "<span style='display:block;'>今日到期</span><span style='display:block;font-size:10px;opacity:0.9;'>Due Today</span>";
      } else if (isOverdue) {
        badgeLabel = "<span style='display:block;'>[逾期] " + Math.abs(item.days) + "天</span><span style='display:block;font-size:10px;opacity:0.9;'>Days Overdue</span>";
      } else {
        badgeLabel = "<span style='display:block;'>还剩 " + item.days + "天</span><span style='display:block;font-size:10px;opacity:0.9;'>Days Left</span>";
      }
      var badgeBg = isOverdue ? "linear-gradient(135deg,#f44336,#d32f2f)" : "linear-gradient(135deg,#f39c12,#e67e22)";
      rows += "\n        <tr style='background-color:" + (i % 2 === 0 ? rowBgAlt : "#ffffff") + ";'>" +
        "<td style='padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;'>" + item.projectName + "</td>" +
        "<td style='padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;'>" + item.milestoneName + "</td>" +
        "<td style='padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;'>" + item.plannedStr + "</td>" +
        "<td style='padding:12px;border-bottom:1px solid #e9ecef;text-align:center;'>" +
        "<div style='background:" + badgeBg + ";color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;'>" + badgeLabel + "</div>" +
        "</td></tr>";
    });
    return "\n      <div style='overflow-x:auto;margin-bottom:20px;'>" +
      "<table style='width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);'>" +
      "<thead><tr style='background:" + headerGrad + ";color:white;'>" +
      "<th style='padding:12px;text-align:left;font-weight:600;'>项目名称<br><span style='font-size:0.8em;opacity:0.9;'>Project</span></th>" +
      "<th style='padding:12px;text-align:left;font-weight:600;'>Milestone<br><span style='font-size:0.8em;opacity:0.9;'>Milestone</span></th>" +
      "<th style='padding:12px;text-align:left;font-weight:600;'>计划日期<br><span style='font-size:0.8em;opacity:0.9;'>Planned Date</span></th>" +
      "<th style='padding:12px;text-align:center;font-weight:600;'>" + (isOverdue ? "逾期天数<br><span style='font-size:0.8em;opacity:0.9;'>Overdue Days</span>" : "剩余天数<br><span style='font-size:0.8em;opacity:0.9;'>Days Left</span>") + "</th>" +
      "</tr></thead><tbody>" + rows + "</tbody></table></div>";
  };

  var body = "\n    <div style='font-family:Arial,sans-serif;max-width:800px;margin:0 auto;background-color:#f8f9fa;padding:20px;'>" +
    "<div style='background:" + bgColor + ";border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid " + accentColor + ";'>" +
    "<h2 style='color:" + darkColor + ";text-align:center;margin-bottom:20px;border-bottom:3px solid " + accentColor + ";padding-bottom:10px;'>" +
    (hasOverdue ? "【逾期提醒】项目Milestone已逾期" : "【临期提醒】项目Milestone即将到期") + "<br>" +
    "<span style='font-size:0.8em;'>Project Milestone Reminder</span></h2>" +
    "<p style='font-size:16px;line-height:1.6;color:" + darkColor + ";'>" +
    "您好" + (leaderName ? " <b>" + leaderName + "</b>" : "") + "！（" + todayStr + "）以下项目 Milestone 需要您关注：<br>" +
    "<span style='font-size:0.9em;opacity:0.85;'>Hello" + (leaderName ? " <b>" + leaderName + "</b>" : "") + "! The following project milestones require your attention (" + todayStr + "):</span></p>" +
    "<p style='font-size:15px;line-height:1.6;color:" + darkColor + ";font-weight:600;background:rgba(0,0,0,0.05);padding:10px 16px;border-radius:6px;margin-top:10px;'>" +
    "⚠️ 请" + (leaderName ? " " + leaderName : "") + "及时跟进并更新项目进度！<br>" +
    "<span style='font-weight:400;font-size:0.9em;opacity:0.85;'>Please update the project progress promptly!</span></p></div>";

  if (overdueItems.length > 0) {
    body += "\n      <div style='background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;'>" +
      "<h3 style='color:#d32f2f;border-bottom:2px solid #f44336;padding-bottom:10px;margin-bottom:20px;'>" +
      "[逾期] 已逾期 Overdue Milestones（" + overdueItems.length + "个）</h3>" +
      buildTable(overdueItems, true) + "</div>";
  }

  if (upcomingItems.length > 0) {
    body += "\n      <div style='background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;'>" +
      "<h3 style='color:#e65100;border-bottom:2px solid #f39c12;padding-bottom:10px;margin-bottom:20px;'>" +
      "[临期] 即将到期 Upcoming Milestones（" + upcomingItems.length + "个）</h3>" +
      buildTable(upcomingItems, false) + "</div>";
  }

  body += "\n      <div style='background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;'>" +
    "<div style='text-align:center;color:" + darkColor + ";font-size:14px;line-height:1.6;'>" +
    "<p style='margin-bottom:10px;font-weight:600;'>请及时跟进并更新项目进度！<br>" +
    "<span style='font-size:0.9em;opacity:0.85;'>Please follow up and update the project progress!</span></p>" +
    "<p style='margin:0;font-style:italic;'>此邮件由系统自动发送，请勿回复。<br>" +
    "<span style='font-size:0.8em;opacity:0.7;'>This email is automatically sent by the system, please do not reply.</span></p></div></div></div>";

  return body;
}

function logSummary(logSheet, count) {
  var timestamp = Utilities.formatDate(new Date(), currentTimeZone, "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([timestamp, count, "Milestone邮件提醒执行完成"]);
}
