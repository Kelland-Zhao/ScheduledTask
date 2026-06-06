// V20260606.1 — Milestone邮件提醒（从 Project_Tracking 迁移至 ScheduledScripts）

const PROJECT_TRACKING_SPREADSHEET_ID = "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24";
const PROJECT_SHEET_NAME = "项目总表";

const MILESTONE_COLS = [
  ["模具/自动化改造", 3, 4],
  ["FAT", 5, 6],
  ["现场安装", 7, 8],
  ["IQ/OQ", 9, 10],
  ["工程测试", 11, 12],
  ["PQ", 13, 14],
  ["Mass Production", 15, 16]
];

function milestoneReminder() {
  const projectSpreadsheet = SpreadsheetApp.openById(PROJECT_TRACKING_SPREADSHEET_ID);
  const projectSheet = projectSpreadsheet.getSheetByName(PROJECT_SHEET_NAME);
  const logSheet = saas.getSheetByName("Log");

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, currentTimeZone, "yyyy-MM-dd");

  const existingKeys = new Set();
  const logLr = logSheet.getLastRow();
  if (logLr > 1) {
    logSheet.getRange(2, 4, logLr - 1, 1).getValues().forEach(r => { if (r[0]) existingKeys.add(r[0]); });
  }

  const lastRow = projectSheet.getLastRow();
  if (lastRow <= 1) {
    logSummary(logSheet, 0);
    return;
  }

  const data = projectSheet.getRange(2, 1, lastRow - 1, 18).getValues();
  const leaderMap = {};
  const newLogRows = [];

  data.forEach(row => {
    const projectName = String(row[0] || "").trim();
    const leaderStr = String(row[1] || "").trim();
    const status = String(row[17] || "").trim();
    if (!projectName || status !== "Ongoing") return;

    const toEmail = extractEmail(leaderStr);
    const leaderName = extractName(leaderStr);

    MILESTONE_COLS.forEach(([milestoneName, pIdx, aIdx]) => {
      const planned = row[pIdx];
      const actual = row[aIdx];
      if (!planned || planned === "NA" || actual) return;
      if (!(planned instanceof Date) && isNaN(new Date(planned).getTime())) return;

      const plannedDate = new Date(planned);
      plannedDate.setHours(0, 0, 0, 0);
      const days = Math.round((plannedDate - today) / 86400000);

      let triggerType = "";
      if (days === 0) triggerType = "到期当天";
      else if (days === 1) triggerType = "提前1天";
      else if (days === 2) triggerType = "提前2天";
      else if (days < 0) triggerType = `超期${Math.abs(days)}天`;
      if (!triggerType) return;

      const dedupKey = `${projectName}|${milestoneName}|${todayStr}`;
      if (existingKeys.has(dedupKey)) return;

      const plannedStr = Utilities.formatDate(plannedDate, currentTimeZone, "yyyy-MM-dd");
      const item = { projectName, milestoneName, plannedStr, triggerType, days };
      if (!leaderMap[toEmail]) leaderMap[toEmail] = { leaderName, overdue: [], upcoming: [] };
      if (days < 0) leaderMap[toEmail].overdue.push(item);
      else leaderMap[toEmail].upcoming.push(item);

      newLogRows.push([todayStr, projectName, milestoneName, dedupKey, triggerType]);
      existingKeys.add(dedupKey);
    });
  });

  let itemCount = 0;
  for (const toEmail in leaderMap) {
    if (!toEmail) continue;
    const { leaderName, overdue, upcoming } = leaderMap[toEmail];
    sendMilestoneEmail(toEmail, leaderName, overdue, upcoming, todayStr);
    itemCount += overdue.length + upcoming.length;
  }

  if (newLogRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogRows.length, 5).setValues(newLogRows);
  }
  logSummary(logSheet, itemCount);
}

function extractEmail(str) {
  const m = str.match(/\(([^)]+@[^)]+)\)/);
  return m ? m[1] : "";
}

function extractName(str) {
  return str.replace(/\s*\([^)]*\)\s*/, "").trim();
}

function sendMilestoneEmail(toEmail, leaderName, overdueItems, upcomingItems, todayStr) {
  if (!toEmail) return;
  const hasOverdue = overdueItems.length > 0;
  const subject = hasOverdue
    ? `[项目Milestone逾期] 有 ${overdueItems.length} 个Milestone已逾期 / Project Milestone Overdue Reminder`
    : `[项目Milestone临期] 有 ${upcomingItems.length} 个Milestone即将到期 / Project Milestone Due Soon`;
  const recipients = toEmail === "lyon_zhang@colpal.com"
    ? toEmail
    : toEmail + ",lyon_zhang@colpal.com";
  GmailApp.sendEmail(recipients, subject, "请使用支持HTML的邮件客户端查看此邮件。", {
    htmlBody: generateMilestoneEmailContent(leaderName, overdueItems, upcomingItems, todayStr),
    cc: "kelland_zhao@colpal.com",
    name: "项目Milestone提醒系统"
  });
}

function generateMilestoneEmailContent(leaderName, overdueItems, upcomingItems, todayStr) {
  const hasOverdue = overdueItems.length > 0;
  const accentColor = hasOverdue ? "#f44336" : "#f39c12";
  const darkColor   = hasOverdue ? "#d32f2f" : "#e65100";
  const bgColor     = hasOverdue ? "#ffebee" : "#fff8e1";

  const buildTable = (items, isOverdue) => {
    const headerGrad = isOverdue
      ? "linear-gradient(135deg,#f44336,#d32f2f)"
      : "linear-gradient(135deg,#f39c12,#e67e22)";
    const rowBgAlt = isOverdue ? "#fff5f5" : "#fffbf0";
    let rows = "";
    items.forEach((item, i) => {
      const badgeLabel = item.days === 0
        ? `<span style="display:block;">今日到期</span><span style="display:block;font-size:10px;opacity:0.9;">Due Today</span>`
        : isOverdue
          ? `<span style="display:block;">[逾期] ${Math.abs(item.days)}天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Overdue</span>`
          : `<span style="display:block;">还剩 ${item.days}天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Left</span>`;
      const badgeBg = isOverdue ? "linear-gradient(135deg,#f44336,#d32f2f)" : "linear-gradient(135deg,#f39c12,#e67e22)";
      rows += `
        <tr style="background-color:${i % 2 === 0 ? rowBgAlt : "#ffffff"};">
          <td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">${item.projectName}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${item.milestoneName}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">${item.plannedStr}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">
            <div style="background:${badgeBg};color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;">${badgeLabel}</div>
          </td>
        </tr>`;
    });
    return `
      <div style="overflow-x:auto;margin-bottom:20px;">
        <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">
          <thead>
            <tr style="background:${headerGrad};color:white;">
              <th style="padding:12px;text-align:left;font-weight:600;">项目名称<br><span style="font-size:0.8em;opacity:0.9;">Project</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">Milestone<br><span style="font-size:0.8em;opacity:0.9;">Milestone</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">计划日期<br><span style="font-size:0.8em;opacity:0.9;">Planned Date</span></th>
              <th style="padding:12px;text-align:center;font-weight:600;">${isOverdue ? "逾期天数<br><span style='font-size:0.8em;opacity:0.9;'>Overdue Days</span>" : "剩余天数<br><span style='font-size:0.8em;opacity:0.9;'>Days Left</span>"}</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  };

  let body = `
    <div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;background-color:#f8f9fa;padding:20px;">
      <div style="background:${bgColor};border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid ${accentColor};">
        <h2 style="color:${darkColor};text-align:center;margin-bottom:20px;border-bottom:3px solid ${accentColor};padding-bottom:10px;">
          ${hasOverdue ? "[逾期提醒] 项目Milestone已逾期" : "[临期提醒] 项目Milestone即将到期"}<br>
          <span style="font-size:0.8em;">${hasOverdue ? "Project Milestone Overdue Reminder" : "Project Milestone Due Soon Reminder"}</span>
        </h2>
        <p style="font-size:16px;line-height:1.6;color:${darkColor};">
          您好${leaderName ? " <b>" + leaderName + "</b>" : ""}！（${todayStr}）以下项目 Milestone 需要您关注：<br>
          <span style="font-size:0.9em;opacity:0.85;">Hello${leaderName ? " <b>" + leaderName + "</b>" : ""}! The following project milestones require your attention (${todayStr}):</span>
        </p>
        <p style="font-size:15px;line-height:1.6;color:${darkColor};font-weight:600;background:rgba(0,0,0,0.05);padding:10px 16px;border-radius:6px;margin-top:10px;">
          ⚠️ 请${leaderName ? " " + leaderName : ""}及时跟进并更新项目进度，在表格中填写实际完成日期！<br>
          <span style="font-weight:400;font-size:0.9em;opacity:0.85;">Please${leaderName ? " " + leaderName + "," : ""} update the project progress and fill in the actual completion date promptly!</span>
        </p>
      </div>`;

  if (overdueItems.length > 0) {
    body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">
        <h3 style="color:#d32f2f;border-bottom:2px solid #f44336;padding-bottom:10px;margin-bottom:20px;">
          [逾期] 已逾期 Overdue Milestones（${overdueItems.length}个）
        </h3>
        ${buildTable(overdueItems, true)}
      </div>`;
  }

  if (upcomingItems.length > 0) {
    body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">
        <h3 style="color:#e65100;border-bottom:2px solid #f39c12;padding-bottom:10px;margin-bottom:20px;">
          [临期] 即将到期 Upcoming Milestones（${upcomingItems.length}个）
        </h3>
        ${buildTable(upcomingItems, false)}
      </div>`;
  }

  body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">
        <div style="text-align:center;color:${darkColor};font-size:14px;line-height:1.6;">
          <p style="margin-bottom:10px;font-weight:600;">请及时跟进并在表格中更新实际完成日期！<br>
          <span style="font-size:0.9em;opacity:0.85;">Please follow up and update the actual completion date in the spreadsheet!</span></p>
          <p style="margin:0;font-style:italic;">此邮件由系统自动发送，请勿回复。<br>
          <span style="font-size:0.8em;opacity:0.7;">This email is automatically sent by the system, please do not reply.</span></p>
        </div>
      </div>
    </div>`;

  return body;
}

function logSummary(logSheet, count) {
  const timestamp = Utilities.formatDate(new Date(), currentTimeZone, "yyyy-MM-dd HH:mm:ss");
  logSheet.appendRow([timestamp, count, "Milestone邮件提醒执行完成"]);
}
