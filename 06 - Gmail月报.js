// V20260524.1 — SmartMeeting Alert Gmail 月报
// 入口：sendGmailReports（周一 08:00 定时 or 手动）

/** 周一 08:00 发送合并月报：当月完成率 + 全员未完成清单 */
function sendGmailReports(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const authorized = getAuthorizedEmails();
    if (authorized === null) {
      writeLog("sendGmailReports", "失败", "权限列 SmartMeeting_Alert 不存在", trigger, "");
      return;
    }

    // 当月数据范围（按会议日期过滤）
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    monthStart.setHours(0, 0, 0, 0);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    monthEnd.setHours(23, 59, 59, 999);
    const monthLabel = Utilities.formatDate(now, currentTimeZone, "yyyy年MM月");
    const dateRange = formatVariableAsDate(monthStart) + " ~ " + formatVariableAsDate(monthEnd);

    const rows = getFollowUpRows();
    // 按 E列「记录时间」过滤当月
    const monthRows = rows.filter(function(r) {
      return r.recordDateObj && r.recordDateObj >= monthStart && r.recordDateObj <= monthEnd;
    });

    const groups = {};
    monthRows.forEach(function(r) {
      r.assignees.forEach(function(email) {
        if (!authorized[email]) return;
        if (!groups[email]) groups[email] = { total: 0, done: 0, pending: 0 };
        groups[email].total++;
        if (r.isDone) groups[email].done++;
        else groups[email].pending++;
      });
    });

    // 当月所有未完成项（至少一个授权人）
    const pendingItems = monthRows.filter(function(r) {
      return !r.isDone && r.assignees.some(function(a) { return authorized[a]; });
    }).sort(function(a, b) {
      // 按截止日期升序，无截止日期排最后
      if (!a.dueDateObj && !b.dueDateObj) return 0;
      if (!a.dueDateObj) return 1;
      if (!b.dueDateObj) return -1;
      return a.dueDateObj - b.dueDateObj;
    });

    const recipients = Object.keys(groups);
    if (recipients.length === 0) {
      writeLog("sendGmailReports", "跳过", "当月无授权人 INJ SDM 数据", trigger, "范围:" + dateRange);
      return;
    }

    const today = formatVariableAsDate(new Date());
    const subject = "【SmartMeeting Alert】 INJ SDM 跟进完成率月报 - " + monthLabel;
    const html = _buildGmailMergedHtml(groups, pendingItems, today, monthLabel, dateRange);

    try {
      sendMail(recipients.join(","), subject, html, GMAIL_CC);
      const summary = "已发送" + monthLabel + "月报，TO=" + recipients.length + "人，CC=" + GMAIL_CC;
      writeLog("sendGmailReports", "成功", summary, trigger, "TO: " + recipients.join(","));
    } catch (err) {
      writeLog("sendGmailReports", "失败", err.message, trigger, "TO: " + recipients.join(","));
      console.error("发送失败: " + err.message);
    }
  } catch (err) {
    writeLog("sendGmailReports", "失败", err.message, trigger, err.stack || "");
    console.error(err.stack || err.message);
  }
}

/** 合并邮件正文：当月范围 + 团队总览 + 个人完成率排行 + 全员未完成清单 */
function _buildGmailMergedHtml(groups, pendingItems, today, monthLabel, dateRange) {
  let teamTotal = 0, teamDone = 0, teamPending = 0;
  Object.keys(groups).forEach(function(email) {
    teamTotal += groups[email].total;
    teamDone += groups[email].done;
    teamPending += groups[email].pending;
  });
  const teamRate = teamTotal > 0 ? Math.round(teamDone / teamTotal * 100) : 0;
  const teamRateColor = teamRate >= 80 ? "#27ae60" : teamRate >= 60 ? "#e67e22" : "#e74c3c";

  let html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto">';
  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">INJ SDM 跟进完成率月报</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">统计口径：记录时间在 ' + dateRange + '（' + monthLabel + '）</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + today + '</p>';
  html += '</div>';

  // ===== 团队总览 =====
  html += '<div style="padding:24px">';
  html += '<h3 style="color:#E60012;margin-top:0;border-left:4px solid #E60012;padding-left:8px">团队总览</h3>';
  html += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px"><tr>';
  html += _gmailCard("总跟进项", teamTotal, "#2c3e50");
  html += _gmailCard("已完成", teamDone, "#27ae60");
  html += _gmailCard("未完成", teamPending, "#e74c3c");
  html += _gmailCard("完成率", teamRate + "%", teamRateColor);
  html += '</tr></table>';

  // ===== 各人完成率排行（按完成率从高到低）=====
  html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;border-bottom:2px solid #E60012;padding-bottom:8px">个人完成率排行</h3>';
  const rankRows = Object.keys(groups).map(function(email) {
    const g = groups[email];
    const rate = g.total > 0 ? Math.round(g.done / g.total * 100) : 0;
    return { email: email, rate: rate, total: g.total, done: g.done, pending: g.pending };
  }).sort(function(a, b) { return b.rate - a.rate; });

  const rankTable = rankRows.map(function(r) {
    return [r.email.split("@")[0], r.total, r.done, r.pending, r.rate + "%"];
  });
  html += buildHtmlTable(["跟进人", "总数", "已完成", "未完成", "完成率"], rankTable, "#E60012");

  // ===== 当月未完成清单（全员一表） =====
  html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;border-bottom:2px solid #E60012;padding-bottom:8px;margin-top:32px">当月未完成清单 (' + pendingItems.length + '项)</h3>';
  if (pendingItems.length === 0) {
    html += '<p style="color:#27ae60;font-weight:bold">★ 当月跟进项已全部完成</p>';
  } else {
    const pendingTable = pendingItems.map(function(item) {
      const od = diffDays(item.dueDateObj);
      let st;
      if (od === null) st = "-";
      else if (od < 0) st = "逾期" + (-od) + "天";
      else if (od === 0) st = "今日截止";
      else st = "剩余" + od + "天";
      const names = item.assignees.map(function(e) { return e.split("@")[0]; }).join(", ");
      return [names, item.recordDate || "-", item.followUp, item.dueDate || "-", st];
    });
    html += buildHtmlTable(["跟进人", "记录日期", "跟进项", "截止日期", "状态"], pendingTable, "#E60012");
  }

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由 SmartMeeting Alert 系统自动发送，TO 包含当月所有授权跟进人，CC ' + GMAIL_CC + '</p>';
  html += '</div></div></body></html>';
  return html;
}

function _gmailCard(label, value, color) {
  return '<td style="text-align:center;padding:16px;border:1px solid #ecf0f1">' +
    '<div style="font-size:28px;font-weight:bold;color:' + color + '">' + value + '</div>' +
    '<div style="color:#7f8c8d;font-size:12px;margin-top:4px">' + label + '</div></td>';
}
