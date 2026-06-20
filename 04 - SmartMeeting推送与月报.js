// V20260524.1 — SmartMeeting Alert 业务逻辑
// 三个推送入口：sendWechatMorning / sendWechatEvening / sendGmailReports
// 手动触发：直接调用（e=undefined）；定时触发：传入 e 事件对象

/** 08:00 早推送：dueDate≤今天 且未完成 */
function sendWechatMorning(e) {
  _runWechatPush("morning", e ? "定时" : "手动");
}

/** 16:00 晚推送：今日未完成 + 明日到期 */
function sendWechatEvening(e) {
  _runWechatPush("evening", e ? "定时" : "手动");
}

function _runWechatPush(mode, trigger) {
  const fnName = mode === "morning" ? "sendWechatMorning" : "sendWechatEvening";
  try {
    const authorized = getAuthorizedEmails();
    if (authorized === null) {
      writeLog(fnName, "失败", "权限列 SmartMeeting_Alert 不存在", trigger, "请在 userID 表 BJ 列添加");
      return;
    }
    const rows = getFollowUpRows();
    const monthStats = getCurrentMonthStats(rows, authorized);
    // 多跟进人：只要任一跟进人有权限即纳入
    const pending = rows.filter(function(r) {
      return !r.isDone && r.assignees.some(function(a) { return authorized[a]; });
    });

    let content, summary;
    if (mode === "morning") {
      const items = pending.filter(function(r) {
        return r.dueDateObj && r.dueDateObj <= todayDate();
      });
      if (items.length === 0) {
        content = "## 🎉 INJ SDM 今日无需跟进项\n\n" + formatVariableAsDate(new Date()) + " 团队继续加油！";
        summary = "无未完成项，推送清零消息";
      } else {
        content = _buildMorningMd(items, authorized, monthStats);
        summary = "推送 " + items.length + " 项未完成，涉及 " + _countAuthorizedAssignees(items, authorized) + " 人";
      }
    } else {
      const today = todayDate();
      const tomorrow = tomorrowDate();
      const todayItems = pending.filter(function(r) {
        return r.dueDateObj && r.dueDateObj.getTime() === today.getTime();
      });
      const tomorrowItems = pending.filter(function(r) {
        return r.dueDateObj && r.dueDateObj.getTime() === tomorrow.getTime();
      });
      if (todayItems.length === 0 && tomorrowItems.length === 0) {
        content = "## 🎉 INJ SDM 今日全部关闭，明日无新到期\n\n" + formatVariableAsDate(new Date());
        summary = "无未完成项，推送清零消息";
      } else {
        content = _buildEveningMd(todayItems, tomorrowItems, authorized, monthStats);
        summary = "今日未完成 " + todayItems.length + " 项，明日到期 " + tomorrowItems.length + " 项";
      }
    }

    const resp = postWechat(content);
    if (resp.errcode === 0) {
      writeLog(fnName, "成功", summary, trigger, "");
    } else {
      writeLog(fnName, "失败", "企微返回: " + JSON.stringify(resp), trigger, summary);
    }
  } catch (e) {
    writeLog(fnName, "失败", e.message, trigger, e.stack || "");
    console.error(e.stack || e.message);
  }
}

function _buildMorningMd(items, authorized, monthStats) {
  const lines = [];
  lines.push("## 📋 INJ SDM 早报：当天需跟进项 " + formatVariableAsDate(new Date()));
  lines.push("");
  _appendGroupedItems(lines, items, authorized, monthStats);
  lines.push("---");
  lines.push("共 **" + items.length + "** 项需要今日处理");
  return lines.join("\n");
}

function _buildEveningMd(todayItems, tomorrowItems, authorized, monthStats) {
  const lines = [];
  lines.push("## 📋 INJ SDM 晚报 " + formatVariableAsDate(new Date()));
  lines.push("");
  lines.push("### ⚠️ 今日应关闭但未完成 (" + todayItems.length + "项)");
  if (todayItems.length === 0) {
    lines.push('<font color="info">今日全部关闭 ✓</font>');
  } else {
    _appendGroupedItems(lines, todayItems, authorized, monthStats);
  }
  lines.push("");
  lines.push("### 🔔 明日到期 (" + tomorrowItems.length + "项)");
  if (tomorrowItems.length === 0) {
    lines.push('<font color="info">明日无到期项</font>');
  } else {
    _appendGroupedItems(lines, tomorrowItems, authorized, monthStats);
  }
  return lines.join("\n");
}

function _appendGroupedItems(lines, items, authorized, monthStats) {
  // 每个授权的跟进人都看到自己的项；非授权人不出现在分组里
  const grouped = _groupByAssignee(items, authorized);
  Object.keys(grouped).sort().forEach(function(email) {
    const arr = grouped[email];
    const name = email.split("@")[0];
    const monthInfo = _formatMonthStat(monthStats[email]);
    lines.push("**【" + name + "】** " + arr.length + "项未完成 | " + monthInfo);
    arr.slice(0, MAX_ITEMS_PER_PERSON).forEach(function(item) {
      lines.push(_formatItemLine(item));
    });
    if (arr.length > MAX_ITEMS_PER_PERSON) {
      lines.push('><font color="comment">...等共 ' + arr.length + ' 项</font>');
    }
    lines.push("");
  });
}

/** 本月完成率短文：本月 done/total 完成率N%，按比例着色 */
function _formatMonthStat(stat) {
  if (!stat || stat.total === 0) {
    return '<font color="comment">本月 0/0 (无新项)</font>';
  }
  const rate = Math.round(stat.done / stat.total * 100);
  const color = rate >= 80 ? "info" : rate >= 60 ? "comment" : "warning";
  return '本月 ' + stat.done + '/' + stat.total + ' <font color="' + color + '">完成率' + rate + '%</font>';
}

function _formatItemLine(item) {
  const fu = item.followUp.length > 25 ? item.followUp.substring(0, 25) + "…" : item.followUp;
  const od = item.dueDateObj ? diffDays(item.dueDateObj) : null;
  let suffix = "";
  if (od !== null) {
    if (od < 0) suffix = ' <font color="warning">逾期' + (-od) + '天</font>';
    else if (od === 0) suffix = ' <font color="warning">今日截止</font>';
    else suffix = ' (' + item.dueDate + ')';
  }
  return ">" + fu + suffix;
}

/** 按授权的跟进人分组；同一项可出现在多人组里 */
function _groupByAssignee(items, authorized) {
  const groups = {};
  items.forEach(function(item) {
    item.assignees.forEach(function(email) {
      if (authorized && !authorized[email]) return; // 跳过未授权
      if (!groups[email]) groups[email] = [];
      groups[email].push(item);
    });
  });
  return groups;
}

/** 统计涉及多少个授权人 */
function _countAuthorizedAssignees(items, authorized) {
  const set = {};
  items.forEach(function(i) {
    i.assignees.forEach(function(a) {
      if (authorized[a]) set[a] = true;
    });
  });
  return Object.keys(set).length;
}

// ========== Gmail 周报（单封合并，仅当月数据） ==========
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

// ========== 数据预览（仅日志，不推送） ==========
function previewData() {
  const authorized = getAuthorizedEmails();
  console.log("=== 权限白名单 ===");
  if (authorized === null) {
    console.log("❌ 权限列 SmartMeeting_Alert 不存在");
  } else {
    console.log("✓ 共 " + Object.keys(authorized).length + " 人：" + Object.keys(authorized).join(", "));
  }

  const rows = getFollowUpRows();
  const pending = rows.filter(function(r) { return !r.isDone; });
  const filtered = authorized
    ? pending.filter(function(r) { return r.assignees.some(function(a) { return authorized[a]; }); })
    : pending;

  console.log("\n=== INJ SDM 数据统计 ===");
  console.log("总跟进项: " + rows.length);
  console.log("未完成: " + pending.length);
  console.log("权限过滤后未完成: " + filtered.length);

  const today = todayDate();
  const tomorrow = tomorrowDate();
  const todayItems = filtered.filter(function(r) { return r.dueDateObj && r.dueDateObj.getTime() === today.getTime(); });
  const tomorrowItems = filtered.filter(function(r) { return r.dueDateObj && r.dueDateObj.getTime() === tomorrow.getTime(); });
  const overdueItems = filtered.filter(function(r) { return r.dueDateObj && r.dueDateObj < today; });

  console.log("\n=== 时段分布 ===");
  console.log("📅 今日截止未完成: " + todayItems.length + " 项");
  console.log("📅 明日到期: " + tomorrowItems.length + " 项");
  console.log("⚠️  历史逾期未完成: " + overdueItems.length + " 项");

  console.log("\n=== 未完成前 10 条 ===");
  filtered.slice(0, 10).forEach(function(r, idx) {
    console.log((idx + 1) + ". [" + r.date + "] " + r.assignees.join("|") + " | " + r.followUp.substring(0, 40) + " | 截止:" + r.dueDate + " | " + r.status);
  });
}

// ========== 跟进人对账诊断 ==========
/** 列出所有 INJ SDM 跟进人，区分白名单内/外 */
function diagnoseAssignees() {
  const authorized = getAuthorizedEmails();
  if (authorized === null) {
    console.log("❌ 权限列不存在，无法诊断");
    return;
  }
  const rows = getFollowUpRows();
  const allAssignees = {};
  rows.forEach(function(r) {
    r.assignees.forEach(function(a) {
      if (!allAssignees[a]) allAssignees[a] = { total: 0, pending: 0 };
      allAssignees[a].total++;
      if (!r.isDone) allAssignees[a].pending++;
    });
  });

  const inWhitelist = [];
  const notInWhitelist = [];
  Object.keys(allAssignees).sort().forEach(function(email) {
    const s = allAssignees[email];
    const line = email + "  总=" + s.total + " 未完成=" + s.pending;
    if (authorized[email]) inWhitelist.push("✅ " + line);
    else notInWhitelist.push("❌ " + line);
  });

  console.log("=== 在白名单（会被处理） " + inWhitelist.length + " 人 ===");
  inWhitelist.forEach(function(l) { console.log(l); });
  console.log("\n=== 不在白名单（被跳过） " + notInWhitelist.length + " 人 ===");
  notInWhitelist.forEach(function(l) { console.log(l); });
}
