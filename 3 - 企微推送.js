// V20260524.1 — SmartMeeting Alert 企微早晚推送
// 入口：sendWechatMorning / sendWechatEvening
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
