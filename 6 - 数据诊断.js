// V20260524.1 — SmartMeeting Alert 数据预览与诊断
// 手动工具，不参与定时推送

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
