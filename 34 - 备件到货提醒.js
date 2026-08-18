// V20260817.01 — 备件到货提醒（每天 08:30 提醒未领出备件的申请人，直至领出自动停止）
// 入口：sendSpareArrivalReminder（定时/手动）；测试入口：testSpareArrivalReminder（仅发 kelland_zhao@colpal.com）
// 数据源：备件项目采购未领出报告（1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU）→ 未领出备件报告
// 逻辑：报告 tab 的 FILTER 公式只含未领用数量>0 的记录 → 每天提醒，领出后记录消失自动停止
// 收件人：申请人（userID 姓名→J列GMail）TO + 直线上级（BI列）CC + kelland_zhao CC
// 兜底：姓名匹配不到 → 汇总一封发 kelland_zhao@colpal.com + Log 警告
// 邮件：遵循 docs/邮件UI规范.md 模式 C

// ========== 配置常量 ==========
const SAR_CONFIG = {
  REPORT_SPREADSHEET_ID: "1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU",
  REPORT_SHEET_NAME: "未领出备件报告",
  SENDER_NAME: "备件到货提醒系统",
  ADMIN_EMAIL: "kelland_zhao@colpal.com",
};

const SAR_FUNC_NAME = "sendSpareArrivalReminder";
const SAR_EMAIL_COL = 9;      // userID J列 GMail（0-indexed）
const SAR_SUPERVISOR_COL = 60; // userID BI列 直线上级邮箱（0-indexed）

// ========== 姓名预处理（用于 userID 匹配，同 08 号模块模式）==========
function _sar_normalizeName(name) {
  let s = String(name || '').trim();
  s = s.replace(/-[A-Za-z0-9]+$/, ''); // 去掉 "-Z" 类后缀
  s = s.replace(/_[A-Za-z0-9_]+$/, ''); // 去掉 "_QC" 类后缀
  return s.trim().toLowerCase();
}

// ========== 读取未领出备件报告 ==========
function _sar_readReportRows(trigger) {
  const ss = SpreadsheetApp.openById(SAR_CONFIG.REPORT_SPREADSHEET_ID);
  const ws = ss.getSheetByName(SAR_CONFIG.REPORT_SHEET_NAME);
  if (!ws) {
    writeLog(SAR_FUNC_NAME, "失败", "找不到工作表: " + SAR_CONFIG.REPORT_SHEET_NAME, trigger, "");
    return [];
  }
  const lastRow = ws.getLastRow();
  if (lastRow < 2) {
    writeLog(SAR_FUNC_NAME, "成功", "报告无数据行，跳过", trigger, "");
    return [];
  }
  // 列：A备件# B名称 C-PO D申请人 E采购原因 F采购数量 G到货日期 H-Stock I-Safety Stock J未领用数量 K在备件房天数 L超期提醒 M备注 N单价 O合价
  const data = ws.getRange(2, 1, lastRow - 1, 15).getDisplayValues();
  const rows = [];
  data.forEach(function(r) {
    const code = String(r[0] || "").trim();
    const applicant = String(r[3] || "").trim();
    const qty = Number(String(r[9] || "0").replace(/,/g, ""));
    if (!code || !applicant || !(qty > 0)) return; // 防御性过滤（报告公式已保证，双保险）
    rows.push({
      code: code,
      name: String(r[1] || "").trim(),
      po: String(r[2] || "").trim(),
      applicant: applicant,
      reason: String(r[4] || "").trim(),
      poQty: String(r[5] || "").trim(),
      arrivalDate: String(r[6] || "").trim(),
      daysInRoom: String(r[10] || "").trim(),
      overdueNote: String(r[11] || "").trim(),
      unitPrice: String(r[13] || "").trim(),
      totalPrice: String(r[14] || "").trim(),
      unpickedQty: String(r[9] || "").trim(),
    });
  });
  writeLog(SAR_FUNC_NAME, "成功", "读取到 " + rows.length + " 条未领出记录", trigger, "");
  return rows;
}

// ========== 构建 userID 姓名→邮箱映射 ==========
function _sar_buildUserMap(trigger) {
  const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
  if (!sheet) {
    writeLog(SAR_FUNC_NAME, "失败", "找不到权限表: " + PERMISSION_SHEET_NAME, trigger, "");
    return {};
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    writeLog(SAR_FUNC_NAME, "失败", "权限表数据不足（lastRow=" + lastRow + "）", trigger, "");
    return {};
  }
  // 读取 A~BI（61列），从第3行开始（跳过2行表头），同 08 号模块模式
  const data = sheet.getRange(3, 1, lastRow - 2, 61).getValues();
  const map = {};
  data.forEach(function(row) {
    const key = _sar_normalizeName(row[1]); // B列 姓名
    if (!key) return;
    const email = String(row[SAR_EMAIL_COL] || "").trim().toLowerCase();
    const supervisorEmail = String(row[SAR_SUPERVISOR_COL] || "").trim().toLowerCase();
    if (!map[key]) {
      map[key] = { email: email, supervisorEmail: supervisorEmail };
    } else if (!map[key].email && email) {
      map[key].email = email; // 首条无邮箱时补
    }
  });
  writeLog(SAR_FUNC_NAME, "成功", "userID 映射构建完成，共 " + Object.keys(map).length + " 人", trigger, "");
  return map;
}

// ========== 临时调试入口（Task 4 删除）==========
function _sar_debugRead() {
  const rows = _sar_readReportRows("手动");
  console.log("共 " + rows.length + " 条未领出记录");
  if (rows.length > 0) {
    console.log(JSON.stringify(rows[0]));
  }
  const userMap = _sar_buildUserMap("手动");
  console.log("userID 映射人数: " + Object.keys(userMap).length);
  const sampleNames = {};
  rows.forEach(function(r) {
    const key = _sar_normalizeName(r.applicant);
    if (!sampleNames[key]) sampleNames[key] = { applicant: r.applicant, u: userMap[key] };
  });
  Object.keys(sampleNames).forEach(function(key) {
    console.log(key + " → " + JSON.stringify(sampleNames[key].u));
  });
}
