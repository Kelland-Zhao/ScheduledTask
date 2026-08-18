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

  const firstKey = Object.keys(sampleNames)[0];
  if (firstKey) {
    const items = rows.filter(function(r) {
      return _sar_normalizeName(r.applicant) === firstKey;
    });
    const body = _sar_buildEmailBody(sampleNames[firstKey].applicant, items, "debug");
    console.log("正文长度: " + body.length);
    console.log(body.substring(0, 500));
  }
}

// ========== 邮件正文（模式 C：白色卡片+色条 + 渐变表头卡片表格）==========
function _sar_buildEmailBody(applicant, items, email) {
  var tableRows = "";
  items.forEach(function(it) {
    var daysCell;
    if (String(it.daysInRoom).trim() === "0") {
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:#27ae60;color:white;font-weight:600;font-size:12px;">今日到货<br><span style="font-size:10px;opacity:0.9;">Arrived Today</span></span>';
    } else {
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:linear-gradient(135deg,#f39c12,#e67e22);color:white;font-weight:600;font-size:12px;">' + escapeHtml(it.daysInRoom) + ' 天<br><span style="font-size:10px;opacity:0.9;">Days</span></span>';
    }
    var overdueCell = it.overdueNote
      ? '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:linear-gradient(135deg,#e74c3c,#c0392b);color:white;font-weight:600;font-size:12px;">' + escapeHtml(it.overdueNote) + '</span>'
      : '—';
    var reasonCell = it.reason === "紧急采购"
      ? '<span style="display:inline-block;padding:4px 8px;border-radius:12px;background:#f39c12;color:white;font-size:12px;">紧急采购</span>'
      : escapeHtml(it.reason || "—");
    var priceCell = it.unitPrice ? (escapeHtml(it.unitPrice) + " / " + escapeHtml(it.totalPrice || "—")) : "—";
    tableRows += '<tr style="background-color:#ffffff;">' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + escapeHtml(it.code) + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:260px;word-wrap:break-word;">' + escapeHtml(it.name || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.po || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + reasonCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.arrivalDate || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + daysCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + overdueCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.unpickedQty || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + priceCell + '</td>' +
      '</tr>';
  });

  var htmlBody =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">' +

    // 模式 C：白色卡片 + 左侧红色色条
    '<div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;border-left:5px solid #E60012;">' +
    '<h2 style="color:#d32f2f;margin:0 0 8px 0;">【备件到货提醒】您有备件待领用</h2>' +
    '<p style="color:#c62828;margin:0;font-size:14px;">Spare Parts Pickup Reminder — ' + items.length + ' 件待领用</p>' +
    '</div>' +

    // 正文卡片
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;">' +
    '<p style="font-size:15px;color:#34495e;line-height:1.8;">' +
    escapeHtml(applicant) + '，您好！以下备件已到货/仍在备件房，请尽快到备件房领用：<br>' +
    '<span style="font-size:13px;opacity:0.8;">Hello! The following spare parts have arrived / are waiting in the spare parts room. Please pick them up as soon as possible.</span>' +
    '</p>' +

    // 渐变表头卡片表格
    '<div style="overflow-x:auto;margin:24px 0;">' +
    '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.1);font-size:13px;">' +
    '<thead><tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">' +
    '<th style="padding:10px;text-align:left;font-weight:600;">备件编码<br><span style="font-size:11px;opacity:0.85;">Code</span></th>' +
    '<th style="padding:10px;text-align:left;font-weight:600;">备件名称<br><span style="font-size:11px;opacity:0.85;">Name</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">PO</th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">采购原因<br><span style="font-size:11px;opacity:0.85;">Reason</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">到货日期<br><span style="font-size:11px;opacity:0.85;">Arrival</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">在备件房<br><span style="font-size:11px;opacity:0.85;">Days</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">超期提醒<br><span style="font-size:11px;opacity:0.85;">Overdue</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">未领数量<br><span style="font-size:11px;opacity:0.85;">Qty</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">单价/合价<br><span style="font-size:11px;opacity:0.85;">Price</span></th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table></div>' +
    '</div>' +

    // 尾部 模式 B：双语
    '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;">' +
    '<p style="margin-bottom:10px;">请及时到备件房领用相关备件。<br>' +
    '<span style="font-size:0.9em;opacity:0.8;">Please pick up the spare parts at the spare parts room promptly.</span></p>' +
    '<p style="margin:0;font-style:italic;">此邮件由备件到货提醒系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.8;">This email is automatically sent by the Spare Parts Arrival Reminder System, please do not reply.</span></p>' +
    '</div></div></body></html>';

  return htmlBody;
}

// ========== 未匹配人员汇总邮件正文 ==========
function _sar_buildUnmatchedEmailBody(list) {
  var tableRows = "";
  list.forEach(function(g) {
    g.items.forEach(function(it) {
      tableRows += '<tr style="background-color:#ffffff;">' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + escapeHtml(it.code) + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:260px;word-wrap:break-word;">' + escapeHtml(it.name || "—") + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.po || "—") + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.applicant) + '</td>' +
        '</tr>';
    });
  });

  var htmlBody =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">' +
    '<div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;border-left:5px solid #E60012;">' +
    '<h2 style="color:#d32f2f;margin:0 0 8px 0;">【备件到货提醒】未匹配人员汇总</h2>' +
    '<p style="color:#c62828;margin:0;font-size:14px;">以下申请人在 userID 表中未找到，无法发送提醒，请人工转告</p>' +
    '</div>' +
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">' +
    '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">' +
    '<th style="padding:10px;text-align:left;">备件编码</th><th style="padding:10px;text-align:left;">备件名称</th>' +
    '<th style="padding:10px;text-align:center;">PO</th><th style="padding:10px;text-align:center;">申请人</th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table></div></div>' +
    '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;margin-top:20px;">' +
    '<p style="margin:0;font-style:italic;">此邮件由备件到货提醒系统自动发送，请勿回复。</p></div>' +
    '</div></body></html>';

  return htmlBody;
}
