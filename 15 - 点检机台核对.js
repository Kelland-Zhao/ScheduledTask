// V20260605.1 — 点检机台核对
// 入口：checkPointCheckMachines（每日 08:25 定时 or 手动）
// 逻辑：比对 MachineList(工序=INJ) 与 Workcenter(C列≠闲置) 的机台差异，
//       差异1（点检有/计划账无）→ MachineList 标黄 + 邮件，
//       差异2（计划账有/点检无）→ 仅邮件，机型=6AX 豁免

// ========== 数据源配置 ==========
const _pc_ID_POINTCHECK = "1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY";
const _pc_SHEET_MACHINELIST = "MachineList";
const _pc_FILTER_PROCESS = "INJ";         // 工序过滤条件

const _pc_ID_PLAN = "12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM";
const _pc_SHEET_WORKCENTER = "Workcenter";
const _pc_EXCLUDE_KEYWORD = "闲置";       // Workcenter C列排除关键字

const _pc_ID_PERMISSION = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const _pc_SHEET_USERID = "userID";
const _pc_PERM_PROCESS_COL = 14;          // O列(0-indexed): 工序
const _pc_PERM_ROLE_COL = 15;             // P列: 职位
const _pc_PERM_EMAIL_COL = 9;             // J列: GMail
const _pc_PERM_PROCESS_VAL = "INJ";
const _pc_PERM_ROLE_VAL = "S&C";

// ========== 主入口 ==========
function checkPointCheckMachines(e) {
  const trigger = e ? "定时" : "手动";
  try {
    console.log("开始执行点检机台核对...");

    // 1. 读取 MachineList，过滤工序=INJ
    const ssPC = SpreadsheetApp.openById(_pc_ID_POINTCHECK);
    const wsML = ssPC.getSheetByName(_pc_SHEET_MACHINELIST);
    const dataML = wsML.getDataRange().getValues();

    if (dataML.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 为空或只有表头", trigger, "");
      return;
    }

    const headerML = dataML[0];
    const injRows = [];              // { rowIndex(1-based), machineNo, rowData[] }
    const injMachineNos = new Set();

    for (let i = 1; i < dataML.length; i++) {
      if (String(dataML[i][0] || "").trim() === _pc_FILTER_PROCESS) {
        const machineNo = String(dataML[i][3] || "").trim();
        injRows.push({ rowIndex: i + 1, machineNo: machineNo, rowData: dataML[i] });
        if (machineNo) injMachineNos.add(machineNo);
      }
    }

    console.log("MachineList INJ 工序行数: " + injRows.length);

    if (injRows.length === 0) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 无 INJ 工序数据", trigger, "");
      return;
    }

    // 2. 读取 Workcenter，过滤 C列≠闲置
    const ssPlan = SpreadsheetApp.openById(_pc_ID_PLAN);
    const wsWC = ssPlan.getSheetByName(_pc_SHEET_WORKCENTER);
    const dataWC = wsWC.getDataRange().getValues();

    if (dataWC.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 为空", trigger, "");
      return;
    }

    const wcMap = {};  // Workcenter → { machineType, finalMachineType, ... }
    for (let i = 1; i < dataWC.length; i++) {
      const colC = String(dataWC[i][2] || "").trim();
      if (colC === _pc_EXCLUDE_KEYWORD) continue;

      const wc = String(dataWC[i][0] || "").trim();
      if (wc) {
        wcMap[wc] = {
          machineType: String(dataWC[i][1] || "").trim(),       // B列 Machine Type
          machineModel: String(dataWC[i][2] || "").trim(),       // C列 机器性能
        };
      }
    }

    console.log("Workcenter 有效行数(排除闲置): " + Object.keys(wcMap).length);

    // 3. 计算差异
    const wcSet = new Set(Object.keys(wcMap));
    const type1 = [];  // 点检有/计划账无
    const type2 = [];  // 计划账有/点检无

    injRows.forEach(function (r) {
      if (r.machineNo && !wcSet.has(r.machineNo)) {
        type1.push(r);
      }
    });

    wcSet.forEach(function (wc) {
      if (!injMachineNos.has(wc)) {
        // 豁免：机型=6AX 的机台不纳入差异类型2
        if (wcMap[wc].machineModel === "6AX") return;
        type2.push({ workcenter: wc, info: wcMap[wc] });
      }
    });

    // 排序
    type1.sort(function (a, b) { return a.machineNo.localeCompare(b.machineNo); });
    type2.sort(function (a, b) { return a.workcenter.localeCompare(b.workcenter); });

    console.log("差异1(点检有/计划账无): " + type1.length + " 台");
    console.log("差异2(计划账有/点检无): " + type2.length + " 台");

    // 4. MachineList 标黄（差异类型1）
    _pc_updateHighlights(wsML, dataML, injRows, type1);

    // 5. 获取收件人
    const recipients = _pc_getRecipients();

    // 6. 发送邮件
    if (recipients.length > 0) {
      const today = formatVariableAsDate(new Date());
      const subject = "[点检核对] 注塑机台差异报告 " + today;
      const html = _pc_buildEmailHtml(type1, type2, today);

      try {
        _pc_sendMail(recipients.join(","), subject, html);
        const summary = "差异1=" + type1.length + "台, 差异2=" + type2.length + "台, TO=" + recipients.length + "人";
        writeLog("checkPointCheckMachines", "成功", summary, trigger, "TO: " + recipients.join(","));
        console.log("邮件发送成功: " + summary);
      } catch (err) {
        writeLog("checkPointCheckMachines", "失败", err.message, trigger, "TO: " + recipients.join(","));
        console.error("发送失败: " + err.message);
      }
    } else {
      writeLog("checkPointCheckMachines", "跳过", "无匹配收件人(O=INJ,P=S&C)", trigger, "");
      console.warn("未找到匹配收件人");
    }

    console.log("点检机台核对执行完毕");

  } catch (err) {
    console.error(err.stack || err.message);
    try { writeLog("checkPointCheckMachines", "失败", err.message, trigger, err.stack || ""); } catch (e2) { }
  }
}

// ========== MachineList 标黄 ==========
/** 批量更新 MachineList 背景色：差异类型1 标黄，其余 INJ 行清除背景 */
function _pc_updateHighlights(ws, dataML, injRows, type1) {
  const lastRow = ws.getLastRow();
  const bgData = ws.getRange(1, 1, lastRow, 5).getBackgrounds();  // A~E列

  const type1IdxSet = new Set(type1.map(function (r) { return r.rowIndex - 1; })); // 0-indexed

  // 遍历所有数据行，只修改 INJ 工序行的背景
  for (let i = 1; i <= lastRow - 1 && i < dataML.length; i++) {
    if (String(dataML[i][0] || "").trim() === _pc_FILTER_PROCESS) {
      if (type1IdxSet.has(i)) {
        bgData[i] = ["#FFFF00", "#FFFF00", "#FFFF00", "#FFFF00", "#FFFF00"];
      } else {
        bgData[i] = [null, null, null, null, null];
      }
    }
  }

  ws.getRange(1, 1, lastRow, 5).setBackgrounds(bgData);
  console.log("MachineList 标黄完成: " + type1.length + " 行");
}

// ========== 收件人 ==========
/** 从 userID 表读取 O列=INJ 且 P列=S&C 的邮箱 */
function _pc_getRecipients() {
  try {
    const sheet = SpreadsheetApp.openById(_pc_ID_PERMISSION).getSheetByName(_pc_SHEET_USERID);
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    const data = sheet.getRange(1, 1, lastRow, Math.max(_pc_PERM_ROLE_COL + 1, _pc_PERM_EMAIL_COL + 1)).getValues();
    const recipients = [];

    for (let i = 2; i < data.length; i++) {
      const process = String(data[i][_pc_PERM_PROCESS_COL] || "").trim();
      const role = String(data[i][_pc_PERM_ROLE_COL] || "").trim();
      const email = String(data[i][_pc_PERM_EMAIL_COL] || "").trim();

      if (process === _pc_PERM_PROCESS_VAL && role === _pc_PERM_ROLE_VAL && email) {
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
function _pc_sendMail(to, subject, htmlBody) {
  GmailApp.sendEmail(to, subject, "", {
    htmlBody: htmlBody,
    name: "PointCheck Alert"
  });
}

// ========== 邮件 HTML ==========
function _pc_buildEmailHtml(type1, type2, today) {
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto">';

  // 红色标题栏
  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">注塑机台差异报告</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">比对范围：MachineList(工序=INJ) ↔ Workcenter(排除闲置)</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + today + '</p>';
  html += '</div>';

  html += '<div style="padding:24px">';

  // ===== 总览 =====
  html += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px"><tr>';
  html += _pc_card("差异类型1<br>点检有/计划账无", type1.length, "#e67e22");
  html += _pc_card("差异类型2<br>计划账有/点检无", type2.length, "#e74c3c");
  html += '</tr></table>';

  // ===== 无差异 =====
  if (type1.length === 0 && type2.length === 0) {
    html += '<p style="color:#27ae60;font-weight:bold;font-size:16px">★ 点检机台与计划账机台完全一致，无差异。</p>';
  }

  // ===== 差异类型1 =====
  if (type1.length > 0) {
    html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px">差异类型1：点检有 / 计划账无 (' + type1.length + '台)</h3>';
    html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">⚠ 请确认机台是否存在，更新点检机台主数据</p>';

    var t1Headers = ["工序", "车间", "机型", "机台号", "点检人"];
    var t1Rows = type1.map(function (r) {
      return [r.rowData[0] || "-", r.rowData[1] || "-", r.rowData[2] || "-", r.rowData[3] || "-", r.rowData[4] || "-"];
    });
    html += buildHtmlTable(t1Headers, t1Rows, "#E60012");
  }

  // ===== 差异类型2 =====
  if (type2.length > 0) {
    html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:32px">差异类型2：计划账有 / 点检无 (' + type2.length + '台)</h3>';
    html += '<p style="color:#e74c3c;font-weight:bold;margin-bottom:8px">⚠ 这些机台在计划账上，但在点检机台主数据中缺失，需要更新点检机台主数据</p>';

    var t2Headers = ["Workcenter", "Machine Type", "机型"];
    var t2Rows = type2.map(function (r) {
      return [r.workcenter, r.info.machineType || "-", r.info.machineModel || "-"];
    });
    html += buildHtmlTable(t2Headers, t2Rows, "#E60012");
  }

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由 PointCheck Alert 系统自动发送</p>';
  html += '</div></div></body></html>';
  return html;
}

function _pc_card(label, value, color) {
  return '<td style="text-align:center;padding:16px;border:1px solid #ecf0f1;width:50%">' +
    '<div style="font-size:32px;font-weight:bold;color:' + color + '">' + value + '</div>' +
    '<div style="color:#7f8c8d;font-size:13px;margin-top:6px;line-height:1.5">' + label + '</div></td>';
}
