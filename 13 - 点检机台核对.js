// V20260930.01 — 点检机台核对
// 入口：checkPointCheckMachines（每日 08:25 定时 or 手动）
// 逻辑：比对 MachineList(工序=INJ) 与 Workcenter(无需检查Y/N≠Y) 的机台差异，
//       差异1（点检有/计划账无）→ MachineList 标黄 + 邮件，
//       差异2（计划账有/点检无）→ 仅邮件，Final Machine Type=6AX/DP/HS 豁免
// 2026-09-30：Workcenter 表由 11 列改为 19 列，改为按表头名定位列

// ========== 数据源配置 ==========
const _pc_ID_POINTCHECK = "1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY";
const _pc_SHEET_MACHINELIST = "MachineList";
const _pc_FILTER_PROCESS = "INJ";         // 工序过滤条件

const _pc_ID_PLAN = "12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM";
const _pc_SHEET_WORKCENTER = "Workcenter";

// ========== 纳入集判定的表头名（按名定位，不硬编码列号）==========
const _pc_HEADER_WORKCENTER = "Workcenter";
const _pc_HEADER_MACHINE_TYPE = "Machine Type";
const _pc_HEADER_FINAL_TYPE = "Final Machine Type";
const _pc_HEADER_FLAG = "无需检查Y/N";
const _pc_FLAG_YES = "Y";
const _pc_EXEMPT_PLASMA = "Plasma";       // 点检侧独有、表11 完全没有的机型，不纳入删除

// 安全阀：单次删除超过这个机台数就只报告不删。稳态日差异应为 0-3 台，
// 超阈值几乎必定是上游异常（表头错位、人工误标）而非真有批量变动
const _pc_DELETE_LIMIT = 10;
const _pc_HIGHLIGHT = "#FFFF00";          // 待补全行的标黄颜色（写入值，大小写皆可）

/**
 * 归一化背景色，供「目标色 ≠ 当前色」比较使用。
 * **必须归一**：`getBackgrounds()` 返回**小写**十六进制，且未设置/已清除的单元格
 * 返回 `"#ffffff"` 而不是 null/空串。不归一的话比较恒不相等 ——
 * 每次运行都会把整段 INJ 行重写一遍，「只写变色区间」的约束会失效。
 * @param {*} value 原始背景色值
 * @returns {string|null} 归一后的小写色值；无色返回 null
 */
function _pc_normalizeBg(value) {
  const s = String(value === undefined || value === null ? "" : value).trim().toLowerCase();
  if (s === "" || s === "#ffffff" || s === "white") return null;
  return s;
}

const _pc_ID_PERMISSION = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const _pc_SHEET_USERID = "userID";
const _pc_PERM_PROCESS_COL = 14;          // O列(0-indexed): 工序
const _pc_PERM_ROLE_COL = 15;             // P列: 职位
const _pc_PERM_EMAIL_COL = 9;             // J列: GMail
const _pc_PERM_PROCESS_VAL = "INJ";
const _pc_PERM_ROLE_VAL = "S&C";

// ========== Workcenter 读取（按表头名定位列） ==========
function _pc_headerIndex(headerRow) {
  const idx = {};
  for (let i = 0; i < headerRow.length; i++) {
    const name = String(headerRow[i] || "").trim();
    if (name && idx[name] === undefined) idx[name] = i; // 重名取第一列
  }
  return idx;
}

/** 机台号比对归一：去空格 + 转大写。非字符串输入安全返回 "" */
function _pc_normalizeMachineNo(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim().toUpperCase();
}

/**
 * 解析 Workcenter，产出纳入集与数据质量问题
 * 纳入集 = 全部行 − (「无需检查Y/N」= Y)；判据按表头名定位列
 * @param {Array<Array>} dataWC getDataRange().getValues() 全量（含表头行）
 * @returns {{included: Object, flagged: Object, missing: Array<string>, duplicates: Array<string>, badFlags: Array<string>}}
 *   missing 非空时调用方应跳过，不要用空 included 继续
 */
function _pc_buildIncludedSet(dataWC) {
  const required = [_pc_HEADER_WORKCENTER, _pc_HEADER_MACHINE_TYPE, _pc_HEADER_FINAL_TYPE, _pc_HEADER_FLAG];
  const cols = _pc_headerIndex(dataWC[0] || []);
  const missing = required.filter(function (n) { return cols[n] === undefined; });
  if (missing.length > 0) return { included: {}, flagged: {}, missing: missing, duplicates: [], badFlags: [] };

  const included = {};
  const flagged = {};
  const duplicates = [];
  const badFlags = [];

  for (let i = 1; i < dataWC.length; i++) {
    const machineNo = _pc_normalizeMachineNo(dataWC[i][cols[_pc_HEADER_WORKCENTER]]);
    if (!machineNo) continue;

    const rawFlag = String(dataWC[i][cols[_pc_HEADER_FLAG]] || "").trim();

    // 非空且不是 Y → 标了却没生效，必须告警而不是静默
    if (rawFlag !== "" && rawFlag !== _pc_FLAG_YES && badFlags.indexOf(rawFlag) < 0) {
      badFlags.push(rawFlag);
    }

    if (rawFlag === _pc_FLAG_YES) {
      flagged[machineNo] = true;
      delete included[machineNo];   // Y 优先：同号任一行标 Y，整台机就不点检
      continue;
    }

    // 已被标 Y 的号优先于其他行，不能因为后面还有一行没标 Y 就把它捞回纳入集
    if (Object.prototype.hasOwnProperty.call(flagged, machineNo)) {
      duplicates.push(machineNo);
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(included, machineNo)) {
      duplicates.push(machineNo);   // 重复键取第一条（沿用模块 12 做法）
      continue;
    }

    included[machineNo] = {
      machineType: String(dataWC[i][cols[_pc_HEADER_MACHINE_TYPE]] || "").trim(),
      machineModel: String(dataWC[i][cols[_pc_HEADER_FINAL_TYPE]] || "").trim(),
    };
  }

  return { included: included, flagged: flagged, missing: [], duplicates: duplicates, badFlags: badFlags };
}

/**
 * 由机台号推导车间：从左数第 2 个字符
 * 实测 Workcenter 全部 326 台机台第二位只有 0/1/2；其他值一律留空交人工（不猜）
 * @param {string} machineNo
 * @returns {string} "TB1" | "TB2" | ""
 */
function _pc_deriveWorkshop(machineNo) {
  const c = String(machineNo || "").charAt(1);
  if (c === "0" || c === "1") return "TB1";
  if (c === "2") return "TB2";
  return "";
}

/**
 * 算出本次同步计划（纯函数，不做任何 IO）
 * 比对键 = 机台号；一台机在 MachineList 中可能占多行，整组一起进、一起出
 * @param {Array<{rowIndex: number, machineNo: string, rowData: Array}>} injRows MachineList 中 工序=INJ 的行
 * @param {{included: Object, flagged: Object}} built _pc_buildIncludedSet 的结果
 * @returns {{append: Array, toDelete: Array, plasmaKept: Array}}
 */
function _pc_computeSyncPlan(injRows, built) {
  // 防御：MachineList 的 INJ 段读空时，绝不能把整个纳入集当成「待追加」灌进去
  if (!injRows || injRows.length === 0) return { append: [], toDelete: [], plasmaKept: [] };

  const included = built.included || {};
  const flagged = built.flagged || {};

  // 按机台号聚拢 MachineList 的行（保留首次出现顺序，供追加位置与稳定性判断）
  const mlGroups = {};
  injRows.forEach(function (r) {
    const no = _pc_normalizeMachineNo(r.machineNo);
    if (!no) return;                       // 机台号为空的行不参与比对
    if (!mlGroups[no]) mlGroups[no] = { rows: [], rowData: [] };
    mlGroups[no].rows.push(r.rowIndex);
    mlGroups[no].rowData.push(r.rowData);
  });

  // 二次防御：injRows 非空但机台号列全读空（列错位/读表异常）→ mlGroups 为空，
  // 不守的话整个纳入集照样会被当成「待追加」批量写进主数据。同一灾难、不同扳机
  if (Object.keys(mlGroups).length === 0) return { append: [], toDelete: [], plasmaKept: [] };

  const toDelete = [];
  const plasmaKept = [];

  Object.keys(mlGroups).forEach(function (no) {
    const g = mlGroups[no];
    if (Object.prototype.hasOwnProperty.call(included, no)) return;   // 两侧都有 → 不动

    // 表11 完全没有 / 已被标Y → 删；但 Plasma 机型豁免
    const isPlasma = g.rowData.some(function (d) {
      return String(d[2] || "").trim() === _pc_EXEMPT_PLASMA;
    });
    if (isPlasma) {
      plasmaKept.push({ machineNo: no, rowIndexes: g.rows });
      return;
    }

    toDelete.push({
      machineNo: no,
      rowIndexes: g.rows,
      reason: Object.prototype.hasOwnProperty.call(flagged, no) ? "FLAG" : "ABSENT",
      snapshot: g.rowData,
    });
  });

  const append = [];
  Object.keys(included).forEach(function (no) {
    if (Object.prototype.hasOwnProperty.call(mlGroups, no)) return;
    append.push({ machineNo: no, workshop: _pc_deriveWorkshop(no) });
  });

  return { append: append, toDelete: toDelete, plasmaKept: plasmaKept };
}

/**
 * 安全阀：是否执行删除。追加不设阀（可逆），只拦删除（不可逆）
 * @param {number} deleteCount 拟删除的**机台数**（不是行数）
 * @param {boolean} isManual 手动运行 = 人工放行，绕过阀
 * @returns {{ok: boolean, reason: string}}
 */
function _pc_shouldDelete(deleteCount, isManual) {
  if (isManual) return { ok: true, reason: "" };
  if (deleteCount > _pc_DELETE_LIMIT) {
    return {
      ok: false,
      reason: "拟删除 " + deleteCount + " 台，超过阈值 " + _pc_DELETE_LIMIT + "，本次仅报告未执行删除",
    };
  }
  return { ok: true, reason: "" };
}

/**
 * 待补全判定：机型(C列,index2) 或 点检人(E列,index4) 为空
 * 追加行天然为空 → 必然待补全，靠标黄+邮件暴露直到人工补齐
 * @param {Array} rowData
 * @returns {boolean}
 */
function _pc_isIncomplete(rowData) {
  const model = String(rowData[2] || "").trim();
  const checker = String(rowData[4] || "").trim();
  return model === "" || checker === "";
}

/**
 * 算出需要改色的区间（纯函数）。只输出「目标色 ≠ 当前色」的连续区间，
 * 颜色没变的行一概不产生动作 —— 非 INJ 行永远不需要改色，因此永远不会被写入
 * @param {Array<Array>} bgData 全表背景色（0-based，含表头行）
 * @param {Set<number>} injRowIdx 需管理的 INJ 行下标（0-based）
 * @param {Set<number>} incompleteIdx 待补全行下标（0-based）
 * @returns {Array<{start: number, count: number, color: string|null}>} start 为 0-based
 */
function _pc_computeColorActions(bgData, injRowIdx, incompleteIdx) {
  const actions = [];
  let cur = null;

  for (let i = 0; i < bgData.length; i++) {
    // 非 INJ 行不受管理：不读它的颜色、不为它产生动作，且**必须打断正在累积的区间**
    // （表内 INJ 段被 PK/TF 行夹断，区间一旦跨过去就会把那些行一起重涂）
    if (!injRowIdx.has(i)) {
      if (cur) { actions.push(cur); cur = null; }
      continue;
    }

    const target = incompleteIdx.has(i) ? _pc_HIGHLIGHT : null;
    // 只读 A 列颜色：本模块写入时整行 A~E 涂同一色，故 A 列即代表整行。
    // 若将来出现「只涂部分列」的写入，这里会静默失效 —— 届时须改为逐列比较
    const current = _pc_normalizeBg((bgData[i] || [])[0]);

    if (_pc_normalizeBg(target) === current) {
      if (cur) { actions.push(cur); cur = null; }
      continue;
    }

    if (cur && cur.color === target) {
      cur.count++;
    } else {
      if (cur) actions.push(cur);
      cur = { start: i, count: 1, color: target };
    }
  }
  if (cur) actions.push(cur);
  return actions;
}

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

    // 2. 读取 Workcenter，纳入集 = 全部行 − (无需检查Y/N=Y)
    const ssPlan = SpreadsheetApp.openById(_pc_ID_PLAN);
    const wsWC = ssPlan.getSheetByName(_pc_SHEET_WORKCENTER);
    const dataWC = wsWC.getDataRange().getValues();

    if (dataWC.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 为空", trigger, "");
      return;
    }

    const built = _pc_buildIncludedSet(dataWC);
    if (built.missing.length > 0) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 表头缺少字段: " + built.missing.join(", "), trigger, "");
      return;
    }
    const wcMap = built.included;  // Workcenter → { machineType, machineModel }

    console.log("Workcenter 纳入集行数(排除无需检查): " + Object.keys(wcMap).length);
    if (built.badFlags.length > 0) console.warn("无需检查Y/N 异常取值: " + built.badFlags.join(", "));
    if (built.duplicates.length > 0) console.warn("Workcenter 重复机台号: " + built.duplicates.join(", "));

    // 3. 计算差异
    const wcSet = new Set(Object.keys(wcMap));
    const type1 = [];  // 点检有/计划账无
    const type2 = [];  // 计划账有/点检无

    const exemptType1Process = ["Plasma"];
    injRows.forEach(function (r) {
      if (r.machineNo && !wcSet.has(r.machineNo)) {
        if (exemptType1Process.includes(String(r.rowData[2] || "").trim())) return; // 机型豁免
        type1.push(r);
      }
    });

    const exemptType2WC = ["V2FTA164", "V2FTA264", "V2FTA364"];
    wcSet.forEach(function (wc) {
      if (!injMachineNos.has(wc)) {
        if (exemptType2WC.includes(wc)) return;                               // 指定机台豁免
        // 豁免：Final Machine Type 为 6AX / DP / HS 的机台不纳入差异类型2
        const exempt = ["6AX", "DP", "HS"];
        if (exempt.includes(wcMap[wc].machineModel)) return;
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
      const subject = "【点检核对】 注塑机台差异报告 " + today;
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

    var t2Headers = ["Workcenter", "Machine Type", "Final Machine Type"];
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
