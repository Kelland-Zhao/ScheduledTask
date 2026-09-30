// V20260930.12 — 点检机台核对 + 主数据双向同步
// 入口：checkPointCheckMachines（每日 08:25 定时 or 手动，手动绕过安全阀）
// 预演：dryRunPointCheckSync（只报告不写入）
// 逻辑：纳入集 = Workcenter 全部行 − (「无需检查Y/N」= Y)
//       纳入集有/MachineList-INJ 无 → 追加到表尾（工序=INJ，车间按第2位推，机型/点检人留空）
//       MachineList-INJ 有/纳入集无 → 删除该机台号全部 INJ 行（Plasma 机型豁免）
//       安全阀：拟删 > 10 台 → 只报告不删；手动运行绕过
// 安全不变量：工序 ≠ INJ 的行，值和背景色一律不写入

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

// ========== 邮件状态字面量 ==========
const _pc_ST_ADDED = "已添加·需人工维护";
const _pc_ST_DEL_FLAG = "已删除·因表11标Y";
const _pc_ST_DEL_ABSENT = "已删除·表11查不到";
const _pc_ST_DEL_SKIPPED = "未执行·超阈值";
const _pc_ST_PLASMA = "仅报告·Plasma豁免";
const _pc_ST_INCOMPLETE = "待补全";
// 预演：什么都没写，状态列必须说实话，否则首跑核对时会被误导
const _pc_ST_ADD_PLANNED = "计划追加·预演未写入";
const _pc_ST_DEL_PLANNED = "计划删除·预演未写入";

// ========== 邮件中的两张表链接 ==========
const _pc_URL_POINTCHECK = "https://docs.google.com/spreadsheets/d/" + _pc_ID_POINTCHECK + "/edit#gid=436306312";
const _pc_URL_PLAN = "https://docs.google.com/spreadsheets/d/" + _pc_ID_PLAN + "/edit#gid=0";

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
/**
 * 主入口：核对 + 同步
 * @param {Object} e 触发器事件对象。undefined（手动）= 绕过安全阀
 */
function checkPointCheckMachines(e) {
  _pc_run(e !== undefined && e !== null, null);
}

/** 预演入口：只报告不写入。用于上线首次核对将要追加/删除的清单 */
function dryRunPointCheckSync() {
  _pc_run(false, { dryRun: true });
}

/**
 * 编排：读表 → 算计划 → 过阀 → 写入 → 标黄 → 发信 → 记日志
 * @param {boolean} isScheduled true=定时（启用安全阀）
 * @param {Object|null} opts {dryRun: true} 时不执行任何写入
 */
function _pc_run(isScheduled, opts) {
  const dryRun = !!(opts && opts.dryRun);
  const trigger = dryRun ? "预演" : (isScheduled ? "定时" : "手动");

  try {
    console.log("开始执行点检机台核对...");

    // 1. 读 MachineList，过滤工序=INJ
    const ssPC = SpreadsheetApp.openById(_pc_ID_POINTCHECK);
    const wsML = ssPC.getSheetByName(_pc_SHEET_MACHINELIST);
    const dataML = wsML.getDataRange().getValues();

    if (dataML.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 为空或只有表头", trigger, "");
      return;
    }

    const injRows = [];
    for (let i = 1; i < dataML.length; i++) {
      if (String(dataML[i][0] || "").trim() === _pc_FILTER_PROCESS) {
        injRows.push({ rowIndex: i + 1, machineNo: String(dataML[i][3] || "").trim(), rowData: dataML[i] });
      }
    }
    console.log("MachineList INJ 工序行数: " + injRows.length);
    if (injRows.length === 0) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 无 INJ 工序数据", trigger, "");
      return;
    }

    // 2. 读 Workcenter，建纳入集
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
    console.log("Workcenter 纳入集机台数: " + Object.keys(built.included).length);

    // 3. 算计划
    const plan = _pc_computeSyncPlan(injRows, built);
    // 4. 过安全阀（只作用于删除）
    const valve = _pc_shouldDelete(plan.toDelete.length, !isScheduled);

    console.log("追加 " + plan.append.length + " 台 / 删除 " + plan.toDelete.length + " 台 / 仅报告 " + plan.plasmaKept.length + " 台");
    if (!valve.ok) console.warn("安全阀: " + valve.reason);

    // 5-8. 写入与标黄
    // 待补全必须取**写入后**的状态：当天新追加的行机型/点检人为空，
    // 若沿用写入前的 injRows，这些行要到次日才进「待补全」区块，
    // 而那正是"安静的欠账"——这个机制存在的全部意义就是让它们当天可见
    let deleted = [];
    let incomplete = [];
    if (dryRun) {
      incomplete = _pc_collectIncomplete(injRows).concat(plan.append.map(function (a) {
        return { machineNo: a.machineNo, rowData: [_pc_FILTER_PROCESS, a.workshop, "", a.machineNo, ""] };
      }));
    } else {
      const exec = _pc_executePlan(wsML, dataML, plan, valve, injRows);
      deleted = exec.deleted;
      incomplete = exec.incomplete;
    }

    // 9. 邮件
    const result = {
      append: plan.append,
      deleted: valve.ok ? plan.toDelete : [],
      plasmaKept: plan.plasmaKept,
      incomplete: incomplete,
      deleteSkipped: !valve.ok,
      valveReason: valve.reason,
      badFlags: built.badFlags,
      duplicates: built.duplicates,
      dryRun: dryRun,
    };

    const recipients = _pc_getRecipients();
    if (recipients.length === 0) {
      writeLog("checkPointCheckMachines", "跳过", "无匹配收件人(O=INJ,P=S&C)", trigger, "");
      console.warn("未找到匹配收件人");
      return;
    }

    const today = formatVariableAsDate(new Date());
    const subject = "【点检同步】 注塑机台差异报告 " + today;
    const html = _pc_buildEmailHtml(result, today);

    // 删除快照写进 Log，邮件会过期、快照是唯一的恢复依据
    const snapshotJson = JSON.stringify(result.deleted.map(function (d) {
      return { machineNo: d.machineNo, reason: d.reason, rows: d.snapshot };
    }));

    try {
      _pc_sendMail(recipients.join(","), subject, html);
      const summary = "追加=" + plan.append.length + "台, 删除=" + result.deleted.length + "台"
        + (valve.ok ? "" : "(安全阀拦下)") + ", 待补全=" + incomplete.length + "行, TO=" + recipients.length + "人";
      writeLog("checkPointCheckMachines", "成功", summary, trigger, snapshotJson);
      console.log("邮件发送成功: " + summary);
    } catch (err) {
      writeLog("checkPointCheckMachines", "失败", err.message, trigger, snapshotJson);
      console.error("发送失败: " + err.message);
    }

    console.log("点检机台核对执行完毕");

  } catch (err) {
    console.error(err.stack || err.message);
    try { writeLog("checkPointCheckMachines", "失败", err.message, trigger, err.stack || ""); } catch (e2) { }
  }
}

/**
 * 执行写入：删除 → 追加 → 重新读背景 → 标黄
 * 安全不变量：工序 ≠ INJ 的行，值和背景色一律不写入
 * @returns {{deleted: Array, incomplete: Array}} 实际删除的条目（安全阀拦下时为空）
 *   与写入后仍待补全的行（含刚刚追加的，它们天然机型/点检人为空）
 */
function _pc_executePlan(wsML, dataML, plan, valve, injRows) {
  const deleted = valve.ok ? plan.toDelete : [];

  // ---- 删除：按连续行号区间批量 deleteRows（只含 INJ 行）----
  if (deleted.length > 0) {
    const targets = [];
    deleted.forEach(function (d) { d.rowIndexes.forEach(function (r) { targets.push(r); }); });
    targets.sort(function (a, b) { return b - a; });   // 从大到小删，避免行号移位

    let runStart = targets[0];
    let runCount = 1;
    for (let i = 1; i <= targets.length; i++) {
      if (i < targets.length && targets[i] === runStart - runCount) {
        runCount++;
      } else {
        wsML.deleteRows(runStart - runCount + 1, runCount);
        if (i < targets.length) { runStart = targets[i]; runCount = 1; }
      }
    }
    console.log("已删除 " + deleted.length + " 台机台，共 " + targets.length + " 行");
  }

  // ---- 追加：一律到表尾，不插任何现有行中间 ----
  if (plan.append.length > 0) {
    const lastRow = wsML.getLastRow();
    const rows = plan.append.map(function (a) {
      return [_pc_FILTER_PROCESS, a.workshop, "", a.machineNo, ""];
    });
    wsML.insertRowsAfter(lastRow, rows.length);
    wsML.getRange(lastRow + 1, 1, rows.length, 5).setValues(rows);
    console.log("已在表尾追加 " + rows.length + " 行");
  }

  // ---- 标黄：行号已因删除/追加而变，必须重新读背景 ----
  const lastRow = wsML.getLastRow();
  if (lastRow <= 1) return { deleted: deleted, incomplete: [] };

  const bgData = wsML.getRange(1, 1, lastRow, 5).getBackgrounds();
  const freshML = wsML.getRange(1, 1, lastRow, 5).getValues();

  const injIdx = new Set();
  const incompleteIdx = new Set();
  const incompleteRows = [];
  for (let i = 1; i < freshML.length; i++) {
    if (String(freshML[i][0] || "").trim() !== _pc_FILTER_PROCESS) continue;
    injIdx.add(i);
    if (_pc_isIncomplete(freshML[i])) {
      incompleteIdx.add(i);
      incompleteRows.push({ machineNo: String(freshML[i][3] || "").trim(), rowData: freshML[i] });
    }
  }

  _pc_computeColorActions(bgData, injIdx, incompleteIdx).forEach(function (a) {
    wsML.getRange(a.start + 1, 1, a.count, 5)
      .setBackgrounds(Array(a.count).fill(Array(5).fill(a.color)));
  });
  console.log("标黄完成：待补全 " + incompleteIdx.size + " 行");

  return { deleted: deleted, incomplete: incompleteRows };
}

/** 收集当前 INJ 中的待补全行，供邮件列出 */
function _pc_collectIncomplete(injRows) {
  return injRows.filter(function (r) { return _pc_isIncomplete(r.rowData); })
    .map(function (r) { return { machineNo: r.machineNo, rowData: r.rowData }; });
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
/**
 * 构造同步报告邮件
 * @param {Object} result 见 Task 6 Interfaces
 * @param {string} today 已格式化的日期串
 * @returns {string} HTML
 */
function _pc_buildEmailHtml(result, today) {
  const statusOf = function (reason) {
    if (result.dryRun) return _pc_ST_DEL_PLANNED;
    return reason === "FLAG" ? _pc_ST_DEL_FLAG : _pc_ST_DEL_ABSENT;
  };

  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:960px;margin:0 auto">';

  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">注塑机台同步报告</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">比对范围：MachineList(工序=INJ) ↔ Workcenter(排除「无需检查Y/N」=Y)</p>';
  html += '<p style="margin:4px 0 0;opacity:0.95;font-size:14px">';
  html += '<a href="' + _pc_URL_POINTCHECK + '" style="color:#fff">点检机台主数据</a>';
  html += ' &nbsp;|&nbsp; ';
  html += '<a href="' + _pc_URL_PLAN + '" style="color:#fff">注塑计划机台</a>';
  html += '</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + today + '</p>';
  html += '</div>';

  if (result.dryRun) {
    html += '<div style="background:#fff3cd;border-left:4px solid #e67e22;padding:12px 24px;color:#7a5b00">';
    html += '<strong>预演模式：仅报告，未写入任何数据。</strong></div>';
  }

  html += '<div style="padding:24px">';

  html += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px"><tr>';
  html += _pc_card(result.dryRun ? "计划追加" : "本次追加", result.append.length, "#27ae60");
  html += _pc_card(result.dryRun ? "计划删除" : "本次删除", result.deleteSkipped ? 0 : result.deleted.length, "#e74c3c");
  html += _pc_card("待补全", result.incomplete.length, "#e67e22");
  html += '</tr></table>';

  const nothing = result.append.length === 0 && result.deleted.length === 0
    && result.plasmaKept.length === 0 && result.incomplete.length === 0;

  if (result.deleteSkipped) {
    html += '<p style="color:#e74c3c;font-weight:bold">⚠ ' + result.valveReason + '</p>';
  }
  if (nothing && !result.deleteSkipped) {
    html += '<p style="color:#27ae60;font-weight:bold;font-size:16px">★ 点检机台与计划账完全一致，无差异。</p>';
  }

  // ===== 本次追加 =====
  if (result.append.length > 0) {
    html += _pc_section("本次追加：" + result.append.length + " 台");
    html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">⚠ 新机台的「机型」「点检人」为空白，请人工补齐（补齐前该行在点检表中标黄）</p>';
    html += buildHtmlTable(
      ["机台号", "车间", "状态"],
      result.append.map(function (r) {
        return [r.machineNo, r.workshop || "(待补全)", result.dryRun ? _pc_ST_ADD_PLANNED : _pc_ST_ADDED];
      }),
      "#27ae60");
  }

  // ===== 本次删除 =====
  if (result.deleted.length > 0 || result.deleteSkipped) {
    const n = result.deleteSkipped ? 0 : result.deleted.length;
    const title = result.dryRun ? "计划删除：" + n + " 台（预演未写入）"
      : "本次删除：" + n + " 台" + (result.deleteSkipped ? "（" + _pc_ST_DEL_SKIPPED + "）" : "");
    html += _pc_section(title);
    if (result.dryRun) {
      html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">以下机台<b>尚未删除</b>。确认无误后手动运行 checkPointCheckMachines 执行</p>';
    } else if (result.deleteSkipped) {
      html += '<p style="color:#e74c3c;font-weight:bold;margin-bottom:8px">以下机台本次<b>未删除</b>，待人工确认后手动运行放行</p>';
    } else {
      html += '<p style="color:#e74c3c;font-weight:bold;margin-bottom:8px">已从点检表移除；如需恢复，请照下表整行字段补回</p>';
    }
    // 快照展开：一台机可能占多行，逐行输出便于恢复
    const rows = [];
    result.deleted.forEach(function (d) {
      const st = statusOf(d.reason);
      d.snapshot.forEach(function (s) {
        rows.push([s[0] || "-", s[1] || "-", s[2] || "-", s[3] || "-", s[4] || "-", st]);
      });
    });
    html += buildHtmlTable(["工序", "车间", "机型", "机台号", "点检人", "状态"], rows, "#E60012");
  }

  // ===== 待补全 =====
  if (result.incomplete.length > 0) {
    html += _pc_section("待补全：" + result.incomplete.length + " 行");
    html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">⚠ 「机型」或「点检人」为空，点检任务可能无法正常派发，请补齐</p>';
    html += buildHtmlTable(
      ["机台号", "车间", "状态"],
      result.incomplete.map(function (r) { return [r.machineNo, (r.rowData[1] || "-"), _pc_ST_INCOMPLETE]; }),
      "#e67e22");
  }

  // ===== 仅报告 =====
  if (result.plasmaKept.length > 0) {
    html += _pc_section("仅报告：" + result.plasmaKept.length + " 台");
    html += '<p style="color:#7f8c8d;margin-bottom:8px">这些机台在表11 中查不到，但机型为 Plasma，按规则不删除</p>';
    html += buildHtmlTable(
      ["机台号", "状态"],
      result.plasmaKept.map(function (r) { return [r.machineNo, _pc_ST_PLASMA]; }),
      "#E60012");
  }

  // ===== 数据质量告警 =====
  if (result.badFlags.length > 0 || result.duplicates.length > 0) {
    html += _pc_section("数据质量告警");
    if (result.badFlags.length > 0) {
      html += '<p style="color:#e74c3c">「无需检查Y/N」出现异常取值（<b>不会被当作 Y 排除</b>）：'
        + result.badFlags.join("、") + '。请改为 Y 或清空。</p>';
    }
    if (result.duplicates.length > 0) {
      html += '<p style="color:#e74c3c">Workcenter 出现重复机台号（已取第一条）：'
        + result.duplicates.join("、") + '</p>';
    }
  }

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由 PointCheck Alert 系统自动发送</p>';
  html += '</div></div></body></html>';
  return html;
}

function _pc_section(title) {
  return '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:32px">' + title + '</h3>';
}

function _pc_card(label, value, color) {
  return '<td style="text-align:center;padding:16px;border:1px solid #ecf0f1;width:33%">' +
    '<div style="font-size:32px;font-weight:bold;color:' + color + '">' + value + '</div>' +
    '<div style="color:#7f8c8d;font-size:13px;margin-top:6px;line-height:1.5">' + label + '</div></td>';
}
