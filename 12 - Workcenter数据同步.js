// V20260930.02 — Workcenter 数据同步（Line Database → Workcenter，按 Workcenter 键的行级同步）
// 入口：syncWorkcenterData（每日 08:20 定时 or 手动）
// 逻辑：以 Workcenter 为键就地更新 A–L 与 S（工艺无需检查Y/N）共 13 个程序列，M–R 与 T（点检无需检查Y/N）共 7 个人工列一律不碰
//   · A/B/C/D/J/L ← 1. Line Database 机台 + Equipment_Number_EAM 设备编号
//   · E–I ← 2. Active Cell 的 M–Q，键为 D 列 New Formed Cell
//   · K 是否主设备 ← 机台号出现在 2. Active Cell 的 D 列即为 Y，否则 N
// 列定位一律按表头名解析，不硬编码列号
// 同步末尾另发「机台主数据维护提醒」：J≠NA 且 M–R 六个人工列任一为空的机台每天提醒，直到补齐
//   · 定时运行 → TO 注塑 S&C、CC 注塑 IDL；手动运行只发操作者（防调试误伤）；测试入口 testWorkcenterMaintenanceReminder

// ========== 数据源配置 ==========
const _ws_ID_PLAN = "11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc";
const _ws_SHEET_PLAN = "1. Line Database";
const _ws_SHEET_ACTIVE_CELL = "2. Active Cell";

const _ws_ID_EQU = "12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM";
const _ws_SHEET_EQU = "Workcenter";
const _ws_SHEET_EQUIPMENT_NUMBER = "Equipment_Number_EAM";

// 本脚本管理的 13 个程序列（按表头名定位）
const _WS_MANAGED_HEADERS = [
  "Workcenter", "Machine Type", "机器性能", "New Formed Cell",
  "HIM/Auto", "VIM-1", "VIM-2", "VIM-3", "VIM-4",
  "Final Machine Type", "是否主设备", "设备编号",
  "工艺无需检查Y/N",
];

// 机组配置的 5 列：来源 2. Active Cell 的 M–Q，键为 New Formed Cell
const _WS_CELL_HEADERS = ["HIM/Auto", "VIM-1", "VIM-2", "VIM-3", "VIM-4"];

// 表头行最多往下找几行（2. Active Cell 表头占两行）
const _WS_HEADER_SCAN_ROWS = 5;

// 安全阀：源表机台数低于表内现有行数这个比例时中止，防止源表异常读空后把整表删光
const _WS_MIN_SOURCE_RATIO = 0.5;

// ========== 机台主数据维护提醒配置 ==========
// 人工维护列（EDS「注塑机台主数据」维护页的维护对象）：任一为空且 J≠NA 即进提醒名单
// 注意：S（工艺无需检查Y/N）是上面 _WS_MANAGED_HEADERS 里的程序列；T 空是「需点检」的合法默认，都不在此列
const _WS_MAINT_HEADERS = ["机型", "设备类型1", "设备类型2", "自动化类型", "责任人", "备份责任人"];
const _WS_MAINT_EXCLUDED_FINAL_TYPE = "NA";   // 报废/闲置机台不要求维护
const _WS_MAINT_SENDER_NAME = "机台主数据维护提醒";
const _ws_MAINT_OPERATOR_EMAIL = "kelland_zhao@colpal.com";   // 手动运行只发这里（防误伤）；测试入口也发这里

// 收件人来源表：userID（前两行为分组/字段表头；模块 13 同款列位）
const _ws_ID_USER = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const _ws_SHEET_USERID = "userID";
const _ws_USER_COL_PROC = 14;   // O 列 工序（EDS 组）
const _ws_USER_COL_POS = 15;    // P 列 职位（EDS 组）
const _ws_USER_COL_MAIL = 9;    // J 列 GMail

// EDS「注塑机台主数据」维护页（EQU-Digital-System 生产部署路由）与 Workcenter 表链接
const _ws_URL_EDS_MM = "https://script.google.com/a/colpal.com/macros/s/AKfycbyaQjG5yFGYxU825DrODhSLl2bdfbYKpqAH4qOIzKoTJ4b-5qU/exec?v=INJ_MachineMaster";
const _ws_URL_WC = "https://docs.google.com/spreadsheets/d/" + _ws_ID_EQU + "/edit#gid=0";

// ========== 主入口 ==========
function syncWorkcenterData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    console.log("开始执行 Workcenter 数据同步...");

    // 1. 读源机台
    const sourceMachines = _ws_readSourceMachines();
    console.log("成功读取 " + sourceMachines.length + " 台机台");

    // 2. 读目标表
    const targetWs = SpreadsheetApp.openById(_ws_ID_EQU).getSheetByName(_ws_SHEET_EQU);
    if (!targetWs) throw new Error("找不到工作表: " + _ws_SHEET_EQU);
    const targetData = targetWs.getDataRange().getValues();
    const targetCols = _ws_requireHeaders(targetData[0], _WS_MANAGED_HEADERS, _ws_SHEET_EQU);

    // 2.5 安全阀：源表异常时中止，宁可不同步也不能把整表删光
    const guard = _ws_checkSourceGuard(sourceMachines.length, targetData.length - 1);
    if (!guard.ok) {
      console.log("⛔ " + guard.reason);
      try { writeLog("syncWorkcenterData", "跳过", guard.reason, trigger, ""); } catch (e2) {}
      return;
    }

    // 3. 建设备编号字典（机台号 → 设备编号）
    const eamMap = _ws_readEquipmentMap();
    console.log("设备编号字典: " + Object.keys(eamMap).length + " 条");

    // 4. 读机组配置（New Formed Cell → HIM/Auto + VIM-1~4）
    const cellSync = _ws_readActiveCellMap();
    console.log("机组配置: " + Object.keys(cellSync.map).length + " 条" + (cellSync.available ? "" : "（本次跳过 E–I 同步）"));

    // 5. 计算新的表内容
    const plan = _ws_planSync({
      targetData: targetData,
      targetCols: targetCols,
      sourceMachines: sourceMachines,
      eamMap: eamMap,
      cellMap: cellSync.map,
      cellsAvailable: cellSync.available,
    });

    // 6. 写回
    _ws_writeBack(targetWs, targetData[0].length, plan.matrix, targetData.length - 1);

    // 7. 机台主数据维护提醒（M–R 人工列未补齐的机台每天提醒；内部自捕获，失败不影响同步结果）
    const maintNote = _ws_runMaintenanceReminder(targetData[0], plan, trigger);

    const summary = "更新 " + plan.report.updated.length +
      " / 新增 " + plan.report.added.length +
      " / 删除 " + plan.report.deleted.length +
      " / E–I 同步 " + plan.report.cellSynced +
      " / E–I 清空 " + plan.report.cellCleared.length +
      maintNote;
    console.log("同步完成：" + summary);
    if (plan.report.deleted.length > 0) {
      console.log("已删除的机台: " + plan.report.deleted.join(", "));
    }
    try { writeLog("syncWorkcenterData", "成功", summary, trigger, ""); } catch (e2) {}
  } catch (err) {
    console.log("❌ 主函数执行错误: " + err.toString());
    console.log("错误堆栈: " + (err.stack || ""));
    try { writeLog("syncWorkcenterData", "失败", err.message, trigger, err.stack || ""); } catch (e2) {}
  }
}

// ========== 数据读取 ==========
function _ws_readSourceMachines() {
  const ws = SpreadsheetApp.openById(_ws_ID_PLAN).getSheetByName(_ws_SHEET_PLAN);
  if (!ws) throw new Error("找不到工作表: " + _ws_SHEET_PLAN);

  const data = ws.getDataRange().getValues();
  const cols = _ws_requireHeaders(data[0], ["Individual Machine", "Machine Type", "机器性能", "New Formed Cell"], _ws_SHEET_PLAN);

  const machines = [];
  for (let i = 1; i < data.length; i++) {
    const workcenter = String(data[i][cols["Individual Machine"]] || "").trim();
    if (!workcenter) continue;

    const item = {
      "Individual Machine": data[i][cols["Individual Machine"]],
      "Machine Type": data[i][cols["Machine Type"]],
      "机器性能": data[i][cols["机器性能"]],
      "New Formed Cell": data[i][cols["New Formed Cell"]],
    };

    // J列 Final Machine Type：机器性能不为空则用机器性能，否则用 Machine Type
    let finalMachineType = "";
    if (item["机器性能"] && item["机器性能"].toString().trim() !== "") {
      finalMachineType = item["机器性能"];
    } else {
      finalMachineType = item["Machine Type"];
    }
    finalMachineType = _ws_convertFinalMachineType(finalMachineType);

    machines.push({
      workcenter: workcenter,
      machineType: item["Machine Type"],
      performance: item["机器性能"] === undefined ? "" : item["机器性能"],
      newFormedCell: String(item["New Formed Cell"] || "").trim(),
      finalMachineType: finalMachineType,
    });
  }
  return machines;
}

// 设备编号字典：Equipment_Number_EAM 的「机台号 - Tag」→「设备」，按机台号匹配
function _ws_readEquipmentMap() {
  const map = {};
  const ws = SpreadsheetApp.openById(_ws_ID_EQU).getSheetByName(_ws_SHEET_EQUIPMENT_NUMBER);
  if (!ws) {
    console.log("⚠️ 找不到工作表: " + _ws_SHEET_EQUIPMENT_NUMBER + "，设备编号将留空");
    return map;
  }

  const data = ws.getDataRange().getValues();
  if (data.length < 2) return map;

  const cols = _ws_requireHeaders(data[0], ["设备", "机台号 - Tag"], _ws_SHEET_EQUIPMENT_NUMBER);
  for (let i = 1; i < data.length; i++) {
    const tag = _ws_cellText(data[i][cols["机台号 - Tag"]]);
    if (!tag || map[tag] !== undefined) continue; // 同一机台多条设备记录时取第一条
    const equipment = _ws_cellText(data[i][cols["设备"]]);
    if (equipment) map[tag] = equipment;
  }
  return map;
}

// 机组配置字典：2. Active Cell 的 New Formed Cell → [HIM/Auto, VIM-1~4]
// 表头占两行（第 1 行分组、第 2 行字段名），所以扫描定位而不是写死行号
function _ws_readActiveCellMap() {
  const ws = SpreadsheetApp.openById(_ws_ID_PLAN).getSheetByName(_ws_SHEET_ACTIVE_CELL);
  if (!ws) {
    console.log("⚠️ 找不到工作表: " + _ws_SHEET_ACTIVE_CELL + "，本次跳过 E–I 同步");
    return { available: false, map: {} };
  }

  const data = ws.getDataRange().getValues();
  const required = ["New Formed Cell"].concat(_WS_CELL_HEADERS);
  const headerRow = _ws_findHeaderRow(data, required);
  if (headerRow < 0) {
    console.log("⚠️ " + _ws_SHEET_ACTIVE_CELL + " 找不到表头（需含 " + required.join(" / ") + "），本次跳过 E–I 同步");
    return { available: false, map: {} };
  }
  const cols = _ws_headerIndex(data[headerRow]);

  const map = {};
  const duplicates = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const key = _ws_cellText(data[i][cols["New Formed Cell"]]);
    if (!key) continue;
    if (map[key] !== undefined) { // 重复键取第一条
      duplicates.push(key);
      continue;
    }
    map[key] = _WS_CELL_HEADERS.map(function (h) {
      const v = data[i][cols[h]];
      return v === undefined || v === null ? "" : v;
    });
  }

  const available = Object.keys(map).length > 0;
  if (!available) console.log("⚠️ " + _ws_SHEET_ACTIVE_CELL + " 无有效数据，本次跳过 E–I 同步");
  if (duplicates.length > 0) {
    console.log("⚠️ " + _ws_SHEET_ACTIVE_CELL + " 中 New Formed Cell 重复 " + duplicates.length +
      " 处，已取第一条: " + duplicates.slice(0, 5).join(", "));
  }
  return { available: available, map: map };
}

// 源表数量安全阀
function _ws_checkSourceGuard(sourceCount, existingCount) {
  if (sourceCount === 0) {
    return { ok: false, reason: "源表 " + _ws_SHEET_PLAN + " 读不到任何机台，已中止（未改动 Workcenter）" };
  }
  if (existingCount > 0 && sourceCount < existingCount * _WS_MIN_SOURCE_RATIO) {
    return {
      ok: false,
      reason: "源表机台数 " + sourceCount + " 不足表内现有 " + existingCount + " 台的一半，已中止（未改动 Workcenter）",
    };
  }
  return { ok: true, reason: "" };
}

// ========== 表头解析 ==========
function _ws_headerIndex(headerRow) {
  const idx = {};
  for (let i = 0; i < headerRow.length; i++) {
    const name = String(headerRow[i] || "").trim();
    if (name && idx[name] === undefined) idx[name] = i; // 重名取第一列
  }
  return idx;
}

function _ws_findHeaderRow(data, requiredNames) {
  const limit = Math.min(_WS_HEADER_SCAN_ROWS, data.length);
  for (let r = 0; r < limit; r++) {
    const idx = _ws_headerIndex(data[r]);
    if (requiredNames.every(function (n) { return idx[n] !== undefined; })) return r;
  }
  return -1;
}

function _ws_requireHeaders(headerRow, names, sheetLabel) {
  const idx = _ws_headerIndex(headerRow);
  const missing = names.filter(function (n) { return idx[n] === undefined; });
  if (missing.length > 0) {
    throw new Error(sheetLabel + " 表头缺少字段: " + missing.join(", ") + "（未改动任何数据）");
  }
  return idx;
}

// ========== 同步计算（纯函数，不碰表格） ==========
function _ws_planSync(opts) {
  const targetData = opts.targetData;
  const cols = opts.targetCols;
  const width = targetData[0].length;
  const rows = targetData.slice(1);

  const byMachine = new Map();
  opts.sourceMachines.forEach(function (m) {
    if (!byMachine.has(m.workcenter)) byMachine.set(m.workcenter, m); // 重复机台号取第一条
  });

  const report = { updated: [], added: [], deleted: [], cellSynced: 0, cellCleared: [] };
  const ctx = {
    cols: cols,
    eamMap: opts.eamMap || {},
    cellMap: opts.cellMap || {},
    cellsAvailable: opts.cellsAvailable === true,
    report: report,
  };
  const matrix = [];

  // 表里已有：源里还在就地更新，源里没了整行丢弃（即删除）
  rows.forEach(function (row) {
    const key = _ws_cellText(row[cols["Workcenter"]]);
    if (!key) return;
    const machine = byMachine.get(key);
    if (!machine) {
      report.deleted.push(key);
      return;
    }
    matrix.push(_ws_buildRow(row, width, machine, ctx));
    report.updated.push(key);
  });

  // 源里有、表里没有：追加到末尾，人工列留空
  const handled = new Set(report.updated);
  opts.sourceMachines.forEach(function (m) {
    if (handled.has(m.workcenter)) return;
    handled.add(m.workcenter);
    matrix.push(_ws_buildRow(new Array(width).fill(""), width, m, ctx));
    report.added.push(m.workcenter);
  });

  return { matrix: matrix, report: report };
}

function _ws_buildRow(row, width, machine, ctx) {
  const cols = ctx.cols;
  const out = [];
  for (let i = 0; i < width; i++) out.push(row[i] === undefined ? "" : row[i]);

  out[cols["Workcenter"]] = machine.workcenter;
  out[cols["Machine Type"]] = machine.machineType;
  out[cols["机器性能"]] = machine.performance;
  out[cols["New Formed Cell"]] = machine.newFormedCell;
  out[cols["Final Machine Type"]] = machine.finalMachineType;
  out[cols["设备编号"]] = ctx.eamMap[machine.workcenter] || "";

  // 是否主设备：机台号出现在 2. Active Cell 的 D 列（New Formed Cell）即为 Y，否则 N
  // 按 A 列 Workcenter 匹配。Active Cell 不可用时保留原值（新增行无原值，留空）
  if (ctx.cellsAvailable) {
    out[cols["是否主设备"]] = Object.prototype.hasOwnProperty.call(ctx.cellMap, machine.workcenter) ? "Y" : "N";
  }

  // E–I：机组配置，键为 New Formed Cell；键为空（闲置）的行不参与
  if (ctx.cellsAvailable && machine.newFormedCell) {
    const cellValues = ctx.cellMap[machine.newFormedCell];
    if (cellValues) {
      _WS_CELL_HEADERS.forEach(function (h, i) { out[cols[h]] = cellValues[i]; });
      ctx.report.cellSynced++;
    } else {
      _WS_CELL_HEADERS.forEach(function (h) { out[cols[h]] = ""; });
      ctx.report.cellCleared.push(machine.workcenter);
    }
  }

  // 工艺无需检查Y/N：三条判据命中任一即 Y，否则留空
  //   1) Final Machine Type=6AX 且非主设备  2) Final Machine Type=NA（已退役）
  //   3) D=H2HTB363 且非主设备（下述一次性豁免）
  // 由 J/K/D 派生，所以放在最后算，取的是本行最终值
  // 紧邻的「点检无需检查Y/N」是人工维护列：不在 _WS_MANAGED_HEADERS 里，
  // 行重写时靠基础行整行带过来，本模块一律不写（见 tests/workcenter-sync.test.mjs 保留用例）
  const finalType = _ws_cellText(out[cols["Final Machine Type"]]);
  const isMain = _ws_cellText(out[cols["是否主设备"]]);
  // 一次性豁免（仅 H2HTB363）：判据绑在 New Formed Cell 这个单元格上、而非机台号名单，
  // 机台进出该单元格时自动跟随，不会像硬编码的机台号那样过期
  const cellExempt = _ws_cellText(out[cols["New Formed Cell"]]) === "H2HTB363" && isMain === "N";
  out[cols["工艺无需检查Y/N"]] =
    ((finalType === "6AX" && isMain === "N") || finalType === "NA" || cellExempt) ? "Y" : "";

  return out;
}

function _ws_cellText(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim();
}

// ========== 机台主数据维护提醒 ==========
// 判据（纯函数）：J≠NA 且 M–R 六个人工列任一为空 → 进名单；返回 [{workcenter, missing, isNew}]
// matrix 为最终写入的内容（新增行人工列天然为空，会被当日提醒；补齐后次日自动停止）
function _ws_findMaintenancePending(cols, matrix, addedSet) {
  const pending = [];
  const seen = new Set();
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const workcenter = _ws_cellText(row[cols["Workcenter"]]);
    if (!workcenter || seen.has(workcenter)) continue;   // 重复机台号取第一条
    seen.add(workcenter);
    if (_ws_cellText(row[cols["Final Machine Type"]]) === _WS_MAINT_EXCLUDED_FINAL_TYPE) continue;

    const missing = _WS_MAINT_HEADERS.filter(function (h) {
      return _ws_cellText(row[cols[h]]) === "";
    });
    if (missing.length === 0) continue;
    pending.push({ workcenter: workcenter, missing: missing, isNew: addedSet.has(workcenter) });
  }
  return pending;
}

// 收件人（纯函数）：userID 前两行为表头，数据从第 3 行起
// TO = 工序 INJ 且职位含 S&C；CC = 工序 INJ 且职位 IDL（职位大小写/空格归一，工序 trim 后精确匹配）
function _ws_parseMaintenanceRecipients(userData) {
  const to = [];
  const cc = [];
  for (let r = 2; r < userData.length; r++) {
    const row = userData[r] || [];
    const proc = _ws_cellText(row[_ws_USER_COL_PROC]);
    const pos = _ws_cellText(row[_ws_USER_COL_POS]).toUpperCase();
    const mail = _ws_cellText(row[_ws_USER_COL_MAIL]).toLowerCase();
    if (!mail || proc !== "INJ") continue;
    if (pos.indexOf("S&C") >= 0) {
      if (to.indexOf(mail) < 0) to.push(mail);
    } else if (pos === "IDL") {
      if (cc.indexOf(mail) < 0) cc.push(mail);
    }
  }
  return { to: to, cc: cc };
}

function _ws_getMaintenanceRecipients() {
  const ws = SpreadsheetApp.openById(_ws_ID_USER).getSheetByName(_ws_SHEET_USERID);
  if (!ws) {
    console.warn("维护提醒：找不到 " + _ws_SHEET_USERID + " 表，收件人为空");
    return { to: [], cc: [] };
  }
  return _ws_parseMaintenanceRecipients(ws.getDataRange().getValues());
}

// 车间推导：机台号第 2 位 0/1 → TB1、2 → TB2，其他留空（与模块 13 同规则）
function _ws_deriveWorkshop(machineNo) {
  const s = _ws_cellText(machineNo);
  if (s.length < 2) return "";
  const c = s.charAt(1);
  if (c === "0" || c === "1") return "TB1";
  if (c === "2") return "TB2";
  return "";
}

// 邮件 HTML：模式 A 红底头部 + buildHtmlTable + 两个入口链接（遵循 docs/邮件UI规范.md）
function _ws_buildMaintenanceEmailHtml(pending, today) {
  let html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto">';
  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">机台主数据维护提醒</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">Workcenter 同步发现 ' + pending.length + ' 台机台的属性列未补齐</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + escapeHtml(today) + '</p>';
  html += '</div>';
  html += '<div style="padding:24px">';
  html += '<p style="color:#e67e22;font-weight:bold;margin:0 0 12px">⚠ 请在 EDS「注塑机台主数据」维护页补齐以下机台（该页面只能改、不能增删机台）</p>';
  html += buildHtmlTable(
    ["机台号", "车间", "缺失字段", "类型"],
    pending.map(function (p) {
      return [p.workcenter, _ws_deriveWorkshop(p.workcenter) || "-", p.missing.join("、"), p.isNew ? "本次新增" : "待补齐"];
    }),
    "#E60012");
  html += '<p style="font-size:13px;color:#34495e;margin-top:16px">维护入口：';
  html += '<a href="' + _ws_URL_EDS_MM + '" style="color:#E60012">注塑机台主数据维护页</a>';
  html += ' ｜ 数据表：<a href="' + _ws_URL_WC + '" style="color:#E60012">Workcenter（注塑计划机台）</a></p>';
  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由' + _WS_MAINT_SENDER_NAME + '自动发送，请勿回复</p>';
  html += '</div></div></body></html>';
  return html;
}

// 发送编排：返回给 Log 摘要的片段；任何失败都自捕获，绝不影响同步结果
function _ws_runMaintenanceReminder(headerRow, plan, trigger) {
  try {
    const cols = _ws_headerIndex(headerRow);
    const missingHeaders = _WS_MAINT_HEADERS.filter(function (n) { return cols[n] === undefined; });
    if (missingHeaders.length > 0) {
      console.log("⚠️ 表头缺少人工列（" + missingHeaders.join(", ") + "），跳过维护提醒");
      return " / 维护提醒跳过（表头缺 " + missingHeaders.join("/") + "）";
    }

    const pending = _ws_findMaintenancePending(cols, plan.matrix, new Set(plan.report.added));
    if (pending.length === 0) return " / 待维护 0 台";

    const isScheduled = trigger === "定时";
    let to = [];
    let cc = [];
    if (isScheduled) {
      const recipients = _ws_getMaintenanceRecipients();
      to = recipients.to;
      cc = recipients.cc;
      if (to.length === 0) {
        console.warn("维护提醒：无匹配收件人（工序=INJ、职位含 S&C），本次不发信");
        return " / 待维护 " + pending.length + " 台（未发提醒：无收件人）";
      }
    } else {
      to = [_ws_MAINT_OPERATOR_EMAIL];   // 手动运行只发操作者，防止调试时误发全员
    }

    const today = formatVariableAsDate(new Date());
    const subject = "【机台主数据维护】 " + pending.length + " 台机台属性未补齐 - " + today;
    const options = { htmlBody: _ws_buildMaintenanceEmailHtml(pending, today), name: _WS_MAINT_SENDER_NAME };
    if (cc.length > 0) options.cc = cc.join(",");
    GmailApp.sendEmail(to.join(","), subject, "", options);
    console.log("维护提醒已发送: " + pending.length + " 台, TO=" + to.length + "人, CC=" + cc.length + "人");
    return " / 待维护 " + pending.length + " 台（已提醒" + (isScheduled ? " TO " + to.length + "/CC " + cc.length : "·仅操作者") + "）";
  } catch (err) {
    console.error("维护提醒失败: " + err.message);
    return " / 维护提醒失败（不影响同步）";
  }
}

// 测试入口：只读生产表、只发操作者，可安全随时运行（不写任何数据）
function testWorkcenterMaintenanceReminder() {
  try {
    const ws = SpreadsheetApp.openById(_ws_ID_EQU).getSheetByName(_ws_SHEET_EQU);
    if (!ws) throw new Error("找不到工作表: " + _ws_SHEET_EQU);
    const data = ws.getDataRange().getValues();
    const cols = _ws_headerIndex(data[0]);
    const missingHeaders = _WS_MAINT_HEADERS.filter(function (n) { return cols[n] === undefined; });
    if (missingHeaders.length > 0) throw new Error("表头缺少字段: " + missingHeaders.join(", "));

    const pending = _ws_findMaintenancePending(cols, data.slice(1), new Set());
    if (pending.length === 0) {
      console.log("测试入口：当前无待维护机台，不发信");
      return;
    }

    const today = formatVariableAsDate(new Date());
    const subject = "【机台主数据维护·测试】 " + pending.length + " 台机台属性未补齐 - " + today;
    const html = _ws_buildMaintenanceEmailHtml(pending, today);
    GmailApp.sendEmail(_ws_MAINT_OPERATOR_EMAIL, subject, "", { htmlBody: html, name: _WS_MAINT_SENDER_NAME });
    console.log("测试邮件已发送: " + pending.length + " 台 → " + _ws_MAINT_OPERATOR_EMAIL);
  } catch (err) {
    console.log("❌ 测试入口错误: " + err.toString());
  }
}

// ========== 写回 ==========
function _ws_writeBack(ws, width, matrix, previousRowCount) {
  if (matrix.length > 0) {
    ws.getRange(2, 1, matrix.length, width).setValues(matrix);
  }
  if (previousRowCount > matrix.length) {
    ws.getRange(2 + matrix.length, 1, previousRowCount - matrix.length, width).clearContent();
  }
}

// ========== Final Machine Type 智能转换 ==========
// 机器性能 列里混着两类内容：细分机型（6AX/3AX/H Auto…）与状态（报废/闲置）。
// 状态类机台机型已无意义 → 统一写 NA。
const _WS_STATUS_PATTERN = /报废|闲置/;

function _ws_convertFinalMachineType(machineType) {
  if (!machineType || typeof machineType !== "string") {
    return machineType;
  }

  const typeStr = machineType.toString().trim();

  // 状态类文本（报废/闲置，含「机台号报废,原665机器闲置」这类混合写法）→ NA
  if (_WS_STATUS_PATTERN.test(typeStr)) {
    return "NA";
  }

  // E 开头 → ENG
  if (typeStr.startsWith("E")) {
    return "ENG";
  }

  // F 开头 → FCS；FT400 例外，保持原值
  if (typeStr.startsWith("F")) {
    return typeStr === "FT400" ? typeStr : "FCS";
  }

  return typeStr;
}

