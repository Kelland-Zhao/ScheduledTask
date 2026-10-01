// V20260930.02 — Workcenter 数据同步（Line Database → Workcenter，按 Workcenter 键的行级同步）
// 入口：syncWorkcenterData（每日 08:20 定时 or 手动）
// 逻辑：以 Workcenter 为键就地更新 A–L 与 S（工艺无需检查Y/N）共 13 个程序列，M–R 与 T（点检无需检查Y/N）共 7 个人工列一律不碰
//   · A/B/C/D/J/L ← 1. Line Database 机台 + Equipment_Number_EAM 设备编号
//   · E–I ← 2. Active Cell 的 M–Q，键为 D 列 New Formed Cell
//   · K 是否主设备 ← 机台号出现在 2. Active Cell 的 D 列即为 Y，否则 N
// 列定位一律按表头名解析，不硬编码列号

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

    const summary = "更新 " + plan.report.updated.length +
      " / 新增 " + plan.report.added.length +
      " / 删除 " + plan.report.deleted.length +
      " / E–I 同步 " + plan.report.cellSynced +
      " / E–I 清空 " + plan.report.cellCleared.length;
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

