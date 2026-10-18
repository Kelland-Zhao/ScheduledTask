// PME（Post-PM Effectiveness）计算任务
// 计算层与 Spreadsheet I/O 分离；内部函数统一使用 _pme_ 前缀。

var _pme_CONFIG = {
  mainSpreadsheetId: "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24",
  pmSheetName: "PM Data",
  masterDataSheetName: "Master Data",
  kpiSummarySheetName: "KPI Summary",
  exclusionSheetName: "KPI排除条件",
  tfTypeSheetName: "TF Machine Type",
  sourceSheetName: "SourceData",
  sourceChunkRows: 10000,
  aemSources: [
    { label: "2025Q1", id: "1QFvpIVWS3yv11Ujoh56cYsrBffM-zaH6Bc3h8sPCdJA", start: "2025-01-01", end: "2025-03-31" },
    { label: "2025Q2", id: "1E-Rp9v_083LXJ5DbCW1jh9Wzuh4p5gPgQP2uOn4BRvg", start: "2025-04-01", end: "2025-06-30" },
    { label: "2025Q3", id: "1TSthgE2c4RiZYsw6MiMkPy535dXWNsJgj6zicg4-OZ0", start: "2025-07-01", end: "2025-09-30" },
    { label: "2025Q4", id: "1yCtPUp84XcaWxVfwIusgGg_LIB9RgtjZ_kmpQ1AhSXs", start: "2025-10-01", end: "2025-12-31" },
    { label: "2026Q1", id: "1_vWImIpjTvUfDVY-s9Ij5Ajg5h-ayQ8XUbGcN5sEO68", start: "2026-01-01", end: "2026-03-31" },
    { label: "2026Q2", id: "1yTkY1yz4zwiXCzSI8DiHSxWYu4CmBkzsQcKH1_2_Fv8", start: "2026-04-01", end: "2026-06-30" },
    { label: "2026Q3", id: "10o2yG9Fzvlo54e0l71KeAM-QobNvUVC3xPcBZHIinC4", start: "2026-07-01", end: "2026-09-30" },
    { label: "2026Q4-current", id: "1pKLCZfh8sKv6EcWuEsxH6jAIsjbWHU5_aAusEODnJ_o", start: "2026-10-01", end: "" },
  ],
};

var _pme_MASTER_HEADERS = [
  "序列号/ Serial No.",
  "工序/ KeyProcess",
  "保养日期/ PM_Date",
  "机台号/ Workcenter",
  "Production_AvailableTime_BEFORE PM",
  "Production_AvailableTime_AFTER PM",
  "AU_BEFORE PM",
  "AU_AFTER PM",
  "UPDT_BEFORE PM",
  "UPDT_AFTER PM",
  "UPDT Gap",
  "保养月份/ YearMonth",
  "子工序/ SubProcess",
  "生产状态/ Production Status",
  "绩效状态/ Performance Status",
];

var _pme_KPI_PROCESS_CONFIG = {
  "IM": { row2025: 2, row2026: 10 },
  "TF": { row2025: 3, row2026: 11 },
  "AFT & USW": { row2025: 4, row2026: 12 },
  "HSL & Bou": { row2025: 5, row2026: 13 },
  "Zar": { row2025: 6, row2026: 14 },
  "PK": { row2025: 7, row2026: 15 },
};

/**
 * 规范化机台号：去掉不影响设备身份的尾部括号后缀。
 * @param {*} value
 * @returns {string}
 */
function _pme_normalizeWorkcenter(value) {
  return String(value == null ? "" : value)
    .trim()
    .replace(/【[^】]*】\s*$/, "")
    .trim();
}

function _pme_hasSwWorkcenter(value) {
  return String(value == null ? "" : value).indexOf("【SW】") !== -1;
}

/**
 * 按表头名称建立列索引；重复表头取第一列。
 * @param {Array<*>} headers
 * @returns {Object<string, number>}
 */
function _pme_headerIndex(headers) {
  var index = {};
  (headers || []).forEach(function (header, columnIndex) {
    var name = String(header == null ? "" : header).trim();
    if (name && index[name] === undefined) index[name] = columnIndex;
  });
  return index;
}

function _pme_requireHeaders(headers, required, sourceName) {
  var available = _pme_headerIndex(headers);
  var missing = (required || []).filter(function (name) { return available[name] === undefined; });
  if (missing.length > 0) {
    throw new Error(String(sourceName || "数据源") + " 缺少必需字段: " + missing.join(", "));
  }
  return available;
}

function _pme_validateSourceCoverage(sources) {
  var ranges = (sources || []).filter(function (source) {
    return source && source.start && source.end;
  }).slice().sort(function (a, b) {
    return String(a.start).localeCompare(String(b.start));
  });

  for (var index = 1; index < ranges.length; index++) {
    var previous = ranges[index - 1];
    var current = ranges[index];
    if (String(current.start) <= String(previous.end)) {
      throw new Error("AEM 来源日期范围重叠: " + previous.label + " / " + current.label);
    }
  }
  return true;
}

/**
 * PM 计划工序映射。
 * @param {*} value
 * @returns {string}
 */
function _pme_mapProcess(value) {
  var process = String(value == null ? "" : value).trim().toUpperCase();
  return process === "INJ" ? "IM" : process;
}

/**
 * 日期值统一为 yyyy-MM-dd。计算层不依赖环境时区解析短日期字符串。
 * @param {*} value
 * @returns {string}
 */
function _pme_dateKey(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    return [value.getFullYear(), value.getMonth() + 1, value.getDate()]
      .map(function (part, index) {
        return index === 0 ? String(part) : String(part).padStart(2, "0");
      })
      .join("-");
  }

  var text = String(value == null ? "" : value).trim();
  var iso = text.match(/^(\d{4})[-\/]([01]\d)[-\/]([0-3]\d)/);
  if (iso) return iso[1] + "-" + iso[2] + "-" + iso[3];

  var slash = text.match(/^([01]?\d)[-\/]([0-3]?\d)[-\/](\d{4})$/);
  if (slash) {
    return slash[3] + "-" + slash[1].padStart(2, "0") + "-" + slash[2].padStart(2, "0");
  }

  return text;
}

function _pme_getCell(row, headerIndex, name) {
  var columnIndex = headerIndex[name];
  return columnIndex === undefined ? "" : row[columnIndex];
}

/**
 * 筛选正式 PM 计划并转换成统一结构。
 * @param {Array<Array<*>>} rows 含表头行
 * @param {Object<string, number>} headerIndex
 * @returns {Array<Object>}
 */
function _pme_selectPmPlans(rows, headerIndex, targetYear, asOfDate) {
  if (!Array.isArray(rows) || rows.length <= 1) return [];

  var targetYearText = targetYear == null ? "" : String(targetYear).trim();
  var asOfDateKey = _pme_dateKey(asOfDate || _pme_getTodayDateKey());
  var plans = [];
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    var row = rows[rowIndex] || [];
    var category = String(_pme_getCell(row, headerIndex, "类别") || "").trim();
    var status = String(_pme_getCell(row, headerIndex, "状态") || "").trim();
    if (category.indexOf("PM") === -1 || status === "取消") continue;

    var pmDate = _pme_dateKey(_pme_getCell(row, headerIndex, "开始日期"));
    var rawWorkcenter = String(_pme_getCell(row, headerIndex, "AEM#") || "").trim();
    if (_pme_hasSwWorkcenter(rawWorkcenter)) continue;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(pmDate)) continue;
    if (targetYearText && pmDate.slice(0, 4) !== targetYearText) continue;
    if (asOfDateKey && /^\d{4}-\d{2}-\d{2}$/.test(asOfDateKey) && pmDate > asOfDateKey) continue;

    var workcenter = _pme_normalizeWorkcenter(rawWorkcenter);
    if (!pmDate || !workcenter) continue;

    var durationValue = Number(_pme_getCell(row, headerIndex, "保养时间"));
    var durationHours = Number.isFinite(durationValue) ? durationValue : 0;
    var startTime = String(_pme_getCell(row, headerIndex, "开始时间") || "").trim() || "08:00";
    var keyProcess = _pme_mapProcess(_pme_getCell(row, headerIndex, "工序"));

    plans.push({
      rowIndex: rowIndex + 1,
      serialNo: pmDate + workcenter,
      keyProcess: keyProcess,
      workcenter: workcenter,
      pmDate: pmDate,
      startTime: startTime,
      durationHours: durationHours,
      category: category,
      status: status,
      rawRow: row.slice(),
    });
  }
  return plans;
}

function _pme_getTodayDateKey() {
  var now = new Date();
  return String(now.getFullYear()) + "-" +
    String(now.getMonth() + 1).padStart(2, "0") + "-" +
    String(now.getDate()).padStart(2, "0");
}

/**
 * 按日期和班次排序。班次 1/2/3 的自然顺序由日期共同决定。
 * @param {{prodDate: string, shift: number}} a
 * @param {{prodDate: string, shift: number}} b
 * @returns {number}
 */
function _pme_compareShift(a, b) {
  var dateCompare = String(a.prodDate || "").localeCompare(String(b.prodDate || ""));
  if (dateCompare !== 0) return dateCompare;
  return Number(a.shift || 0) - Number(b.shift || 0);
}

function _pme_number(value) {
  var number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function _pme_shiftKey(prodDate, shift) {
  return _pme_dateKey(prodDate) + "-" + Number(shift);
}

function _pme_shiftKeyParts(shiftKey) {
  var match = String(shiftKey || "").match(/^(\d{4}-\d{2}-\d{2})-(\d+)$/);
  if (!match) return null;
  return { prodDate: match[1], shift: Number(match[2]) };
}

function _pme_shiftOrdinal(shiftKey) {
  var parts = _pme_shiftKeyParts(shiftKey);
  if (!parts || ![1, 2, 3].includes(parts.shift)) return null;
  var dateParts = parts.prodDate.split("-").map(Number);
  var day = Math.floor(Date.UTC(dateParts[0], dateParts[1] - 1, dateParts[2]) / 86400000);
  return day * 3 + parts.shift - 1;
}

function _pme_shiftKeyFromOrdinal(ordinal) {
  var day = Math.floor(ordinal / 3);
  var slot = ordinal - day * 3;
  var date = new Date(Date.UTC(1970, 0, 1) + day * 86400000);
  var dateKey = [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()]
    .map(function (part, index) {
      return index === 0 ? String(part) : String(part).padStart(2, "0");
    })
    .join("-");
  return dateKey + "-" + (slot + 1);
}

function _pme_shiftNumberFromTime(value) {
  var match = String(value || "08:00").trim().match(/^(\d{1,2})(?::(\d{2}))?/);
  var hour = match ? Number(match[1]) : 8;
  if (hour < 8) return 1;
  if (hour < 16) return 2;
  return 3;
}

function _pme_compareShiftKey(a, b) {
  var aParts = _pme_shiftKeyParts(a);
  var bParts = _pme_shiftKeyParts(b);
  return _pme_compareShift(aParts || {}, bParts || {});
}

/**
 * 按 Workcenter + ProdDate + Shift 聚合 AEM 分段。
 * @param {Array<Array<*>>} rows 含表头行
 * @param {Object<string, number>} headerIndex
 * @returns {Array<Object>}
 */
function _pme_aggregateAemShifts(rows, headerIndex) {
  if (!Array.isArray(rows) || rows.length <= 1) return [];
  var grouped = {};

  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    var row = rows[rowIndex] || [];
    var prodDate = _pme_dateKey(_pme_getCell(row, headerIndex, "ProdDate"));
    var workcenter = _pme_normalizeWorkcenter(_pme_getCell(row, headerIndex, "Workcenter"));
    var shift = Number(_pme_getCell(row, headerIndex, "Shift"));
    if (!prodDate || !workcenter || ![1, 2, 3].includes(shift)) continue;

    var key = workcenter + "|" + _pme_shiftKey(prodDate, shift);
    if (!grouped[key]) {
      grouped[key] = {
        shiftKey: _pme_shiftKey(prodDate, shift),
        prodDate: prodDate,
        workcenter: workcenter,
        shift: shift,
        prodHours: 0,
        netRuntime: 0,
        plantAvailableTime: 0,
        co: 0,
        hasCO: false,
        rawRows: [],
      };
    }

    var result = grouped[key];
    result.prodHours += _pme_number(_pme_getCell(row, headerIndex, "Prod_Hours"));
    result.netRuntime += _pme_number(_pme_getCell(row, headerIndex, "NetRuntime"));
    result.plantAvailableTime += _pme_number(_pme_getCell(row, headerIndex, "PlantAvailableTime"));
    result.co += _pme_number(_pme_getCell(row, headerIndex, "CO"));
    if (_pme_number(_pme_getCell(row, headerIndex, "CO")) > 0) result.hasCO = true;
    result.rawRows.push(row.slice());
  }

  return Object.keys(grouped).map(function (key) { return grouped[key]; })
    .sort(_pme_compareShift);
}

/**
 * 根据计划开始班次和连续保养时长计算窗口。
 * @param {{pmDate: string, startTime: string, durationHours: number}} plan
 * @returns {{startShiftKey: string, endShiftKey: string, afterStartShiftKey: string}}
 */
function _pme_getMaintenanceWindow(plan) {
  var startShift = _pme_shiftNumberFromTime(plan.startTime || "08:00");
  var startShiftKey = _pme_shiftKey(plan.pmDate, startShift);
  var startOrdinal = _pme_shiftOrdinal(startShiftKey);
  var shiftCount = Math.max(1, Math.ceil(_pme_number(plan.durationHours) / 8));
  var endShiftKey = _pme_shiftKeyFromOrdinal(startOrdinal + shiftCount - 1);
  var afterStartShiftKey = _pme_shiftKeyFromOrdinal(startOrdinal + shiftCount);
  return {
    startShiftKey: startShiftKey,
    endShiftKey: endShiftKey,
    afterStartShiftKey: afterStartShiftKey,
  };
}

/**
 * 选择 BEFORE PM 的有效班次。
 * IM 按整班应用保养前结存轮空；TF/PK 不应用该规则。
 * @param {string} process
 * @param {Array<Object>} shifts
 * @param {string} pmStartShiftKey
 * @returns {Array<Object>}
 */
function _pme_selectBeforeShifts(process, shifts, pmStartShiftKey) {
  var before = (shifts || [])
    .filter(function (shift) { return _pme_compareShiftKey(shift.shiftKey, pmStartShiftKey) < 0; })
    .sort(_pme_compareShift);

  var positive = before.filter(function (shift) { return shift.prodHours > 0; });
  var skipped = {};
  if (String(process || "").toUpperCase() === "IM") {
    if (positive.length > 0) skipped[positive[positive.length - 1].shiftKey] = true;

    for (var index = 1; index < before.length; index++) {
      var previous = before[index - 1];
      var current = before[index];
      if (previous.prodHours > 0 && current.prodHours === 0) {
        skipped[previous.shiftKey] = true;
      }
    }
  }

  positive = positive.filter(function (shift) { return !skipped[shift.shiftKey]; });

  // 与 AFTER PM 对称：从 PM 前最近的有效班次向前累计，
  // 首次达到或超过 144 小时后停止，最终恢复为时间正序。
  var selected = [];
  var totalProdHours = 0;
  for (var positiveIndex = positive.length - 1; positiveIndex >= 0; positiveIndex--) {
    var shift = positive[positiveIndex];
    selected.unshift(shift);
    totalProdHours += _pme_number(shift.prodHours);
    if (totalProdHours >= 144) break;
  }
  return selected;
}

/**
 * 从 AFTER PM 起点累计至首次达到或超过 144 小时。
 * @param {string} process 保留参数，便于后续工序策略扩展；当前 TF/PK 不做 CO 筛选
 * @param {Array<Object>} shifts
 * @param {string} afterStartShiftKey
 * @returns {{rows: Array<Object>, status: string, totalProdHours: number}}
 */
function _pme_selectAfterWindow(process, shifts, afterStartShiftKey) {
  var selected = [];
  var totalProdHours = 0;
  var ordered = (shifts || [])
    .filter(function (shift) { return _pme_compareShiftKey(shift.shiftKey, afterStartShiftKey) >= 0; })
    .sort(_pme_compareShift);

  for (var index = 0; index < ordered.length; index++) {
    var shift = ordered[index];
    selected.push(shift);
    totalProdHours += _pme_number(shift.prodHours);
    if (totalProdHours >= 144) {
      return { rows: selected, status: "生产", totalProdHours: totalProdHours };
    }
  }

  return { rows: selected, status: "未达144", totalProdHours: totalProdHours };
}

function _pme_sumShiftValues(shifts) {
  return (shifts || []).reduce(function (result, shift) {
    result.prodHours += _pme_number(shift.prodHours);
    result.netRuntime += _pme_number(shift.netRuntime);
    result.plantAvailableTime += _pme_number(shift.plantAvailableTime);
    return result;
  }, { prodHours: 0, netRuntime: 0, plantAvailableTime: 0 });
}

function _pme_ratio(numerator, denominator) {
  return denominator > 0 ? numerator / denominator : null;
}

function _pme_metricValue(value) {
  return value === null || value === undefined || !Number.isFinite(value) ? null : value;
}

/**
 * 计算 Before/After 的累计时间、AU、UPDT 和最终状态。
 * @param {Array<Object>} beforeShifts
 * @param {Array<Object>} afterShifts
 * @returns {Object}
 */
function _pme_calculateMetrics(beforeShifts, afterShifts) {
  var before = _pme_sumShiftValues(beforeShifts);
  var after = _pme_sumShiftValues(afterShifts);
  var beforeAU = _pme_ratio(before.netRuntime, before.plantAvailableTime);
  var afterAU = _pme_ratio(after.netRuntime, after.plantAvailableTime);
  var beforeUPDT = _pme_ratio(before.prodHours - before.netRuntime, before.prodHours);
  var afterUPDT = _pme_ratio(after.prodHours - after.netRuntime, after.prodHours);
  var productionStatus = after.prodHours >= 144 ? "生产" : "未达144";
  var updtGap = beforeUPDT !== null && afterUPDT !== null ? afterUPDT - beforeUPDT : null;
  var performanceStatus = "";

  if (productionStatus === "生产" && beforeUPDT !== null && afterUPDT !== null) {
    performanceStatus = afterUPDT <= beforeUPDT ? "好/Good" : "差/Bad";
  }

  return {
    beforeProductionAvailableTime: before.prodHours,
    afterProductionAvailableTime: after.prodHours,
    beforeAU: _pme_metricValue(beforeAU),
    afterAU: _pme_metricValue(afterAU),
    beforeUPDT: _pme_metricValue(beforeUPDT),
    afterUPDT: _pme_metricValue(afterUPDT),
    updtGap: _pme_metricValue(updtGap),
    productionStatus: productionStatus,
    performanceStatus: performanceStatus,
  };
}

/**
 * 按 PM 开始日期的 20 日切点生成 PME YearMonth。
 * @param {*} pmDate
 * @returns {string}
 */
function _pme_getYearMonth(pmDate) {
  var dateKey = _pme_dateKey(pmDate);
  var parts = dateKey.split("-").map(Number);
  if (parts.length !== 3 || parts.some(function (part) { return !Number.isFinite(part); })) return "";

  var date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (parts[2] > 20) date.setUTCDate(date.getUTCDate() + 11);
  return String(date.getUTCFullYear()) + String(date.getUTCMonth() + 1).padStart(2, "0");
}

/**
 * 判断记录是否不纳入最终 KPI。
 * @param {{keyProcess: string, workcenter: string, durationHours: number}} plan
 * @param {Array<{workcenter: string, duration: number, status?: string}>} exclusionRules
 * @returns {boolean}
 */
function _pme_isKpiExcluded(plan, exclusionRules) {
  var process = String(plan && plan.keyProcess || "").trim().toUpperCase();
  var workcenter = _pme_normalizeWorkcenter(plan && plan.workcenter);
  var durationHours = _pme_number(plan && plan.durationHours);

  if (/^S1DP000[1-4]$/.test(workcenter)) return true;
  if (process === "PK" && /^PK[SF]/.test(workcenter.toUpperCase())) return true;

  return (exclusionRules || []).some(function (rule) {
    var status = String(rule.status == null ? "启用" : rule.status).trim();
    if (status !== "启用") return false;
    return _pme_normalizeWorkcenter(rule.workcenter) === workcenter &&
      _pme_number(rule.duration) === durationHours;
  });
}

/**
 * 判断计划是否完全不生成 Master Data 明细。
 * 目前仅豁免员工维护的 S1DP0001–0004；其他 KPI 排除规则仍保留明细。
 * @param {{workcenter: string}=} plan
 * @returns {boolean}
 */
function _pme_isMasterDataExcluded(plan) {
  var workcenter = _pme_normalizeWorkcenter(plan && plan.workcenter);
  return /^S1DP000[1-4]$/.test(workcenter);
}

/**
 * Master Data 的 15 列标准行。
 * @param {Object} plan
 * @param {Object} metrics
 * @param {string} subprocess
 * @param {Object} exclusionInfo 保留参数，用于调用方记录排除原因；不新增输出列
 * @returns {Array<*>}
 */
function _pme_buildMasterRow(plan, metrics, subprocess, exclusionInfo) {
  void exclusionInfo;
  return [
    plan.serialNo || "",
    plan.keyProcess || "",
    plan.pmDate || "",
    plan.workcenter || "",
    Math.round(_pme_number(metrics.beforeProductionAvailableTime)),
    Math.round(_pme_number(metrics.afterProductionAvailableTime)),
    metrics.beforeAU,
    metrics.afterAU,
    metrics.beforeUPDT,
    metrics.afterUPDT,
    metrics.updtGap,
    _pme_getYearMonth(plan.pmDate),
    subprocess || "",
    metrics.productionStatus || "",
    metrics.performanceStatus || "",
  ];
}

/**
 * 判断 Master Data 行是否具备 Good/Bad KPI 资格。
 * 排除规则由 _pme_isKpiExcluded 在计划层单独处理。
 * @param {Array<*>} row
 * @returns {boolean}
 */
function _pme_calculateKpiEligibility(row) {
  if (!Array.isArray(row)) return false;
  return row[13] === "生产" && (row[14] === "好/Good" || row[14] === "差/Bad");
}

function _pme_emptyKpiBucket() {
  return { good: 0, bad: 0, total: 0 };
}

function _pme_normalizeKpiMonths(months) {
  return Array.from(new Set((months || []).map(function (month) {
    return String(month == null ? "" : month).trim();
  }).filter(function (month) {
    return /^\d{6}$/.test(month);
  })));
}

function _pme_isDirectKpiExcludedRow(row) {
  if (!Array.isArray(row)) return false;
  var process = String(row[1] == null ? "" : row[1]).trim().toUpperCase();
  var workcenter = _pme_normalizeWorkcenter(row[3]);
  if (/^S1DP000[1-4]$/.test(workcenter)) return true;
  return process === "PK" && /^PK[SF]/.test(workcenter.toUpperCase());
}

function _pme_incrementKpiBucket(summary, process, yearMonth, performanceStatus) {
  if (!summary[process] || !summary[process][yearMonth]) return;
  var bucket = summary[process][yearMonth];
  if (performanceStatus === "好/Good") bucket.good++;
  if (performanceStatus === "差/Bad") bucket.bad++;
  bucket.total = bucket.good + bucket.bad;
}

/**
 * 汇总 Master Data 的 PME 结果。传入的 Master Data 行不含表头。
 * @param {Array<Array<*>>} masterRows
 * @param {Array<string>} months
 * @param {Set<string>=} excludedSerials
 * @returns {Object<string, Object<string, {good:number,bad:number,total:number}>>}
 */
function _pme_calculateKpiSummary(masterRows, months, excludedSerials) {
  var selectedMonths = _pme_normalizeKpiMonths(months);
  var summary = {};
  Object.keys(_pme_KPI_PROCESS_CONFIG).forEach(function (process) {
    summary[process] = {};
    selectedMonths.forEach(function (month) {
      summary[process][month] = _pme_emptyKpiBucket();
    });
  });

  (masterRows || []).forEach(function (row) {
    if (!_pme_calculateKpiEligibility(row)) return;
    var yearMonth = String(row[11] == null ? "" : row[11]).trim();
    var keyProcess = String(row[1] == null ? "" : row[1]).trim();
    var serialNo = String(row[0] == null ? "" : row[0]).trim();
    if (selectedMonths.indexOf(yearMonth) === -1) return;
    if (excludedSerials && typeof excludedSerials.has === "function" && excludedSerials.has(serialNo)) return;
    if (_pme_isDirectKpiExcludedRow(row)) return;
    if (!_pme_KPI_PROCESS_CONFIG[keyProcess]) return;

    _pme_incrementKpiBucket(summary, keyProcess, yearMonth, row[14]);
    if (keyProcess === "TF") {
      var subprocess = String(row[12] == null ? "" : row[12]).trim();
      if (_pme_KPI_PROCESS_CONFIG[subprocess]) {
        _pme_incrementKpiBucket(summary, subprocess, yearMonth, row[14]);
      }
    }
  });

  return summary;
}

function _pme_formatKpiPercentage(value) {
  if (typeof value !== "number" || !Number.isFinite(value)) return "0.0%";
  return (value * 100).toFixed(1).replace(/\.0$/, "") + "%";
}

function _pme_parseKpiGoal(value) {
  if (typeof value === "string") {
    var text = value.trim();
    if (!text) return 0;
    var parsedText = parseFloat(text.replace("%", ""));
    return Number.isFinite(parsedText) ? parsedText : 0;
  }
  if (typeof value === "number" && Number.isFinite(value)) return value < 1 ? value * 100 : value;
  return 0;
}

function _pme_getKpiRow(process, year) {
  var config = _pme_KPI_PROCESS_CONFIG[process];
  if (!config) return null;
  return year === 2025 ? config.row2025 : year === 2026 ? config.row2026 : null;
}

function _pme_getKpiMonthColumn(yearMonth) {
  var text = String(yearMonth || "");
  if (!/^\d{6}$/.test(text)) return null;
  var month = Number(text.slice(4, 6));
  return month >= 1 && month <= 12 ? month + 3 : null;
}

function _pme_writeKpiCell(sheet, row, column, bucket, goal) {
  var value = bucket && bucket.total > 0
    ? _pme_formatKpiPercentage(bucket.good / bucket.total)
    : "-";
  var range = sheet.getRange(row, column);
  range.setValue(value);
  if (value === "-") {
    range.setBackground(null);
    return value;
  }

  var actual = parseFloat(value.replace("%", ""));
  range.setBackground(actual >= goal ? "#90EE90" : "#FFB6C1");
  return value;
}

/**
 * 将 PME 汇总写入既有 KPI Summary 布局。
 * @param {Object} summary
 * @param {Array<string>} months
 * @param {Object} sheet
 */
function _pme_writeKpiSummary(summary, months, sheet) {
  if (!sheet) throw new Error("KPI Summary Sheet 不存在");
  var selectedMonths = _pme_normalizeKpiMonths(months);
  var years = Array.from(new Set(selectedMonths.map(function (month) { return Number(month.slice(0, 4)); })));

  years.forEach(function (year) {
    Object.keys(_pme_KPI_PROCESS_CONFIG).forEach(function (process) {
      var row = _pme_getKpiRow(process, year);
      if (!row) return;
      var goal = _pme_parseKpiGoal(sheet.getRange(row, 3).getValue());

      selectedMonths.filter(function (month) { return Number(month.slice(0, 4)) === year; })
        .forEach(function (month) {
          var column = _pme_getKpiMonthColumn(month);
          if (!column) return;
          _pme_writeKpiCell(sheet, row, column,
            summary[process] && summary[process][month], goal);
        });

      var ytd = _pme_emptyKpiBucket();
      selectedMonths.filter(function (month) { return Number(month.slice(0, 4)) === year; })
        .forEach(function (month) {
          var bucket = summary[process] && summary[process][month];
          if (!bucket) return;
          ytd.good += bucket.good;
          ytd.bad += bucket.bad;
          ytd.total += bucket.total;
        });
      _pme_writeKpiCell(sheet, row, 16, ytd, goal);
    });
  });
}

function _pme_processMatches(planProcess, aemProcess) {
  return String(planProcess || "").trim().toUpperCase() ===
    String(aemProcess || "").trim().toUpperCase();
}

function _pme_aemPlanIndexKey(workcenter) {
  return _pme_normalizeWorkcenter(workcenter);
}

/**
 * 只遍历一次 AEM，建立机台号索引，避免每条 PM 计划重复扫描全量 AEM。
 * @param {Array<Array<*>>} aemRows 含表头行
 * @param {Object<string, number>} headerIndex
 * @param {Set<string>=} eligibleWorkcenters 有效 PM Plan 涉及的机台号
 * @returns {Object<string, Array<Array<*>>>}
 */
function _pme_buildAemPlanIndex(aemRows, headerIndex, eligibleWorkcenters) {
  var index = {};
  if (!Array.isArray(aemRows)) return index;

  for (var rowIndex = 1; rowIndex < aemRows.length; rowIndex++) {
    var row = aemRows[rowIndex] || [];
    var rawWorkcenter = String(_pme_getCell(row, headerIndex, "Workcenter") || "").trim();
    if (_pme_hasSwWorkcenter(rawWorkcenter)) continue;
    var workcenter = _pme_normalizeWorkcenter(rawWorkcenter);
    var process = String(_pme_getCell(row, headerIndex, "process") || "").trim().toUpperCase();
    var workcenterGroup = String(_pme_getCell(row, headerIndex, "WorkcenterGroup") || "").trim();
    if (!workcenter || !process) continue;
    if (eligibleWorkcenters && typeof eligibleWorkcenters.has === "function" &&
        !eligibleWorkcenters.has(workcenter)) continue;
    if (process === "IM" && (workcenterGroup === "OPK" || workcenterGroup === "Dummy")) continue;

    var key = _pme_aemPlanIndexKey(workcenter);
    if (!index[key]) index[key] = [];
    index[key].push(row);
  }
  return index;
}

function _pme_getAemRowsForPlan(plan, aemPlanIndex, headerRow, headerIndex) {
  var key = _pme_aemPlanIndexKey(plan.workcenter);
  var rows = aemPlanIndex && aemPlanIndex[key] ? aemPlanIndex[key] : [];
  if (headerIndex && headerIndex.process !== undefined) {
    rows = rows.filter(function (row) {
      return _pme_processMatches(plan.keyProcess, _pme_getCell(row, headerIndex, "process"));
    });
  }
  return [headerRow].concat(rows);
}

function _pme_findAemPmRecord(plan, aemRows, headerIndex) {
  if (!Array.isArray(aemRows)) return null;
  for (var rowIndex = 1; rowIndex < aemRows.length; rowIndex++) {
    var row = aemRows[rowIndex] || [];
    var rowDate = _pme_dateKey(_pme_getCell(row, headerIndex, "ProdDate"));
    var rowWorkcenter = _pme_normalizeWorkcenter(_pme_getCell(row, headerIndex, "Workcenter"));
    var rowProcess = _pme_getCell(row, headerIndex, "process");
    if (rowDate === plan.pmDate && rowWorkcenter === plan.workcenter &&
        _pme_processMatches(plan.keyProcess, rowProcess) &&
        _pme_number(_pme_getCell(row, headerIndex, "PM")) > 0) {
      return row;
    }
  }
  return null;
}

function _pme_filterAemRowsForPlan(plan, aemRows, headerIndex) {
  var filtered = [aemRows[0]];
  for (var rowIndex = 1; rowIndex < aemRows.length; rowIndex++) {
    var row = aemRows[rowIndex] || [];
    var workcenter = _pme_normalizeWorkcenter(_pme_getCell(row, headerIndex, "Workcenter"));
    var process = _pme_getCell(row, headerIndex, "process");
    var workcenterGroup = String(_pme_getCell(row, headerIndex, "WorkcenterGroup") || "").trim();
    if (workcenter !== plan.workcenter || !_pme_processMatches(plan.keyProcess, process)) continue;
    if (plan.keyProcess === "IM" && (workcenterGroup === "OPK" || workcenterGroup === "Dummy")) continue;
    filtered.push(row);
  }
  return filtered;
}

/**
 * 根据已读取的标准化来源计算 Master Data 行。
 * 返回数组本身，并在数组属性 warnings 中附带数据质量告警，保持接口为 Array<Array>。
 * @param {{pmPlans: Array<Object>, aemRows: Array<Array<*>>, aemHeaderIndex: Object, exclusionRules: Array<Object>, tfTypeMap: Object}} source
 * @returns {Array<Array<*>>}
 */
function _pme_calculateMasterRows(source) {
  var result = [];
  var warnings = [];
  var unmatchedPlanSerials = [];
  var aemRows = source && source.aemRows || [];
  var headerIndex = source && source.aemHeaderIndex || _pme_headerIndex(aemRows[0] || []);
  var aemPlanIndex = source && source.aemPlanIndex || _pme_buildAemPlanIndex(aemRows, headerIndex);
  var headerRow = aemRows[0] || [];

  (source && source.pmPlans || []).forEach(function (plan) {
    if (_pme_isMasterDataExcluded(plan)) return;

    var planRows = _pme_getAemRowsForPlan(plan, aemPlanIndex, headerRow, headerIndex);
    if (!_pme_findAemPmRecord(plan, planRows, headerIndex)) {
      warnings.push("未找到对应 AEM PM 记录: " + plan.serialNo);
      unmatchedPlanSerials.push(plan.serialNo);
      return;
    }

    var shifts = _pme_aggregateAemShifts(planRows, headerIndex);
    var window = _pme_getMaintenanceWindow(plan);
    var beforeShifts = _pme_selectBeforeShifts(plan.keyProcess, shifts, window.startShiftKey);
    var afterWindow = _pme_selectAfterWindow(plan.keyProcess, shifts, window.afterStartShiftKey);
    var metrics = _pme_calculateMetrics(beforeShifts, afterWindow.rows);
    var subprocess = plan.keyProcess === "TF"
      ? String((source.tfTypeMap || {})[plan.workcenter] || "")
      : "";
    var excluded = _pme_isKpiExcluded(plan, source.exclusionRules || []);
    result.push(_pme_buildMasterRow(plan, metrics, subprocess, { excluded: excluded }));
  });

  result.warnings = warnings;
  result.unmatchedPlanSerials = unmatchedPlanSerials;
  return result;
}

/**
 * 按序列号幂等更新 Master Data。输入和输出均含表头行。
 * @param {Array<Array<*>>} existingRows
 * @param {Array<Array<*>>} calculatedRows 不含表头
 * @param {Array<Object>=} pmPlans 当前执行年份的有效 PM Plan
 * @param {number=} targetYear 当前执行年份；仅清理该年份的失效 Master Data 行
 * @param {Array<string>=} unmatchedPlanSerials 当前有效计划中未匹配到 AEM PM 记录的序列号
 * @returns {Array<Array<*>>}
 */
function _pme_upsertMasterData(existingRows, calculatedRows, pmPlans, targetYear, unmatchedPlanSerials) {
  var current = Array.isArray(existingRows) && existingRows.length > 0
    ? existingRows.map(function (row) { return row.slice(); })
    : [_pme_MASTER_HEADERS.slice()];
  var header = current[0] || [];
  _pme_validateMasterDataHeaders(header);
  var width = header.length;
  var masterHeaderIndex = _pme_headerIndex(header);
  var serialColumn = masterHeaderIndex["序列号/ Serial No."];
  var pmDateColumn = masterHeaderIndex["保养日期/ PM_Date"];

  if (targetYear != null && Array.isArray(pmPlans)) {
    var targetYearText = String(targetYear).trim();
    var activePlanSerials = new Set();
    pmPlans.forEach(function (plan) {
      if (_pme_isMasterDataExcluded(plan)) return;
      var serialNo = String(plan && plan.serialNo || "").trim();
      if (serialNo) activePlanSerials.add(serialNo);
    });

    current = [current[0]].concat(current.slice(1).filter(function (row) {
      var pmDate = _pme_dateKey(row[pmDateColumn]);
      if (pmDate.slice(0, 4) !== targetYearText) return true;
      return activePlanSerials.has(String(row[serialColumn] || "").trim());
    }));
  }

  var unmatchedSerialSet = new Set(unmatchedPlanSerials || []);
  var productionStatusColumn = masterHeaderIndex["生产状态/ Production Status"];
  if (productionStatusColumn !== undefined && unmatchedSerialSet.size > 0) {
    for (var unmatchedRowIndex = 1; unmatchedRowIndex < current.length; unmatchedRowIndex++) {
      var unmatchedSerial = String(current[unmatchedRowIndex][serialColumn] || "").trim();
      if (unmatchedSerialSet.has(unmatchedSerial)) {
        current[unmatchedRowIndex][productionStatusColumn] = "未匹配到保养记录";
      }
    }
  }

  var rowIndexBySerial = {};
  for (var rowIndex = 1; rowIndex < current.length; rowIndex++) {
    var serialNo = String(current[rowIndex][serialColumn] || "").trim();
    if (serialNo && rowIndexBySerial[serialNo] === undefined) rowIndexBySerial[serialNo] = rowIndex;
  }

  (calculatedRows || []).forEach(function (row) {
    var serialNo = String(row && row[0] || "").trim();
    if (!serialNo) return;
    if (rowIndexBySerial[serialNo] === undefined) {
      rowIndexBySerial[serialNo] = current.length;
      current.push(_pme_mergeMasterDataRow(row, null, width));
    } else {
      var rowIndex = rowIndexBySerial[serialNo];
      current[rowIndex] = _pme_mergeMasterDataRow(row, current[rowIndex], width);
    }
  });
  return current;
}

function _pme_validateMasterDataHeaders(headers) {
  if (!Array.isArray(headers) || headers.length < _pme_MASTER_HEADERS.length ||
      _pme_MASTER_HEADERS.some(function (name, index) { return headers[index] !== name; })) {
    throw new Error("Master Data 表头前 15 列与 PME 结构不一致");
  }
  return true;
}

function _pme_mergeMasterDataRow(calculatedRow, existingRow, width) {
  var merged = (calculatedRow || []).slice(0, _pme_MASTER_HEADERS.length);
  while (merged.length < _pme_MASTER_HEADERS.length) merged.push("");
  for (var column = _pme_MASTER_HEADERS.length; column < width; column++) {
    merged.push(existingRow && existingRow[column] !== undefined ? existingRow[column] : "");
  }
  return merged;
}

/**
 * 追加大量行时使用逐行 push，避免将整批行展开为函数参数导致栈溢出。
 * @param {Array<Array<*>>} target
 * @param {Array<Array<*>>} additions
 */
function _pme_appendRows(target, additions) {
  if (!Array.isArray(target) || !Array.isArray(additions)) return target;
  for (var index = 0; index < additions.length; index++) {
    target.push(additions[index]);
  }
  return target;
}

function _pme_readSheetRows(sheet, chunkRows) {
  if (!sheet) throw new Error("目标 Sheet 不存在");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];

  var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var rows = [header];
  var size = chunkRows || _pme_CONFIG.sourceChunkRows;
  for (var startRow = 2; startRow <= lastRow; startRow += size) {
    var count = Math.min(size, lastRow - startRow + 1);
    var chunk = sheet.getRange(startRow, 1, count, lastColumn).getValues();
    _pme_appendRows(rows, chunk);
  }
  return rows;
}

function _pme_filterAemRowsBySourceRange(rows, headerIndex, source) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("AEM " + String(source && source.label || "来源") + " SourceData 为空");
  }

  var start = _pme_dateKey(source && source.start);
  var end = source && source.end ? _pme_dateKey(source.end) : "";
  var dateColumn = headerIndex && headerIndex.ProdDate;
  var filtered = rows.slice(1).filter(function (row) {
    var dateKey = _pme_dateKey(dateColumn === undefined ? "" : row[dateColumn]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return false;
    if (start && dateKey < start) return false;
    if (end && dateKey > end) return false;
    return true;
  });

  if (filtered.length === 0) {
    throw new Error("AEM " + String(source && source.label || "来源") + " 在配置日期范围内没有有效数据");
  }
  return [rows[0]].concat(filtered);
}

function _pme_filterAemRowsByWorkcenters(rows, headerIndex, eligibleWorkcenters) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  if (!eligibleWorkcenters || typeof eligibleWorkcenters.has !== "function") return rows;

  var filtered = rows.slice(1).filter(function (row) {
    var rawWorkcenter = String(_pme_getCell(row, headerIndex, "Workcenter") || "").trim();
    if (_pme_hasSwWorkcenter(rawWorkcenter)) return false;
    var workcenter = _pme_normalizeWorkcenter(rawWorkcenter);
    return workcenter && eligibleWorkcenters.has(workcenter);
  });
  return [rows[0]].concat(filtered);
}

function _pme_loadExclusionRules(mainSpreadsheet) {
  var sheet = mainSpreadsheet.getSheetByName(_pme_CONFIG.exclusionSheetName);
  if (!sheet || sheet.getLastRow() < 1) return [];
  var rows = _pme_readSheetRows(sheet);
  var index = _pme_requireHeaders(rows[0], ["机台号/Workcenter", "排除的保养时间(h)", "状态"], "KPI排除条件");
  return rows.slice(1).filter(function (row) {
    return String(row[index["状态"]] || "").trim() === "启用";
  }).map(function (row) {
    return {
      workcenter: _pme_normalizeWorkcenter(row[index["机台号/Workcenter"]]),
      duration: _pme_number(row[index["排除的保养时间(h)"]]),
      status: "启用",
    };
  }).filter(function (rule) { return rule.workcenter && rule.duration >= 0; });
}

function _pme_loadTfTypeMap(mainSpreadsheet) {
  var sheet = mainSpreadsheet.getSheetByName(_pme_CONFIG.tfTypeSheetName);
  if (!sheet || sheet.getLastRow() < 1) return {};
  var rows = _pme_readSheetRows(sheet);
  var headers = rows[0] || [];
  var index = _pme_headerIndex(headers);
  var workcenterColumn = index.Workcenter;
  if (workcenterColumn === undefined) workcenterColumn = index["机台号"];
  var typeColumn = index["Machine Type_修正"];
  if (typeColumn === undefined) typeColumn = index["Machine Type"];
  if (workcenterColumn === undefined || typeColumn === undefined) return {};

  var result = {};
  rows.slice(1).forEach(function (row) {
    var workcenter = _pme_normalizeWorkcenter(row[workcenterColumn]);
    if (workcenter) result[workcenter] = String(row[typeColumn] || "").trim();
  });
  return result;
}

function _pme_readAemSources(eligibleWorkcenters) {
  _pme_validateSourceCoverage(_pme_CONFIG.aemSources);
  var allRows = [];
  var sourceHeaders = null;
  var requiredHeaders = [
    "ProdDate", "Workcenter", "Shift", "Prod_Hours", "NetRuntime",
    "PlantAvailableTime", "PM", "process", "WorkcenterGroup", "CO",
  ];

  _pme_CONFIG.aemSources.forEach(function (source) {
    var spreadsheet = SpreadsheetApp.openById(source.id);
    var sheet = spreadsheet.getSheetByName(_pme_CONFIG.sourceSheetName);
    if (!sheet) throw new Error("AEM 来源缺少 SourceData: " + source.label);
    var rows = _pme_readSheetRows(sheet);
    if (rows.length === 0) throw new Error("AEM " + source.label + " SourceData 为空");

    var headerIndex = _pme_requireHeaders(rows[0], requiredHeaders, "AEM " + source.label + " SourceData");
    var sourceRows = _pme_filterAemRowsBySourceRange(rows, headerIndex, source);
    sourceRows = _pme_filterAemRowsByWorkcenters(sourceRows, headerIndex, eligibleWorkcenters);
    if (!sourceHeaders) {
      sourceHeaders = rows[0].slice();
      allRows.push(sourceHeaders);
    } else if (rows[0].length !== sourceHeaders.length ||
               rows[0].some(function (name, index) { return name !== sourceHeaders[index]; })) {
      throw new Error("AEM 来源表头不一致: " + source.label);
    }
    void headerIndex;
    _pme_appendRows(allRows, sourceRows.slice(1));
  });

  return {
    rows: allRows,
    headerIndex: _pme_headerIndex(sourceHeaders || []),
  };
}

/**
 * 读取 PME 计算所需的全部来源。
 * @returns {{pmPlans: Array<Object>, aemRows: Array<Array<*>>, aemHeaderIndex: Object, exclusionRules: Array<Object>, tfTypeMap: Object}}
 */
function _pme_buildEligibleWorkcenters(pmPlans) {
  var workcenters = new Set();
  (pmPlans || []).forEach(function (plan) {
    var workcenter = _pme_normalizeWorkcenter(plan && plan.workcenter);
    if (workcenter) workcenters.add(workcenter);
  });
  return workcenters;
}

function _pme_loadSourceData(targetYear, asOfDate) {
  var mainSpreadsheet = SpreadsheetApp.openById(_pme_CONFIG.mainSpreadsheetId);
  var pmSheet = mainSpreadsheet.getSheetByName(_pme_CONFIG.pmSheetName);
  var pmRows = _pme_readSheetRows(pmSheet);
  var pmHeaderIndex = _pme_requireHeaders(
    pmRows[0] || [],
    ["类别", "状态", "开始日期", "开始时间", "AEM#", "保养时间", "工序"],
    "PM Data",
  );
  var pmPlans = _pme_selectPmPlans(pmRows, pmHeaderIndex, targetYear, asOfDate);
  var eligibleWorkcenters = _pme_buildEligibleWorkcenters(pmPlans);
  var aem = _pme_readAemSources(eligibleWorkcenters);

  return {
    pmPlans: pmPlans,
    aemRows: aem.rows,
    aemHeaderIndex: aem.headerIndex,
    exclusionRules: _pme_loadExclusionRules(mainSpreadsheet),
    tfTypeMap: _pme_loadTfTypeMap(mainSpreadsheet),
  };
}

function _pme_writeMasterDataRows(sheet, rows) {
  if (!sheet) throw new Error("Master Data Sheet 不存在");
  _pme_validateMasterDataHeaders(rows[0] || []);
  var width = _pme_MASTER_HEADERS.length;
  var values = (rows || []).map(function (row) {
    return _pme_mergeMasterDataRow(row, null, width);
  });
  var existingLastRow = typeof sheet.getLastRow === "function" ? sheet.getLastRow() : values.length;
  if (existingLastRow > values.length && typeof sheet.deleteRows === "function") {
    sheet.deleteRows(values.length + 1, existingLastRow - values.length);
  }
  sheet.getRange(1, 1, values.length, width).setValues(values);
  var dataRowCount = values.length - 1;
  if (dataRowCount > 0) {
    sheet.getRange(2, 5, dataRowCount, 2).setNumberFormat("0");
    sheet.getRange(2, 11, dataRowCount, 1).setNumberFormat("0.00%");
  }
}

function _pme_buildExcludedSerials(pmPlans, exclusionRules) {
  var excluded = new Set();
  (pmPlans || []).forEach(function (plan) {
    if (_pme_isKpiExcluded(plan, exclusionRules || [])) excluded.add(plan.serialNo);
  });
  return excluded;
}

function _pme_generateCurrentMonths() {
  var now = new Date();
  var year = now.getFullYear();
  var currentMonth = now.getMonth() + 1;
  var months = [];
  for (var month = 1; month <= currentMonth; month++) {
    months.push(String(year) + String(month).padStart(2, "0"));
  }
  return months;
}

function _pme_getExecutionMonths(event) {
  if (event && Array.isArray(event.months) && event.months.length > 0) {
    return _pme_normalizeKpiMonths(event.months);
  }
  return _pme_generateCurrentMonths();
}

function _pme_getExecutionYear(event) {
  var requestedYear = event && event.targetYear;
  if (/^\d{4}$/.test(String(requestedYear == null ? "" : requestedYear).trim())) {
    return Number(requestedYear);
  }
  return new Date().getFullYear();
}

function _pme_normalizeTriggerLabel(trigger) {
  var raw = trigger == null ? "" : String(trigger);
  if (raw === "scheduled" || raw === "定时") return "定时";
  if (raw === "manual" || raw === "手动" || raw === "") return "手动";
  return raw;
}

function _pme_writeExecutionLog(status, detail, trigger) {
  if (typeof writeLog !== "function") return;
  try {
    writeLog("executePMECalculation", status, detail, _pme_normalizeTriggerLabel(trigger), "");
  } catch (error) {
    console.warn("PME 写入 Log 失败: " + error.message);
  }
}

/**
 * PME 入口：完成来源读取、计算、Master Data 更新和 KPI Summary 汇总。
 * @param {{triggerType?: string}=} event
 * @returns {{trigger: string, planCount: number, calculatedCount: number, warningCount: number, months: Array<string>}}
 */
function executePMECalculation(event) {
  var trigger = _pme_normalizeTriggerLabel(event && event.triggerType ? String(event.triggerType) : "manual");
  try {
    var targetYear = _pme_getExecutionYear(event);
    var source = _pme_loadSourceData(targetYear);
    var calculatedRows = _pme_calculateMasterRows(source);
    var mainSpreadsheet = SpreadsheetApp.openById(_pme_CONFIG.mainSpreadsheetId);
    var masterDataSheet = mainSpreadsheet.getSheetByName(_pme_CONFIG.masterDataSheetName);
    var kpiSummarySheet = mainSpreadsheet.getSheetByName(_pme_CONFIG.kpiSummarySheetName);
    if (!masterDataSheet) throw new Error("Master Data Sheet 不存在");
    if (!kpiSummarySheet) throw new Error("KPI Summary Sheet 不存在");

    var existingRows = _pme_readSheetRows(masterDataSheet);
    var mergedRows = _pme_upsertMasterData(
      existingRows,
      calculatedRows,
      source.pmPlans,
      targetYear,
      calculatedRows.unmatchedPlanSerials,
    );
    var months = _pme_getExecutionMonths(event);
    var excludedSerials = _pme_buildExcludedSerials(source.pmPlans, source.exclusionRules);
    var summary = _pme_calculateKpiSummary(mergedRows.slice(1), months, excludedSerials);

    _pme_writeMasterDataRows(masterDataSheet, mergedRows);
    _pme_writeKpiSummary(summary, months, kpiSummarySheet);

    var result = {
      trigger: trigger,
      planCount: source.pmPlans.length,
      calculatedCount: calculatedRows.length,
      warningCount: (calculatedRows.warnings || []).length,
      months: months,
    };
    _pme_writeExecutionLog(
      "成功",
      "计划 " + result.planCount + " 条，计算 " + result.calculatedCount + " 条，告警 " + result.warningCount + " 条",
      trigger,
    );
    return result;
  } catch (error) {
    _pme_writeExecutionLog("失败", error.message, trigger);
    throw error;
  }
}
