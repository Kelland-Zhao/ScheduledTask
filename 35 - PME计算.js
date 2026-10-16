// PME（Post-PM Effectiveness）计算任务
// 计算层与 Spreadsheet I/O 分离；内部函数统一使用 _pme_ 前缀。

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
function _pme_selectPmPlans(rows, headerIndex) {
  if (!Array.isArray(rows) || rows.length <= 1) return [];

  var plans = [];
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    var row = rows[rowIndex] || [];
    var category = String(_pme_getCell(row, headerIndex, "类别") || "").trim();
    var status = String(_pme_getCell(row, headerIndex, "状态") || "").trim();
    if (category.indexOf("PM") === -1 || status === "取消") continue;

    var pmDate = _pme_dateKey(_pme_getCell(row, headerIndex, "开始日期"));
    var workcenter = _pme_normalizeWorkcenter(_pme_getCell(row, headerIndex, "AEM#"));
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
  if (String(process || "").toUpperCase() !== "IM") return positive;

  var skipped = {};
  if (positive.length > 0) skipped[positive[positive.length - 1].shiftKey] = true;

  for (var index = 1; index < before.length; index++) {
    var previous = before[index - 1];
    var current = before[index];
    if (previous.prodHours > 0 && current.prodHours === 0) {
      skipped[previous.shiftKey] = true;
    }
  }

  return positive.filter(function (shift) { return !skipped[shift.shiftKey]; });
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
    metrics.beforeProductionAvailableTime,
    metrics.afterProductionAvailableTime,
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
