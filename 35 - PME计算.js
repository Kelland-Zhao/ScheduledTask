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
