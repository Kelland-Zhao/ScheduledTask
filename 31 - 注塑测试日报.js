var ITR_CONFIG = {
  SPREADSHEET_ID: "17ys3UDFWjhfaPnk0TErqqeU0FnMP7nsRoRsTmlmm2fg",
  TEST_SHEET: "注塑测试",
  NOTIFICATION_SHEET: "通知清单",
  NAME_SHEET: "Name Database",
  TIME_ZONE: "Asia/Shanghai",
  TEST_EMAIL: "kelland_zhao@colpal.com",
  FROM_EMAIL: "CSX_PlantSystem@colpal.com"
};

function _itr_normalizeDate(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return isNaN(value.getTime())
      ? ""
      : Utilities.formatDate(value, ITR_CONFIG.TIME_ZONE, "yyyy-MM-dd");
  }

  var text = String(value).trim();
  var match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (match) {
    return match[1] + "-" +
      ("0" + match[2]).slice(-2) + "-" +
      ("0" + match[3]).slice(-2);
  }

  var parsed = new Date(text);
  return isNaN(parsed.getTime())
    ? ""
    : Utilities.formatDate(parsed, ITR_CONFIG.TIME_ZONE, "yyyy-MM-dd");
}

function _itr_isSkippedStatus(status) {
  return ["满产", "模具不在线", "取消"].indexOf(String(status || "").trim()) !== -1;
}

function _itr_checkAfterProblems(record) {
  var status = String(record.values[15] || "").trim();
  if (_itr_isSkippedStatus(status)) return [];
  if (!status) return ["完成状态未维护"];
  if (status === "已完成" && !String(record.values[17] || "").trim()) {
    return ["无测试记录"];
  }
  return [];
}

function _itr_checkBeforeProblems(record) {
  var status = String(record.values[15] || "").trim();
  if (_itr_isSkippedStatus(status)) return [];

  var problems = [];
  if (!String(record.values[11] || "").trim()) {
    problems.push("测试机台未维护");
  }
  if (!String(record.values[14] || "").trim()) {
    problems.push("测试负责人未维护");
  }
  if (!String(record.values[16] || "").trim()) {
    problems.push("测试样单链接未维护");
  }
  return problems;
}

function _itr_uniqueEmails(emails) {
  var seen = {};
  return (emails || []).reduce(function (result, email) {
    var normalized = String(email || "").trim().toLowerCase();
    var valid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized);
    if (valid && !seen[normalized]) {
      seen[normalized] = true;
      result.push(normalized);
    }
    return result;
  }, []);
}

function _itr_escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
