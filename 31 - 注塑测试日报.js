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
  var dateMatch = text.match(/^(\d{4})([\/-])(\d{2})\2(\d{2})$/);
  if (dateMatch) {
    var year = Number(dateMatch[1]);
    var month = Number(dateMatch[3]);
    var day = Number(dateMatch[4]);
    var calendarDate = new Date(Date.UTC(year, month - 1, day));
    if (calendarDate.getUTCFullYear() !== year ||
        calendarDate.getUTCMonth() !== month - 1 ||
        calendarDate.getUTCDate() !== day) {
      return "";
    }
    return dateMatch[1] + "-" + dateMatch[3] + "-" + dateMatch[4];
  }

  var isoMatch = text.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-]\d{2}:\d{2})$/
  );
  if (!isoMatch) return "";

  var isoYear = Number(isoMatch[1]);
  var isoMonth = Number(isoMatch[2]);
  var isoDay = Number(isoMatch[3]);
  var isoHour = Number(isoMatch[4]);
  var isoMinute = Number(isoMatch[5]);
  var isoSecond = Number(isoMatch[6]);
  var isoCalendarDate = new Date(Date.UTC(isoYear, isoMonth - 1, isoDay));
  if (isoCalendarDate.getUTCFullYear() !== isoYear ||
      isoCalendarDate.getUTCMonth() !== isoMonth - 1 ||
      isoCalendarDate.getUTCDate() !== isoDay ||
      isoHour > 23 ||
      isoMinute > 59 ||
      isoSecond > 59) {
    return "";
  }

  var offset = isoMatch[7];
  if (offset !== "Z") {
    var offsetParts = offset.slice(1).split(":");
    if (Number(offsetParts[0]) > 23 || Number(offsetParts[1]) > 59) return "";
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
    var parts = normalized.split("@");
    var local = parts[0] || "";
    var domain = parts[1] || "";
    var localValid = parts.length === 2 &&
      /^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+$/.test(local) &&
      local.charAt(0) !== "." &&
      local.charAt(local.length - 1) !== "." &&
      local.indexOf("..") === -1;
    var labels = domain.split(".");
    var domainValid = labels.length >= 2 && labels.every(function (label) {
      return /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$/.test(label);
    });
    var valid = localValid && domainValid;
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
