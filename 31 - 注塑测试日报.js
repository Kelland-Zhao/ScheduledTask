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

function _itr_addDays(date, days) {
  var result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

function _itr_dateKey(date) {
  return Utilities.formatDate(date, ITR_CONFIG.TIME_ZONE, "yyyy-MM-dd");
}

function _itr_getRichTextLink(richTextValue) {
  if (!richTextValue) return "";

  var directLink = richTextValue.getLinkUrl();
  if (directLink) return directLink;

  var runs = richTextValue.getRuns() || [];
  for (var i = 0; i < runs.length; i++) {
    var runLink = runs[i].getLinkUrl();
    if (runLink) return runLink;
  }
  return "";
}

function _itr_getRecordsByDate(sheet, targetDateKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var range = sheet.getRange(2, 1, lastRow - 1, 21);
  var values = range.getValues();
  var richTextValues = range.getRichTextValues();
  var records = [];

  values.forEach(function (rowValues, index) {
    if (_itr_normalizeDate(rowValues[1]) !== targetDateKey) return;

    var richTextRow = richTextValues[index] || [];
    records.push({
      rowNumber: index + 2,
      values: rowValues,
      links: {
        sample: _itr_getRichTextLink(richTextRow[16]),
        record: _itr_getRichTextLink(richTextRow[17])
      }
    });
  });
  return records;
}

function _itr_fillSmartChipLinks(records, sheetName) {
  var escapedSheetName = "'" + String(sheetName).replace(/'/g, "''") + "'";
  var linkFields = [
    { valueIndex: 16, column: "Q", key: "sample" },
    { valueIndex: 17, column: "R", key: "record" }
  ];

  (records || []).forEach(function (record) {
    record.links = record.links || {};
    record.links.sample = record.links.sample || "";
    record.links.record = record.links.record || "";
    linkFields.forEach(function (field) {
      if (!String(record.values[field.valueIndex] || "").trim() ||
          record.links[field.key]) {
        return;
      }

      var range = escapedSheetName + "!" + field.column + record.rowNumber;
      try {
        var response = Sheets.Spreadsheets.get(ITR_CONFIG.SPREADSHEET_ID, {
          ranges: [range],
          includeGridData: true,
          fields: "sheets.data.rowData.values(chipRuns)"
        });
        var sheets = response.sheets || [];
        var data = sheets[0] && sheets[0].data || [];
        var rowData = data[0] && data[0].rowData || [];
        var cells = rowData[0] && rowData[0].values || [];
        var chipRuns = cells[0] && cells[0].chipRuns || [];
        var uri = "";
        for (var i = 0; i < chipRuns.length; i++) {
          var properties = chipRuns[i].chip && chipRuns[i].chip.richLinkProperties;
          if (properties && properties.uri) {
            uri = properties.uri;
            break;
          }
        }
        if (uri) record.links[field.key] = uri;
      } catch (error) {
        console.warn("Smart Chip link read failed for " + range + ": " + error.message);
      }
    });
  });
  return records;
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

function _itr_getRecipients(spreadsheet, records) {
  var notificationSheet = spreadsheet.getSheetByName(ITR_CONFIG.NOTIFICATION_SHEET);
  if (!notificationSheet) {
    throw new Error("缺少必需工作表：通知清单");
  }

  var nameSheet = spreadsheet.getSheetByName(ITR_CONFIG.NAME_SHEET);
  if (!nameSheet) {
    throw new Error("缺少必需工作表：Name Database");
  }

  var emails = [];
  var notificationRows = notificationSheet.getDataRange().getValues();
  notificationRows.forEach(function (notificationRow) {
    var notificationType = String(notificationRow[1] || "");
    if (notificationType.indexOf("测试后日报") !== -1 ||
        notificationType.indexOf("测试前日报") !== -1) {
      emails.push(notificationRow[0]);
    }
  });

  var ownerNames = {};
  (records || []).forEach(function (record) {
    var ownerName = String(record.values[12] || "").trim();
    if (ownerName) ownerNames[ownerName] = true;
  });

  var nameRows = nameSheet.getDataRange().getValues();
  nameRows.forEach(function (nameRow) {
    var name = String(nameRow[0] || "").trim();
    if (ownerNames[name]) emails.push(nameRow[1]);
  });

  return _itr_uniqueEmails(emails);
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
