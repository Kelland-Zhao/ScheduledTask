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
  return _itr_getRecordsByDateKeys(sheet, [targetDateKey])[targetDateKey] || [];
}

function _itr_getRecordsByDateKeys(sheet, targetDateKeys) {
  var grouped = {};
  (targetDateKeys || []).forEach(function (key) {
    grouped[key] = [];
  });
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return grouped;

  var range = sheet.getRange(2, 1, lastRow - 1, 21);
  var values = range.getValues();
  var richTextValues = range.getRichTextValues();

  values.forEach(function (rowValues, index) {
    var dateKey = _itr_normalizeDate(rowValues[1]);
    if (!Object.prototype.hasOwnProperty.call(grouped, dateKey)) return;

    var richTextRow = richTextValues[index] || [];
    grouped[dateKey].push({
      rowNumber: index + 2,
      values: rowValues,
      links: {
        sample: _itr_getRichTextLink(richTextRow[16]),
        record: _itr_getRichTextLink(richTextRow[17])
      }
    });
  });
  return grouped;
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

function _itr_cell(record, index) {
  return String(record && record.values && record.values[index] || "").trim();
}

function _itr_linkHtml(label, url) {
  var safeLabel = _itr_escapeHtml(label || "查看 / View");
  var linkUrl = String(url || "").trim();
  var labelUrl = String(label || "").trim();
  if (!linkUrl && /^https?:\/\//i.test(labelUrl)) linkUrl = labelUrl;
  if (!/^https?:\/\//i.test(linkUrl)) {
    return safeLabel === "查看 / View" ? "-" : safeLabel;
  }
  return '<a href="' + _itr_escapeHtml(linkUrl) +
    '" style="color:#E60012;text-decoration:none;">' + safeLabel + "</a>";
}

function _itr_statusHtml(status, problems) {
  var hasProblems = (problems || []).length > 0;
  var text = hasProblems
    ? (problems || []).join("；")
    : "正常 / Normal";
  var color = hasProblems ? "#b3261e" : "#137333";
  var background = hasProblems ? "#fce8e6" : "#e6f4ea";
  return '<span style="display:inline-block;padding:3px 8px;border-radius:10px;' +
    "color:" + color + ";background:" + background + ';">' +
    _itr_escapeHtml(text) + "</span>";
}

function _itr_buildSectionHtml(titleCn, titleEn, dateKey, records, sourceUrl) {
  if (!records || records.length === 0) return "";

  var headers = [
    ["序号", "No."], ["产品名称", "Product"], ["测试说明", "Description"],
    ["日期", "Date"], ["机台", "Machine"], ["项目负责人", "Project Owner"],
    ["测试负责人", "Test Owner"], ["状态", "Status"], ["样单", "Sample"],
    ["记录", "Record"], ["检查结果", "Check Result"]
  ];
  var headerHtml = headers.map(function (pair) {
    return '<th style="padding:5px 6px;border:1px solid #ddd;background:#E60012;' +
      'color:#fff;font-size:12px;font-weight:bold;text-align:center;' +
      'vertical-align:middle;">' +
      _itr_escapeHtml(pair[0]) + "<br>" + _itr_escapeHtml(pair[1]) + "</th>";
  }).join("");

  var rowsHtml = records.map(function (record, index) {
    var problems = record.problems || [];
    var rowBackground = problems.length ? "#fce8e6" : "#e6f4ea";
    var rowUrl = String(sourceUrl || "") +
      "#gid=636324977&range=A" + record.rowNumber;
    var cells = [
      _itr_linkHtml(String(index + 1), rowUrl),
      _itr_escapeHtml(_itr_cell(record, 5)),
      _itr_escapeHtml(_itr_cell(record, 6)),
      _itr_escapeHtml(_itr_normalizeDate(_itr_cell(record, 1)) || dateKey),
      _itr_escapeHtml(_itr_cell(record, 11)),
      _itr_escapeHtml(_itr_cell(record, 12)),
      _itr_escapeHtml(_itr_cell(record, 14)),
      _itr_escapeHtml(_itr_cell(record, 15)),
      _itr_linkHtml(_itr_cell(record, 16), record.links && record.links.sample),
      _itr_linkHtml(_itr_cell(record, 17), record.links && record.links.record),
      _itr_statusHtml(_itr_cell(record, 15), problems)
    ];
    return '<tr style="background:' + rowBackground + ';">' +
      cells.map(function (cell) {
        return '<td style="padding:5px 6px;border:1px solid #ddd;' +
          'font-size:12px;vertical-align:middle;text-align:center;">' +
          (cell || "-") + "</td>";
      }).join("") + "</tr>";
  }).join("");

  return '<div style="margin-top:22px;">' +
    '<div style="font-size:13px;font-weight:700;color:#6c757d;' +
    'border-left:3px solid #E60012;padding-left:10px;margin:18px 0 10px;' +
    'letter-spacing:0.5px;">' +
    _itr_escapeHtml(titleCn) + "<br>" +
    '<span style="font-weight:400;">' +
    _itr_escapeHtml(titleEn.toUpperCase()) + "</span>" +
    " · " + _itr_escapeHtml(dateKey) + "</div>" +
    '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;' +
    'font-family:Arial,sans-serif;">' +
    "<thead><tr>" + headerHtml + "</tr></thead><tbody>" + rowsHtml +
    "</tbody></table></div></div>";
}

function _itr_buildEmailHtml(model) {
  var yesterdayRecords = model.yesterday.records || [];
  var tomorrowRecords = model.tomorrow.records || [];
  var allRecords = yesterdayRecords.concat(tomorrowRecords);
  var abnormalCount = allRecords.filter(function (record) {
    return (record.problems || []).length > 0;
  }).length;

  return '<div style="font-family:Arial,Microsoft YaHei,sans-serif;color:#333;' +
    'max-width:1200px;margin:0 auto;">' +
    '<div style="background:#E60012;color:#fff;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h1 style="margin:0;font-size:24px;">注塑测试日报 / Injection Test Daily Report</h1>' +
    '<div style="margin-top:8px;font-size:14px;">日报日期 / Report Date: ' +
    _itr_escapeHtml(model.reportDate) + "</div></div>" +
    '<div style="padding:14px 20px;background:#fff5f5;border:1px solid #f3cccc;">' +
    "昨日 " + yesterdayRecords.length + " · 明日 " + tomorrowRecords.length +
    " · 异常 " + abnormalCount + "</div>" +
    _itr_buildSectionHtml(
      "昨日测试复盘", "Yesterday Review",
      model.yesterday.date,
      yesterdayRecords,
      model.sourceUrl
    ) +
    _itr_buildSectionHtml(
      "明日测试提醒", "Tomorrow Reminder",
      model.tomorrow.date,
      tomorrowRecords,
      model.sourceUrl
    ) +
    '<div style="margin-top:20px;color:#777;font-size:11px;">' +
    "此邮件由 ScheduledScripts 自动生成 / Generated automatically.</div></div>";
}

function _itr_sendEmail(recipients, subject, htmlBody) {
  var options = {
    htmlBody: htmlBody,
    name: "注塑测试日报"
  };
  var aliases = GmailApp.getAliases() || [];
  var configuredAlias = String(ITR_CONFIG.FROM_EMAIL).toLowerCase();
  var aliasAvailable = aliases.some(function (alias) {
    return String(alias).toLowerCase() === configuredAlias;
  });
  if (aliasAvailable) options.from = ITR_CONFIG.FROM_EMAIL;

  GmailApp.sendEmail(
    recipients.join(","),
    subject,
    "请使用支持 HTML 的邮件客户端查看注塑测试日报。",
    options
  );
}

function _itr_run(options) {
  options = options || {};
  var now = options.now || new Date();
  var reportDate = _itr_dateKey(now);
  var yesterdayDate = _itr_dateKey(_itr_addDays(now, -1));
  var tomorrowDate = _itr_dateKey(_itr_addDays(now, 1));
  var spreadsheet = options.spreadsheet ||
    SpreadsheetApp.openById(ITR_CONFIG.SPREADSHEET_ID);
  var testSheet = spreadsheet.getSheetByName(ITR_CONFIG.TEST_SHEET);
  if (!testSheet) {
    throw new Error("缺少必需工作表：" + ITR_CONFIG.TEST_SHEET);
  }

  var recordsByDate = _itr_getRecordsByDateKeys(
    testSheet,
    [yesterdayDate, tomorrowDate]
  );
  var yesterdayRecords = recordsByDate[yesterdayDate] || [];
  var tomorrowRecords = recordsByDate[tomorrowDate] || [];
  if (yesterdayRecords.length === 0 && tomorrowRecords.length === 0) {
    return {
      status: "skipped",
      reportDate: reportDate,
      yesterdayCount: 0,
      tomorrowCount: 0,
      abnormalCount: 0,
      recipients: []
    };
  }

  _itr_fillSmartChipLinks(yesterdayRecords.concat(tomorrowRecords), ITR_CONFIG.TEST_SHEET);
  yesterdayRecords.forEach(function (record) {
    record.problems = _itr_checkAfterProblems(record);
  });
  tomorrowRecords.forEach(function (record) {
    record.problems = _itr_checkBeforeProblems(record);
  });

  var recipients = options.testMode
    ? [ITR_CONFIG.TEST_EMAIL]
    : _itr_getRecipients(spreadsheet, yesterdayRecords.concat(tomorrowRecords));
  recipients = _itr_uniqueEmails(recipients);
  if (recipients.length === 0) {
    throw new Error("无有效收件人");
  }

  var sourceUrl = typeof spreadsheet.getUrl === "function"
    ? spreadsheet.getUrl()
    : "https://docs.google.com/spreadsheets/d/" + ITR_CONFIG.SPREADSHEET_ID + "/edit";
  var model = {
    reportDate: reportDate,
    sourceUrl: sourceUrl,
    yesterday: { date: yesterdayDate, records: yesterdayRecords },
    tomorrow: { date: tomorrowDate, records: tomorrowRecords }
  };
  var subject = (options.testMode ? "【测试】 " : "") +
    "【注塑测试日报】" + reportDate + " 昨日复盘 & 明日提醒";
  _itr_sendEmail(recipients, subject, _itr_buildEmailHtml(model));

  var abnormalCount = yesterdayRecords.concat(tomorrowRecords)
    .filter(function (record) {
      return record.problems.length > 0;
    }).length;
  return {
    status: "sent",
    reportDate: reportDate,
    yesterdayCount: yesterdayRecords.length,
    tomorrowCount: tomorrowRecords.length,
    abnormalCount: abnormalCount,
    recipients: recipients
  };
}

function _itr_logResult(functionName, result, trigger) {
  if (result.status === "skipped") {
    _itr_safeWriteLog(functionName, "跳过", "昨日与明日均无测试记录", trigger, "");
    return;
  }
  var detail = "昨日 " + result.yesterdayCount +
    " 条，明日 " + result.tomorrowCount +
    " 条，异常 " + result.abnormalCount +
    " 条，收件人 " + result.recipients.length + " 人";
  _itr_safeWriteLog(
    functionName,
    "成功",
    detail,
    trigger,
    "TO: " + result.recipients.join(",")
  );
}

function _itr_safeWriteLog(functionName, status, detail, trigger, remark) {
  try {
    writeLog(functionName, status, detail, trigger, remark);
    return true;
  } catch (error) {
    console.error("注塑测试日报日志写入失败: " + error.message);
    return false;
  }
}

function _itr_execute(functionName, trigger, options) {
  try {
    var result = _itr_run(options || {});
    _itr_logResult(functionName, result, trigger);
    return result;
  } catch (error) {
    _itr_safeWriteLog(
      functionName,
      "失败",
      error.message,
      trigger,
      error.stack || ""
    );
    throw error;
  }
}

function sendInjectionTestDailyReport(e) {
  return _itr_execute(
    "sendInjectionTestDailyReport",
    e && e.triggerType === "scheduled" ? "定时" : "手动",
    {}
  );
}

function testInjectionTestDailyReport() {
  return _itr_execute(
    "testInjectionTestDailyReport",
    "测试",
    { testMode: true }
  );
}
