// PME 年度归档：整年门槛、PME剪切、KPI快照及邮件提醒

var _pmeBackup_CONFIG = {
  sourceSpreadsheetId: "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24",
  sourceMasterSheetName: "Master Data",
  sourceKpiSheetName: "KPI Summary",
  destinationSpreadsheetId: "1PMZDeHuE-qwiV-s62b6n7HsNUMemTc8YNJmWHEnMqdM",
  destinationKpiSheetName: "KPI",
  recipient: "kelland_zhao@colpal.com",
  timeZone: "Asia/Shanghai",
  headers: {
    serial: "序列号/ Serial No.",
    yearMonth: "保养月份/ YearMonth",
    pmDate: "保养日期/ PM_Date",
    workcenter: "机台号/ Workcenter",
    productionStatus: "生产状态/ Production Status",
    exempt: "豁免与否/ Exempt?",
    exemptionReason: "豁免原因/ Exemption Reason",
  },
  successHashPropertyPrefix: "PME_BACKUP_SUCCESS_HASH_",
  successPendingPropertyPrefix: "PME_BACKUP_SUCCESS_PENDING_",
};

var _pmeBackup_STYLE_METHODS = [
  ["getNumberFormats", "setNumberFormats"],
  ["getBackgrounds", "setBackgrounds"],
  ["getFontColors", "setFontColors"],
  ["getFontFamilies", "setFontFamilies"],
  ["getFontSizes", "setFontSizes"],
  ["getFontWeights", "setFontWeights"],
  ["getFontStyles", "setFontStyles"],
  ["getFontLines", "setFontLines"],
  ["getHorizontalAlignments", "setHorizontalAlignments"],
  ["getVerticalAlignments", "setVerticalAlignments"],
  ["getWraps", "setWraps"],
];

/**
 * PME年度归档统一入口。定时运行仅限每周一，手动运行不受日期限制。
 * @param {Object=} e 调度器事件对象
 * @returns {Object} 本次运行摘要
 */
function backupPmeData(e) {
  var triggerLabel = e ? "定时" : "手动";
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    _pmeBackup_log("跳过", "另一个 PME 年度归档任务正在运行", triggerLabel, "");
    return { status: "locked" };
  }

  try {
    var now = new Date();
    var localDate = Utilities.formatDate(now, _pmeBackup_CONFIG.timeZone, "yyyy-MM-dd");
    var localWeekday = Number(Utilities.formatDate(now, _pmeBackup_CONFIG.timeZone, "u"));
    var targetYear = Number(localDate.slice(0, 4)) - 1;

    if (e && localWeekday !== 1) {
      _pmeBackup_log("跳过", "定时归档仅在每周一运行；当前日期 " + localDate, triggerLabel, "");
      return { status: "outside-window", targetYear: targetYear };
    }

    var sourceSpreadsheet = SpreadsheetApp.openById(_pmeBackup_CONFIG.sourceSpreadsheetId);
    var sourceMasterSheet = sourceSpreadsheet.getSheetByName(_pmeBackup_CONFIG.sourceMasterSheetName);
    var sourceKpiSheet = sourceSpreadsheet.getSheetByName(_pmeBackup_CONFIG.sourceKpiSheetName);
    if (!sourceMasterSheet) throw new Error("未找到来源工作表: " + _pmeBackup_CONFIG.sourceMasterSheetName);
    if (!sourceKpiSheet) throw new Error("未找到来源工作表: " + _pmeBackup_CONFIG.sourceKpiSheetName);

    var sourceRows = _pmeBackup_readTable(sourceMasterSheet, "Master Data");
    var sourceHeaders = sourceRows[0];
    var headerIndex = _pmeBackup_headerIndex(sourceHeaders, "Master Data");
    var preflight = _pmeBackup_collectYearRows(sourceRows, headerIndex, targetYear);

    if (preflight.blockers.length > 0) {
      try {
        _pmeBackup_sendBlockerEmail(
          targetYear,
          preflight.blockers,
          sourceSpreadsheet.getUrl(),
          headerIndex,
        );
      } catch (mailError) {
        _pmeBackup_log(
          "异常",
          "待处理 PME 提醒邮件发送失败: " + mailError.message,
          triggerLabel,
          sourceSpreadsheet.getUrl(),
        );
      }
    }

    if (preflight.invalidYearMonth.length > 0 || preflight.blockers.length > 0) {
      var invalidDetail = preflight.invalidYearMonth.length > 0
        ? "；YearMonth 无效/空白 " + preflight.invalidYearMonth.length + " 行: " +
          _pmeBackup_summarizeRows(preflight.invalidYearMonth)
        : "";
      var blockerDetail = preflight.blockers.length > 0
        ? "；目标年未处理 " + preflight.blockers.length + " 行"
        : "";
      _pmeBackup_log(
        "警告",
        targetYear + " 年 PME/KPI 整年归档暂停" + blockerDetail + invalidDetail,
        triggerLabel,
        sourceSpreadsheet.getUrl(),
      );
      return {
        status: "blocked",
        targetYear: targetYear,
        blockerCount: preflight.blockers.length,
        invalidYearMonthCount: preflight.invalidYearMonth.length,
      };
    }

    var targetSpreadsheet = SpreadsheetApp.openById(_pmeBackup_CONFIG.destinationSpreadsheetId);
    var destinationPmeSheetName = String(targetYear);
    var destinationPmeSheet = targetSpreadsheet.getSheetByName(destinationPmeSheetName);
    var destinationKpiSheet = targetSpreadsheet.getSheetByName(_pmeBackup_CONFIG.destinationKpiSheetName);
    if (!destinationKpiSheet) {
      throw new Error("目标工作簿缺少标签页: " + _pmeBackup_CONFIG.destinationKpiSheetName);
    }

    var sourceSerialCounts = _pmeBackup_countSourceSerials(sourceRows, headerIndex);
    _pmeBackup_validateCandidateSerials(preflight.targetRows, sourceSerialCounts, headerIndex);

    var existingPmeRows = _pmeBackup_readArchiveRows(destinationPmeSheet, sourceHeaders);
    _pmeBackup_validateArchiveSerials(existingPmeRows.rows, headerIndex, targetYear);
    var mergedPmeRows = _pmeBackup_mergePmeRows(
      existingPmeRows.rows,
      preflight.targetRows,
      headerIndex,
    );

    var sourceKpiBlock = _pmeBackup_findKpiBlock(sourceKpiSheet, targetYear);
    var destinationKpiBlocks = _pmeBackup_findKpiBlocks(destinationKpiSheet, targetYear);
    if (destinationKpiBlocks.length > 1) {
      throw new Error("目标 KPI 标签页中存在多个 " + targetYear + " 年区块，无法唯一更新");
    }
    var destinationKpiBlock = destinationKpiBlocks.length === 1 ? destinationKpiBlocks[0] : null;
    var sourceSnapshotHash = _pmeBackup_hash(sourceRows);

    if (preflight.targetRows.length > 0) {
      if (!destinationPmeSheet) {
        destinationPmeSheet = targetSpreadsheet.insertSheet(destinationPmeSheetName);
      }
      _pmeBackup_writePmeArchive(
        destinationPmeSheet,
        sourceMasterSheet,
        sourceHeaders,
        mergedPmeRows,
        headerIndex,
      );
    }

    var writtenKpiBlock = _pmeBackup_writeKpiBlock(
      destinationKpiSheet,
      sourceKpiBlock,
      destinationKpiBlock,
    );
    SpreadsheetApp.flush();

    var verifiedPmeRows = preflight.targetRows.length > 0
      ? _pmeBackup_verifyPmeArchive(destinationPmeSheet, sourceHeaders, mergedPmeRows)
      : (destinationPmeSheet ? existingPmeRows.rows : []);
    var verifiedKpiBlock = _pmeBackup_verifyKpiBlock(
      destinationKpiSheet,
      targetYear,
      sourceKpiBlock,
    );

    var currentSourceRows = _pmeBackup_readTable(sourceMasterSheet, "Master Data");
    if (_pmeBackup_hash(currentSourceRows) !== sourceSnapshotHash) {
      throw new Error("归档期间 Master Data 发生变化；目标副本已写入但未删除任何源记录，请重跑校验");
    }

    _pmeBackup_deleteSourceRows(sourceMasterSheet, preflight.targetRows);
    SpreadsheetApp.flush();
    _pmeBackup_verifySourceRowsRemoved(sourceMasterSheet, targetYear, headerIndex);

    var archiveVersion = _pmeBackup_hash({
      year: targetYear,
      pmeRows: verifiedPmeRows,
      kpiRows: verifiedKpiBlock.values,
    });
    var notificationResult = _pmeBackup_sendSuccessIfChanged({
      year: targetYear,
      now: new Date(),
      movedCount: preflight.targetRows.length,
      totalPmeCount: verifiedPmeRows.length,
      destinationUrl: targetSpreadsheet.getUrl(),
      archiveVersion: archiveVersion,
      triggerLabel: triggerLabel,
    });
    var notificationWarning = notificationResult === "failed" || notificationResult === "sent-state-unrecorded";

    _pmeBackup_log(
      notificationWarning ? "警告" : "成功",
      targetYear + " 年 PME/KPI 整批归档完成；本次剪切 " + preflight.targetRows.length +
        " 行，年度 PME 共 " + verifiedPmeRows.length + " 行；KPI 区块已验证" +
        (notificationResult === "failed" ? "；成功邮件待重试" :
          (notificationResult === "sent-state-unrecorded" ? "；成功邮件状态未能保存" : "")),
      triggerLabel,
      targetSpreadsheet.getUrl(),
    );

    return {
      status: notificationResult === "failed" ? "archived-notification-pending" :
        (notificationResult === "sent-state-unrecorded" ? "archived-notification-state-warning" : "archived"),
      targetYear: targetYear,
      movedCount: preflight.targetRows.length,
      totalPmeCount: verifiedPmeRows.length,
      kpiRows: writtenKpiBlock.values.length,
      notification: notificationResult,
    };
  } catch (error) {
    _pmeBackup_log("异常", "PME 年度归档失败: " + error.message, triggerLabel, "");
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function _pmeBackup_readTable(sheet, label) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) throw new Error(label + " 为空或缺少表头");
  return sheet.getRange(1, 1, lastRow, lastColumn).getValues();
}

function _pmeBackup_headerIndex(headers, label) {
  var duplicates = {};
  var index = {};
  (headers || []).forEach(function (header, column) {
    var name = String(header == null ? "" : header).trim();
    if (!name) return;
    if (index[name] !== undefined) duplicates[name] = true;
    else index[name] = column;
  });
  var duplicateNames = Object.keys(duplicates);
  if (duplicateNames.length > 0) {
    throw new Error(label + " 存在重复表头: " + duplicateNames.join(", "));
  }
  var required = Object.keys(_pmeBackup_CONFIG.headers).map(function (key) {
    return _pmeBackup_CONFIG.headers[key];
  });
  var missing = required.filter(function (name) { return index[name] === undefined; });
  if (missing.length > 0) throw new Error(label + " 缺少必需字段: " + missing.join(", "));
  return index;
}

function _pmeBackup_collectYearRows(rows, headerIndex, targetYear) {
  var yearColumn = headerIndex[_pmeBackup_CONFIG.headers.yearMonth];
  var statusColumn = headerIndex[_pmeBackup_CONFIG.headers.productionStatus];
  var exemptColumn = headerIndex[_pmeBackup_CONFIG.headers.exempt];
  var reasonColumn = headerIndex[_pmeBackup_CONFIG.headers.exemptionReason];
  var result = { targetRows: [], blockers: [], invalidYearMonth: [] };

  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    var row = rows[rowIndex];
    if (_pmeBackup_isBlankRow(row)) continue;
    var yearMonth = _pmeBackup_parseYearMonth(row[yearColumn]);
    var serialColumn = headerIndex[_pmeBackup_CONFIG.headers.serial];
    var record = {
      rowNumber: rowIndex + 1,
      values: row,
      yearMonth: yearMonth.text,
      serial: String(row[serialColumn] == null ? "" : row[serialColumn]).trim(),
    };
    if (!yearMonth.valid) {
      result.invalidYearMonth.push(record);
      continue;
    }
    if (yearMonth.year !== targetYear) continue;

    result.targetRows.push(record);
    var status = String(row[statusColumn] == null ? "" : row[statusColumn]).trim();
    var reason = String(row[reasonColumn] == null ? "" : row[reasonColumn]).trim();
    var exempt = _pmeBackup_isChecked(row[exemptColumn]);
    var isResolved = status === "生产" || (status === "未达144" && exempt && reason !== "");
    if (!isResolved) {
      record.status = status;
      record.exempt = row[exemptColumn];
      record.reason = row[reasonColumn];
      result.blockers.push(record);
    }
  }
  return result;
}

function _pmeBackup_parseYearMonth(value) {
  var text = String(value == null ? "" : value).trim();
  var match = text.match(/^(\d{4})(0[1-9]|1[0-2])$/);
  if (!match) return { valid: false, text: text, year: null, month: null };
  return { valid: true, text: text, year: Number(match[1]), month: Number(match[2]) };
}

function _pmeBackup_isBlankRow(row) {
  return (row || []).every(function (value) {
    return value === "" || value === null;
  });
}

function _pmeBackup_isChecked(value) {
  if (value === true) return true;
  return String(value == null ? "" : value).trim().toLowerCase() === "true";
}

function _pmeBackup_countSourceSerials(rows, headerIndex) {
  var serialColumn = headerIndex[_pmeBackup_CONFIG.headers.serial];
  var counts = {};
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    var row = rows[rowIndex];
    if (_pmeBackup_isBlankRow(row)) continue;
    var serial = String(row[serialColumn] == null ? "" : row[serialColumn]).trim();
    if (!serial) continue;
    counts[serial] = (counts[serial] || 0) + 1;
  }
  return counts;
}

function _pmeBackup_validateCandidateSerials(records, sourceCounts, headerIndex) {
  var serialColumn = headerIndex[_pmeBackup_CONFIG.headers.serial];
  var seen = {};
  records.forEach(function (record) {
    var serial = String(record.values[serialColumn] == null ? "" : record.values[serialColumn]).trim();
    if (!serial) throw new Error("目标年度 PME 第 " + record.rowNumber + " 行缺少 Serial No.");
    if (seen[serial]) throw new Error("目标年度 PME 存在重复 Serial No.: " + serial);
    if (sourceCounts[serial] !== 1) throw new Error("Master Data 中 Serial No. 无法唯一匹配: " + serial);
    seen[serial] = true;
  });
}

function _pmeBackup_readArchiveRows(sheet, sourceHeaders) {
  if (!sheet || sheet.getLastRow() === 0) return { rows: [], headerPresent: false };
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastColumn > sourceHeaders.length) {
    throw new Error("目标 PME 年度标签页存在超出来源表头的列: " + sheet.getName());
  }
  var grid = sheet.getRange(1, 1, lastRow, sourceHeaders.length).getValues();
  var header = grid[0] || [];
  if (_pmeBackup_isBlankRow(header)) {
    if (grid.slice(1).some(function (row) { return !_pmeBackup_isBlankRow(row); })) {
      throw new Error("目标 PME 年度标签页表头为空但存在记录: " + sheet.getName());
    }
    return { rows: [], headerPresent: false };
  }
  if (!_pmeBackup_rowsEqual([header], [sourceHeaders])) {
    throw new Error("目标 PME 年度标签页表头与 Master Data 不一致: " + sheet.getName());
  }
  var rows = [];
  var blankSeen = false;
  for (var i = 1; i < grid.length; i++) {
    if (_pmeBackup_isBlankRow(grid[i])) {
      blankSeen = true;
      continue;
    }
    if (blankSeen) throw new Error("目标 PME 年度标签页存在内部空行: " + sheet.getName());
    rows.push(grid[i]);
  }
  return { rows: rows, headerPresent: true };
}

function _pmeBackup_validateArchiveSerials(rows, headerIndex, targetYear) {
  var serialColumn = headerIndex[_pmeBackup_CONFIG.headers.serial];
  var yearMonthColumn = headerIndex[_pmeBackup_CONFIG.headers.yearMonth];
  var seen = {};
  rows.forEach(function (row) {
    var serial = String(row[serialColumn] == null ? "" : row[serialColumn]).trim();
    if (!serial) throw new Error("目标 PME 年度标签页存在空白 Serial No.");
    if (seen[serial]) throw new Error("目标 PME 年度标签页存在重复 Serial No.: " + serial);
    var yearMonth = _pmeBackup_parseYearMonth(row[yearMonthColumn]);
    if (!yearMonth.valid || yearMonth.year !== targetYear) {
      throw new Error(
        "目标 PME " + targetYear + " 年标签页存在 YearMonth 无效或年份不符的记录: " + serial,
      );
    }
    seen[serial] = true;
  });
}

function _pmeBackup_mergePmeRows(existingRows, sourceRecords, headerIndex) {
  var serialColumn = headerIndex[_pmeBackup_CONFIG.headers.serial];
  var rows = (existingRows || []).map(function (row) { return row.slice(); });
  var indexBySerial = {};
  rows.forEach(function (row, index) {
    indexBySerial[String(row[serialColumn]).trim()] = index;
  });
  sourceRecords.forEach(function (record) {
    var serial = String(record.values[serialColumn]).trim();
    if (indexBySerial[serial] !== undefined) rows[indexBySerial[serial]] = record.values.slice();
    else {
      indexBySerial[serial] = rows.length;
      rows.push(record.values.slice());
    }
  });
  return rows;
}

function _pmeBackup_writePmeArchive(sheet, sourceSheet, headers, rows, headerIndex) {
  var values = [headers.slice()].concat(rows.map(function (row) { return row.slice(); }));
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  var targetHeaderRange = sheet.getRange(1, 1, 1, headers.length);
  var sourceHeaderRange = sourceSheet.getRange(1, 1, 1, headers.length);
  _pmeBackup_copyStyles(sourceHeaderRange, targetHeaderRange, false);
  if (rows.length > 0 && sourceSheet.getLastRow() >= 2) {
    var sourceTemplate = sourceSheet.getRange(2, 1, 1, headers.length);
    var targetBody = sheet.getRange(2, 1, rows.length, headers.length);
    _pmeBackup_copyStyles(sourceTemplate, targetBody, true);
    _pmeBackup_syncCheckboxValidation(sheet, rows, 2, headerIndex);
  }
}

function _pmeBackup_findKpiBlocks(sheet, targetYear) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];
  var grid = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var blocks = [];
  var yearText = String(targetYear);

  for (var rowIndex = 0; rowIndex < grid.length; rowIndex++) {
    var row = grid[rowIndex];
    for (var processColumn = 0; processColumn < row.length; processColumn++) {
      if (String(row[processColumn] == null ? "" : row[processColumn]).trim() !== "Process") continue;
      var monthColumns = {};
      var annualColumns = [];
      for (var column = processColumn + 1; column < row.length; column++) {
        var heading = String(row[column] == null ? "" : row[column]).trim().replace(/\s+/g, "");
        var monthMatch = heading.match(new RegExp("^" + yearText + "(0[1-9]|1[0-2])$"));
        if (monthMatch) monthColumns[Number(monthMatch[1])] = column;
        if (heading === yearText + "A" || heading === yearText + "YTD") annualColumns.push(column);
      }
      if (Object.keys(monthColumns).length !== 12 || annualColumns.length === 0) continue;

      var endColumn = Math.max.apply(null, Object.keys(monthColumns).map(function (month) {
        return monthColumns[month];
      }).concat(annualColumns));
      var rowCount = 1;
      for (var nextRow = rowIndex + 1; nextRow < grid.length; nextRow++) {
        var processName = String(grid[nextRow][processColumn] == null ? "" : grid[nextRow][processColumn]).trim();
        if (!processName || processName === "Process") break;
        rowCount++;
      }
      if (rowCount < 2) continue;

      blocks.push({
        startRow: rowIndex + 1,
        startColumn: processColumn + 1,
        rowCount: rowCount,
        width: endColumn - processColumn + 1,
      });
    }
  }
  return blocks;
}

function _pmeBackup_findKpiBlock(sheet, targetYear) {
  var blocks = _pmeBackup_findKpiBlocks(sheet, targetYear);
  if (blocks.length === 0) throw new Error("KPI Summary 中找不到 " + targetYear + " 年完整 KPI 区块");
  if (blocks.length > 1) throw new Error("KPI Summary 中存在多个 " + targetYear + " 年 KPI 区块");
  var block = blocks[0];
  var range = sheet.getRange(block.startRow, block.startColumn, block.rowCount, block.width);
  block.values = range.getValues();
  block.styles = _pmeBackup_readStyles(range);
  return block;
}

function _pmeBackup_writeKpiBlock(sheet, sourceBlock, existingBlock) {
  var startRow;
  var startColumn;
  if (existingBlock) {
    startRow = existingBlock.startRow;
    startColumn = existingBlock.startColumn;
    if (sourceBlock.rowCount > existingBlock.rowCount) {
      sheet.insertRowsAfter(
        existingBlock.startRow + existingBlock.rowCount - 1,
        sourceBlock.rowCount - existingBlock.rowCount,
      );
    }
    sheet.getRange(
      existingBlock.startRow,
      existingBlock.startColumn,
      existingBlock.rowCount,
      existingBlock.width,
    ).clear();
  } else {
    startRow = sheet.getLastRow() > 0 ? sheet.getLastRow() + 2 : 1;
    startColumn = sourceBlock.startColumn;
  }

  var destinationRange = sheet.getRange(startRow, startColumn, sourceBlock.rowCount, sourceBlock.width);
  destinationRange.setValues(sourceBlock.values);
  _pmeBackup_applyStyles(sourceBlock.styles, destinationRange, false);
  return {
    startRow: startRow,
    startColumn: startColumn,
    rowCount: sourceBlock.rowCount,
    width: sourceBlock.width,
    values: sourceBlock.values,
    styles: sourceBlock.styles,
  };
}

function _pmeBackup_verifyKpiBlock(sheet, targetYear, sourceBlock) {
  var blocks = _pmeBackup_findKpiBlocks(sheet, targetYear);
  if (blocks.length !== 1) throw new Error("目标 KPI 标签页的 " + targetYear + " 年区块未能唯一验证");
  var block = blocks[0];
  if (block.rowCount !== sourceBlock.rowCount || block.width !== sourceBlock.width) {
    throw new Error("目标 KPI 区块行列范围与来源不一致");
  }
  var range = sheet.getRange(block.startRow, block.startColumn, block.rowCount, block.width);
  var values = range.getValues();
  if (!_pmeBackup_rowsEqual(values, sourceBlock.values)) throw new Error("目标 KPI 区块值校验失败");
  var formulas = range.getFormulas();
  if (formulas.some(function (row) { return row.some(function (formula) { return formula !== ""; }); })) {
    throw new Error("目标 KPI 快照中出现公式，静态值归档校验失败");
  }
  var styles = _pmeBackup_readStyles(range);
  if (_pmeBackup_hash(styles) !== _pmeBackup_hash(sourceBlock.styles)) {
    throw new Error("目标 KPI 区块格式校验失败");
  }
  block.values = values;
  return block;
}

function _pmeBackup_verifyPmeArchive(sheet, headers, expectedRows) {
  if (!sheet) throw new Error("目标 PME 年度标签页不存在");
  var grid = sheet.getRange(1, 1, sheet.getLastRow(), headers.length).getValues();
  if (!_pmeBackup_rowsEqual([grid[0] || []], [headers])) throw new Error("目标 PME 表头校验失败");
  var actualRows = grid.slice(1);
  if (!_pmeBackup_rowsEqual(actualRows, expectedRows)) throw new Error("目标 PME 年度数据字段校验失败");
  return actualRows;
}

function _pmeBackup_deleteSourceRows(sheet, records) {
  if (!records || records.length === 0) return;
  var recordByRow = {};
  records.forEach(function (record) { recordByRow[record.rowNumber] = record; });
  var width = records[0].values.length;
  var rowNumbers = records.map(function (record) { return record.rowNumber; })
    .sort(function (left, right) { return right - left; });
  var groups = [];
  rowNumbers.forEach(function (rowNumber) {
    var last = groups[groups.length - 1];
    if (last && rowNumber === last.startRow - 1) {
      last.startRow = rowNumber;
      last.count++;
    } else {
      groups.push({ startRow: rowNumber, count: 1 });
    }
  });
  groups.forEach(function (group) {
    var currentRows = sheet.getRange(group.startRow, 1, group.count, width).getValues();
    for (var offset = 0; offset < group.count; offset++) {
      var expected = recordByRow[group.startRow + offset];
      if (!expected || !_pmeBackup_rowsEqual([currentRows[offset]], [expected.values])) {
        throw new Error("删除源 PME 行前身份/字段复核失败，未删除当前行块: R" + group.startRow);
      }
    }
    sheet.deleteRows(group.startRow, group.count);
  });
}

function _pmeBackup_verifySourceRowsRemoved(sheet, targetYear, headerIndex) {
  if (sheet.getLastRow() < 2) return;
  var rows = _pmeBackup_readTable(sheet, "Master Data");
  var check = _pmeBackup_collectYearRows(rows, headerIndex, targetYear);
  if (check.invalidYearMonth.length > 0) {
    throw new Error("剪切后 Master Data 出现空白/无效 YearMonth，无法确认年度完整性");
  }
  var remaining = check.targetRows;
  if (remaining.length > 0) {
    throw new Error("源表仍有 " + targetYear + " 年 PME 记录，剪切结果未完整完成");
  }
}

function _pmeBackup_syncCheckboxValidation(sheet, rows, startRow, headerIndex) {
  if (!rows || rows.length === 0) return;
  var column = headerIndex[_pmeBackup_CONFIG.headers.exempt] + 1;
  var validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var usesCheckbox = function (row) {
    var value = row[column - 1];
    if (value === true || value === false) return true;
    var text = String(value == null ? "" : value).trim().toLowerCase();
    return text === "true" || text === "false";
  };
  var groupStart = 0;
  var current = usesCheckbox(rows[0]);
  for (var index = 1; index <= rows.length; index++) {
    var next = index < rows.length ? usesCheckbox(rows[index]) : !current;
    if (index === rows.length || next !== current) {
      var range = sheet.getRange(startRow + groupStart, column, index - groupStart, 1);
      if (current) range.setDataValidation(validation);
      else range.clearDataValidations();
      groupStart = index;
      current = next;
    }
  }
}

function _pmeBackup_copyStyles(sourceRange, destinationRange, repeatSingleRow) {
  var styles = _pmeBackup_readStyles(sourceRange);
  _pmeBackup_applyStyles(styles, destinationRange, repeatSingleRow);
}

function _pmeBackup_readStyles(range) {
  var styles = {};
  _pmeBackup_STYLE_METHODS.forEach(function (methods) {
    var getter = methods[0];
    if (typeof range[getter] === "function") styles[getter] = range[getter]();
  });
  return styles;
}

function _pmeBackup_applyStyles(styles, destinationRange, repeatSingleRow) {
  _pmeBackup_STYLE_METHODS.forEach(function (methods) {
    var getter = methods[0];
    var setter = methods[1];
    if (!styles || !styles[getter] || typeof destinationRange[setter] !== "function") return;
    var matrix = styles[getter];
    if (repeatSingleRow && matrix.length === 1 && destinationRange.getNumRows() > 1) {
      var template = matrix[0];
      matrix = [];
      for (var row = 0; row < destinationRange.getNumRows(); row++) matrix.push(template.slice());
    }
    destinationRange[setter](matrix);
  });
}

function _pmeBackup_sendBlockerEmail(targetYear, blockers, sourceUrl, headerIndex) {
  if (!blockers || blockers.length === 0) return;
  var subject = "[PME年度备份提醒] " + targetYear + " 年有 " + blockers.length + " 条待处理记录";
  var details = blockers.map(function (record) {
    var values = record.values;
    return {
      rowNumber: record.rowNumber,
      serial: _pmeBackup_cellText(values, _pmeBackup_CONFIG.headers.serial, headerIndex),
      yearMonth: record.yearMonth,
      workcenter: _pmeBackup_cellText(values, _pmeBackup_CONFIG.headers.workcenter, headerIndex),
      pmDate: _pmeBackup_formatCell(_pmeBackup_cell(values, _pmeBackup_CONFIG.headers.pmDate, headerIndex)),
      status: record.status || _pmeBackup_cellText(values, _pmeBackup_CONFIG.headers.productionStatus, headerIndex),
      exempt: record.exempt === undefined
        ? _pmeBackup_cellText(values, _pmeBackup_CONFIG.headers.exempt, headerIndex)
        : _pmeBackup_formatCell(record.exempt),
      reason: record.reason === undefined
        ? _pmeBackup_cellText(values, _pmeBackup_CONFIG.headers.exemptionReason, headerIndex)
        : _pmeBackup_formatCell(record.reason),
    };
  });
  var bodyLines = [
    targetYear + " 年 PME 年度归档暂停：有待处理记录。",
    "请完成记录状态或豁免信息后，归档程序会在后续运行重新检查。",
    "",
  ];
  details.forEach(function (item) {
    bodyLines.push(
      "Master Data 第 " + item.rowNumber + " 行 | Serial No.: " + item.serial +
      " | YearMonth: " + item.yearMonth + " | Workcenter: " + item.workcenter +
      " | PM_Date: " + item.pmDate + " | Production Status: " + item.status +
      " | Exempt?: " + item.exempt + " | Exemption Reason: " + item.reason,
    );
  });
  bodyLines.push("", "来源表: " + sourceUrl);

  var html = "<p>" + targetYear + " 年 PME 年度归档暂停：有待处理记录。</p>" +
    "<p>请完成记录状态或豁免信息后，归档程序会在后续运行重新检查。</p>" +
    "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse'>" +
    "<tr><th>Master Data 行</th><th>Serial No.</th><th>YearMonth</th><th>Workcenter</th>" +
    "<th>PM_Date</th><th>Production Status</th><th>Exempt?</th><th>Exemption Reason</th></tr>" +
    details.map(function (item) {
      return "<tr><td>" + item.rowNumber + "</td><td>" + _pmeBackup_escapeHtml(item.serial) +
        "</td><td>" + _pmeBackup_escapeHtml(item.yearMonth) + "</td><td>" + _pmeBackup_escapeHtml(item.workcenter) +
        "</td><td>" + _pmeBackup_escapeHtml(item.pmDate) + "</td><td>" + _pmeBackup_escapeHtml(item.status) +
        "</td><td>" + _pmeBackup_escapeHtml(item.exempt) + "</td><td>" + _pmeBackup_escapeHtml(item.reason) + "</td></tr>";
    }).join("") + "</table><p><a href='" + _pmeBackup_escapeHtml(sourceUrl) + "'>打开 PME 来源表</a></p>";

  GmailApp.sendEmail(
    _pmeBackup_CONFIG.recipient,
    subject,
    bodyLines.join("\n"),
    { htmlBody: html, name: "PME年度数据备份" },
  );
}

function _pmeBackup_sendSuccessIfChanged(info) {
  var properties = PropertiesService.getScriptProperties();
  var hashKey = _pmeBackup_CONFIG.successHashPropertyPrefix + info.year;
  var pendingKey = _pmeBackup_CONFIG.successPendingPropertyPrefix + info.year;
  var lastHash = properties.getProperty(hashKey) || "";
  var pendingHash = properties.getProperty(pendingKey) || "";
  if (info.archiveVersion === lastHash) {
    if (pendingHash) {
      try {
        properties.deleteProperty(pendingKey);
      } catch (cleanupError) {
        _pmeBackup_log("异常", "清理 PME 成功通知状态失败: " + cleanupError.message, info.triggerLabel, info.destinationUrl);
      }
    }
    return "unchanged";
  }

  try {
    properties.setProperty(pendingKey, info.archiveVersion);
  } catch (propertyError) {
    _pmeBackup_log("异常", "无法记录 PME 成功通知待发送状态: " + propertyError.message, info.triggerLabel, info.destinationUrl);
  }

  var timestamp = Utilities.formatDate(info.now, _pmeBackup_CONFIG.timeZone, "yyyy-MM-dd HH:mm:ss");
  var subject = "[PME年度备份成功] " + info.year + " 年归档完成/更新";
  var body = [
    info.year + " 年 PME 年度数据归档已完成并通过目标校验。",
    "本次剪切 PME 记录: " + info.movedCount + " 行",
    "归档后年度 PME 记录: " + info.totalPmeCount + " 行",
    "KPI 区块: 已更新并验证",
    "完成时间: " + timestamp,
    "目标工作簿: " + info.destinationUrl,
  ].join("\n");
  var html = "<p><b>" + info.year + " 年 PME 年度数据归档已完成并通过目标校验。</b></p>" +
    "<ul><li>本次剪切 PME 记录: " + info.movedCount + " 行</li>" +
    "<li>归档后年度 PME 记录: " + info.totalPmeCount + " 行</li>" +
    "<li>KPI 区块: 已更新并验证</li><li>完成时间: " + timestamp + "</li></ul>" +
    "<p><a href='" + _pmeBackup_escapeHtml(info.destinationUrl) + "'>打开 PME 历史归档工作簿</a></p>";

  try {
    GmailApp.sendEmail(
      _pmeBackup_CONFIG.recipient,
      subject,
      body,
      { htmlBody: html, name: "PME年度数据备份" },
    );
  } catch (mailError) {
    _pmeBackup_log("异常", "PME 成功邮件发送失败，将在后续运行重试: " + mailError.message, info.triggerLabel, info.destinationUrl);
    return "failed";
  }

  try {
    properties.setProperty(hashKey, info.archiveVersion);
  } catch (stateError) {
    _pmeBackup_log("异常", "PME 成功邮件已发送，但无法记录通知版本: " + stateError.message, info.triggerLabel, info.destinationUrl);
    return "sent-state-unrecorded";
  }
  try {
    properties.deleteProperty(pendingKey);
  } catch (cleanupError) {
    _pmeBackup_log("异常", "PME 成功邮件已发送，但清理待通知标记失败: " + cleanupError.message, info.triggerLabel, info.destinationUrl);
  }
  return "sent";
}

function _pmeBackup_cell(row, header, headerIndex) {
  if (!headerIndex || headerIndex[header] === undefined) return "";
  return row[headerIndex[header]];
}

function _pmeBackup_cellText(row, header, headerIndex) {
  var value = _pmeBackup_cell(row, header, headerIndex);
  return _pmeBackup_formatCell(value);
}

function _pmeBackup_formatCell(value) {
  if (value instanceof Date) return Utilities.formatDate(value, _pmeBackup_CONFIG.timeZone, "yyyy-MM-dd");
  return String(value == null ? "" : value).trim();
}

function _pmeBackup_escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function _pmeBackup_summarizeRows(records) {
  return records.slice(0, 20).map(function (record) {
    var serial = record.serial || "";
    return "R" + record.rowNumber + "/" + String(serial == null ? "" : serial);
  }).join(", ") + (records.length > 20 ? " 等" : "");
}

function _pmeBackup_rowsEqual(left, right) {
  return _pmeBackup_hash(left) === _pmeBackup_hash(right);
}

function _pmeBackup_hash(value) {
  var normalized = _pmeBackup_normalize(value);
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(normalized),
    Utilities.Charset.UTF_8,
  );
  return digest.map(function (byte) {
    return ("0" + (byte & 0xff).toString(16)).slice(-2);
  }).join("");
}

function _pmeBackup_normalize(value) {
  if (value instanceof Date) {
    return { type: "date", value: Utilities.formatDate(value, _pmeBackup_CONFIG.timeZone, "yyyy-MM-dd'T'HH:mm:ss") };
  }
  if (Array.isArray(value)) return value.map(_pmeBackup_normalize);
  if (value && typeof value === "object") {
    var result = {};
    Object.keys(value).sort().forEach(function (key) {
      result[key] = _pmeBackup_normalize(value[key]);
    });
    return result;
  }
  return value;
}

function _pmeBackup_log(status, detail, triggerLabel, remark) {
  try {
    writeLog("backupPmeData", status, detail || "", triggerLabel || "", remark || "");
  } catch (error) {
    console.error("PME备份 Log 写入失败: " + error.message);
  }
}
