// V20260607.1 — 保养数据年度备份模块

// ========== 保养数据备份配置 ==========
const PM_SPREADSHEET_ID = "1Y7FclPNn_yHWzwZiRCzSy350fppgXZ3NYgwA1OXQgD4";
const PM_SOURCE_SHEET_NAME = "PM_Records";
const PM_BACKUP_FOLDER_ID = "1PQyR1bQGVN0a4n1ynGVmq5a3uA2Z2FqD";
const PM_DATE_COL_IDX = 7; // H列 EndDate (0-based)

/**
 * 备份上一年保养数据到 Drive 文件夹，并从源表删除
 * 采用「筛选保留行 → 全量回写」方案，避免逐行 deleteRow 的 N 次 API 调用
 * @param {Object} e - 触发器事件对象（定时触发时传入，手动调用时为 undefined）
 */
function backupPMRecords(e) {
  var triggerLabel = e ? "定时" : "手动";
  try {
    var now = new Date();
    var prevYear = now.getFullYear() - 1;

    // 1. 读取源表
    var sourceSS = SpreadsheetApp.openById(PM_SPREADSHEET_ID);
    var sourceSheet = sourceSS.getSheetByName(PM_SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      writeLog("backupPMRecords", "异常", "未找到源表: " + PM_SOURCE_SHEET_NAME, triggerLabel, "");
      return;
    }

    var allData = sourceSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      writeLog("backupPMRecords", "跳过", "源表无数据行", triggerLabel, "");
      return;
    }

    var header = allData[0];
    var colCount = header.length;

    // 2. 单次遍历：拆分为「上年数据(备份)」和「保留数据(回写)」
    var prevYearRows = [];
    var keepRows = [header];

    for (var i = 1; i < allData.length; i++) {
      var dateVal = allData[i][PM_DATE_COL_IDX];
      var d = dateVal ? _pm_parseDate(dateVal) : null;

      if (d && d.getFullYear() === prevYear) {
        prevYearRows.push(allData[i]);
      } else {
        keepRows.push(allData[i]);
      }
    }

    if (prevYearRows.length === 0) {
      writeLog("backupPMRecords", "跳过", "无 " + prevYear + " 年数据", triggerLabel, "");
      return;
    }

    // 3. 创建备份文件 + 写入上年数据
    var timestamp = Utilities.formatDate(now, currentTimeZone, "yyyy-MM-dd HH:mm:ss");
    var backupFileName = "EDS_DPM_" + prevYear + "_备份时间:" + timestamp;
    var destFolder = DriveApp.getFolderById(PM_BACKUP_FOLDER_ID);

    var backupSS = SpreadsheetApp.create(backupFileName);
    DriveApp.getFileById(backupSS.getId()).moveTo(destFolder);

    var backupSheet = backupSS.getSheetByName(PM_SOURCE_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = backupSS.insertSheet(PM_SOURCE_SHEET_NAME);
      var defaultSheet = backupSS.getSheetByName("Sheet1");
      if (defaultSheet && backupSS.getSheets().length > 1) {
        backupSS.deleteSheet(defaultSheet);
      }
    }

    var backupData = [header].concat(prevYearRows);
    _pm_writeChunked(backupSheet, backupData, colCount);

    // 4. 验证备份
    var backupRowCount = backupSheet.getLastRow() - 1;
    if (backupRowCount !== prevYearRows.length) {
      writeLog("backupPMRecords", "异常",
        "备份行数不一致: 预期 " + prevYearRows.length + " 行, 实际 " + backupRowCount + " 行",
        triggerLabel, "");
      return;
    }

    // 5. 源表批量替换：清空 → 回写保留行
    sourceSheet.clearContents();
    _pm_writeChunked(sourceSheet, keepRows, colCount);

    // 清理多余空白行
    var newLastRow = keepRows.length;
    var maxRows = sourceSheet.getMaxRows();
    if (maxRows - newLastRow > 500) {
      sourceSheet.deleteRows(newLastRow + 1, maxRows - newLastRow);
    }

    writeLog("backupPMRecords", "成功",
      backupFileName + " 已剪切 " + prevYearRows.length + " 行",
      triggerLabel, "");

  } catch (err) {
    writeLog("backupPMRecords", "异常", "备份失败: " + err.message, triggerLabel, "");
    console.error("backupPMRecords error:", err);
  }
}

/** 解析日期值（兼容 Date 对象、ISO 字符串、YYYY-M-D 字符串） */
function _pm_parseDate(val) {
  if (val instanceof Date) {
    var d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  var str = String(val).trim();
  if (!str) return null;
  var normalized = str.replace(/-/g, "/");
  var d = new Date(normalized);
  if (!isNaN(d.getTime())) return d;
  d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/** 分批写入 sheet */
function _pm_writeChunked(sheet, data, colCount) {
  var CHUNK = 50000;
  for (var r = 0; r < data.length; r += CHUNK) {
    var chunk = data.slice(r, r + CHUNK);
    sheet.getRange(r + 1, 1, chunk.length, colCount).setValues(chunk);
  }
}
