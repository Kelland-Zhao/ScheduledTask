// V20260607.2 — 交接班数据年度备份模块（批量读写重构）

// ========== 交接班数据备份配置 ==========
const SHIFT_SPREADSHEET_ID = "10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w";
const SHIFT_SOURCE_SHEET_NAME = "Shift_Records";
const SHIFT_BACKUP_FOLDER_ID = "1LyqKxmCrTgZZuXGU2fmNQIPk_GaBXgpm";
const SHIFT_DATE_COL_IDX = 11; // L列 提交日期 (0-based)

/**
 * 备份上一年交接班数据到 Drive 文件夹，并从源表删除
 * 采用「筛选保留行 → 全量回写」方案，避免逐行 deleteRow 的 N 次 API 调用
 * @param {Object} e - 触发器事件对象（定时触发时传入，手动调用时为 undefined）
 */
function backupShiftRecords(e) {
  var triggerLabel = e ? "定时" : "手动";
  try {
    var now = new Date();
    var prevYear = now.getFullYear() - 1;

    // 1. 读取源表
    var sourceSS = SpreadsheetApp.openById(SHIFT_SPREADSHEET_ID);
    var sourceSheet = sourceSS.getSheetByName(SHIFT_SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      writeLog("backupShiftRecords", "异常", "未找到源表: " + SHIFT_SOURCE_SHEET_NAME, triggerLabel, "");
      return;
    }

    var allData = sourceSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      writeLog("backupShiftRecords", "跳过", "源表无数据行", triggerLabel, "");
      return;
    }

    var header = allData[0];
    var colCount = header.length;

    // 2. 单次遍历：拆分为「上年数据(备份)」和「保留数据(回写)」
    var prevYearRows = [];
    var keepRows = [header]; // 保留行以表头开头

    for (var i = 1; i < allData.length; i++) {
      var dateVal = allData[i][SHIFT_DATE_COL_IDX];
      var d = dateVal ? _sb_parseDate(dateVal) : null;

      if (d && d.getFullYear() === prevYear) {
        prevYearRows.push(allData[i]);
      } else {
        keepRows.push(allData[i]);
      }
    }

    if (prevYearRows.length === 0) {
      writeLog("backupShiftRecords", "跳过", "无 " + prevYear + " 年数据", triggerLabel, "");
      return;
    }

    // 3. 创建备份文件 + 写入上年数据
    var timestamp = Utilities.formatDate(now, currentTimeZone, "yyyy-MM-dd HH:mm:ss");
    var backupFileName = "EDS_DS_" + prevYear + "_备份时间:" + timestamp;
    var destFolder = DriveApp.getFolderById(SHIFT_BACKUP_FOLDER_ID);

    var backupSS = SpreadsheetApp.create(backupFileName);
    DriveApp.getFileById(backupSS.getId()).moveTo(destFolder);

    var backupSheet = backupSS.getSheetByName(SHIFT_SOURCE_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = backupSS.insertSheet(SHIFT_SOURCE_SHEET_NAME);
      var defaultSheet = backupSS.getSheetByName("Sheet1");
      if (defaultSheet && backupSS.getSheets().length > 1) {
        backupSS.deleteSheet(defaultSheet);
      }
    }

    var backupData = [header].concat(prevYearRows);
    _sb_writeChunked(backupSheet, backupData, colCount);

    // 4. 验证备份
    var backupRowCount = backupSheet.getLastRow() - 1; // 减去表头
    if (backupRowCount !== prevYearRows.length) {
      writeLog("backupShiftRecords", "异常",
        "备份行数不一致: 预期 " + prevYearRows.length + " 行, 实际 " + backupRowCount + " 行",
        triggerLabel, "");
      return;
    }

    // 5. 源表批量替换：清空 → 回写保留行（仅 1 次 clear + 1 次 write，替代 N 次 deleteRow）
    sourceSheet.clearContents();
    _sb_writeChunked(sourceSheet, keepRows, colCount);

    // 清理多余空白行，避免 sheet 膨胀
    var newLastRow = keepRows.length;
    var maxRows = sourceSheet.getMaxRows();
    if (maxRows - newLastRow > 500) {
      sourceSheet.deleteRows(newLastRow + 1, maxRows - newLastRow);
    }

    writeLog("backupShiftRecords", "成功",
      backupFileName + " 已剪切 " + prevYearRows.length + " 行",
      triggerLabel, "");

  } catch (err) {
    writeLog("backupShiftRecords", "异常", "备份失败: " + err.message, triggerLabel, "");
    console.error("backupShiftRecords error:", err);
  }
}

/** 分批写入 sheet，避免单次 setValues 超限 */
function _sb_writeChunked(sheet, data, colCount) {
  var CHUNK = 50000;
  for (var r = 0; r < data.length; r += CHUNK) {
    var chunk = data.slice(r, r + CHUNK);
    sheet.getRange(r + 1, 1, chunk.length, colCount).setValues(chunk);
  }
}

/**
 * 解析日期值（兼容 Date 对象、ISO 字符串、YYYY-M-D 字符串）
 * @param {*} val - 日期值
 * @returns {Date|null}
 */
function _sb_parseDate(val) {
  if (val instanceof Date) {
    var d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  // 字符串格式：处理 "2025-1-2" 或 "2025/1/2" 或 ISO 格式
  var str = String(val).trim();
  if (!str) return null;
  // 尝试替换 "-" 为 "/" 以兼容 Safari/GAS 的 Date 解析
  var normalized = str.replace(/-/g, "/");
  var d = new Date(normalized);
  if (!isNaN(d.getTime())) return d;
  // 回退：原始字符串
  d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}
