// V20260607.1 — 交接班数据年度备份模块

// ========== 交接班数据备份配置 ==========
const SHIFT_SPREADSHEET_ID = "10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w";
const SHIFT_SOURCE_SHEET_NAME = "Shift_Records";
const SHIFT_BACKUP_FOLDER_ID = "1LyqKxmCrTgZZuXGU2fmNQIPk_GaBXgpm";
const SHIFT_DATE_COL_IDX = 11; // L列 提交日期 (0-based)

/**
 * 备份上一年交接班数据到 Drive 文件夹，并从源表删除
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

    // 2. 筛选上一年数据行
    var prevYearRows = [];
    var prevYearRowNums = []; // 1-based sheet row numbers（用于删除）

    for (var i = 1; i < allData.length; i++) {
      var dateVal = allData[i][SHIFT_DATE_COL_IDX];
      if (!dateVal) continue;

      var d = _sb_parseDate(dateVal);
      if (!d) continue;

      if (d.getFullYear() === prevYear) {
        prevYearRows.push(allData[i]);
        prevYearRowNums.push(i + 1); // sheet 行号 = 数组索引 + 1
      }
    }

    if (prevYearRows.length === 0) {
      writeLog("backupShiftRecords", "跳过", "无 " + prevYear + " 年数据", triggerLabel, "");
      return;
    }

    // 3. 创建备份文件
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

    // 4. 分批写入备份（表头 + 数据行）
    var backupData = [header].concat(prevYearRows);
    var CHUNK = 50000;
    for (var r = 0; r < backupData.length; r += CHUNK) {
      var chunk = backupData.slice(r, r + CHUNK);
      backupSheet.getRange(r + 1, 1, chunk.length, colCount).setValues(chunk);
    }

    // 5. 验证备份
    var backupRowCount = backupSheet.getLastRow() - 1; // 减去表头
    if (backupRowCount !== prevYearRows.length) {
      writeLog("backupShiftRecords", "异常",
        "备份行数不一致: 预期 " + prevYearRows.length + " 行, 实际 " + backupRowCount + " 行",
        triggerLabel, "");
      return;
    }

    // 6. 从源表删除（从下往上删，避免行号偏移）
    for (var k = prevYearRowNums.length - 1; k >= 0; k--) {
      sourceSheet.deleteRow(prevYearRowNums[k]);
    }

    writeLog("backupShiftRecords", "成功",
      backupFileName + " 已剪切 " + prevYearRows.length + " 行",
      triggerLabel, "");

    // 同步更新源表行数，避免空白行残留
    var remainingRows = sourceSheet.getMaxRows() - sourceSheet.getLastRow();
    if (remainingRows > 500) {
      sourceSheet.deleteRows(sourceSheet.getLastRow() + 1, remainingRows);
    }

  } catch (err) {
    writeLog("backupShiftRecords", "异常", "备份失败: " + err.message, triggerLabel, "");
    console.error("backupShiftRecords error:", err);
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
