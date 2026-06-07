// V20260607.1 — 点检数据年度备份模块

// ========== 点检数据备份配置 ==========
const PC_SPREADSHEET_ID = "1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY";
const PC_SOURCE_SHEET_NAME = "InspectionRecords";
const PC_BACKUP_FOLDER_ID = "1mX6i2uAdK3VbuoM88TtMhHEup9yW0iAc";
const PC_DATE_COL_IDX = 7; // H列 Submit Date (0-based)

/**
 * 备份上一年点检数据到 Drive 文件夹，并从源表删除
 * 采用「筛选保留行 → 全量回写」方案，避免逐行 deleteRow 的 N 次 API 调用
 * @param {Object} e - 触发器事件对象（定时触发时传入，手动调用时为 undefined）
 */
function backupInspectionRecords(e) {
  var triggerLabel = e ? "定时" : "手动";
  try {
    var now = new Date();
    var prevYear = now.getFullYear() - 1;

    // 1. 读取源表
    var sourceSS = SpreadsheetApp.openById(PC_SPREADSHEET_ID);
    var sourceSheet = sourceSS.getSheetByName(PC_SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      writeLog("backupInspectionRecords", "异常", "未找到源表: " + PC_SOURCE_SHEET_NAME, triggerLabel, "");
      return;
    }

    var allData = sourceSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      writeLog("backupInspectionRecords", "跳过", "源表无数据行", triggerLabel, "");
      return;
    }

    var header = allData[0];
    var colCount = header.length;

    // 2. 单次遍历：拆分为「上年数据(备份)」和「保留数据(回写)」
    var prevYearRows = [];
    var keepRows = [header];

    for (var i = 1; i < allData.length; i++) {
      var dateVal = allData[i][PC_DATE_COL_IDX];
      var d = dateVal ? _pc_parseDate(dateVal) : null;

      if (d && d.getFullYear() === prevYear) {
        prevYearRows.push(allData[i]);
      } else {
        keepRows.push(allData[i]);
      }
    }

    if (prevYearRows.length === 0) {
      writeLog("backupInspectionRecords", "跳过", "无 " + prevYear + " 年数据", triggerLabel, "");
      return;
    }

    // 3. 创建备份文件 + 写入上年数据
    var timestamp = Utilities.formatDate(now, currentTimeZone, "yyyy-MM-dd HH:mm:ss");
    var backupFileName = "Database_PointCheck" + prevYear + "_备份时间:" + timestamp;
    var destFolder = DriveApp.getFolderById(PC_BACKUP_FOLDER_ID);

    var backupSS = SpreadsheetApp.create(backupFileName);
    DriveApp.getFileById(backupSS.getId()).moveTo(destFolder);

    var backupSheet = backupSS.getSheetByName(PC_SOURCE_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = backupSS.insertSheet(PC_SOURCE_SHEET_NAME);
      var defaultSheet = backupSS.getSheetByName("Sheet1");
      if (defaultSheet && backupSS.getSheets().length > 1) {
        backupSS.deleteSheet(defaultSheet);
      }
    }

    var backupData = [header].concat(prevYearRows);
    _pc_writeChunked(backupSheet, backupData, colCount);

    // 4. 验证备份
    var backupRowCount = backupSheet.getLastRow() - 1;
    if (backupRowCount !== prevYearRows.length) {
      writeLog("backupInspectionRecords", "异常",
        "备份行数不一致: 预期 " + prevYearRows.length + " 行, 实际 " + backupRowCount + " 行",
        triggerLabel, "");
      return;
    }

    // 5. 源表批量替换：清空 → 回写保留行
    sourceSheet.clearContents();
    _pc_writeChunked(sourceSheet, keepRows, colCount);

    // 清理多余空白行
    var newLastRow = keepRows.length;
    var maxRows = sourceSheet.getMaxRows();
    if (maxRows - newLastRow > 500) {
      sourceSheet.deleteRows(newLastRow + 1, maxRows - newLastRow);
    }

    writeLog("backupInspectionRecords", "成功",
      backupFileName + " 已剪切 " + prevYearRows.length + " 行",
      triggerLabel, "");

  } catch (err) {
    writeLog("backupInspectionRecords", "异常", "备份失败: " + err.message, triggerLabel, "");
    console.error("backupInspectionRecords error:", err);
  }
}

/**
 * 解析日期值
 * 兼容 Date 对象 / ISO 字符串 / YYYY-M-D / YYYY/MM/DD-HH:MM 等格式
 * @param {*} val - 日期值
 * @returns {Date|null}
 */
function _pc_parseDate(val) {
  if (val instanceof Date) {
    var d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  var str = String(val).trim();
  if (!str) return null;

  // 剥离时间部分（-HH:MM 或 空格HH:MM 或 T...），保留纯日期
  var dateOnly = str.replace(/[-\s]\d{1,2}:\d{2}.*$/, "").replace(/T.*$/, "");

  // 尝试直接解析日期部分
  var d = new Date(dateOnly);
  if (!isNaN(d.getTime())) return d;

  // 备选：替换 - 为 / 兼容不同引擎
  var normalized = dateOnly.replace(/-/g, "/");
  d = new Date(normalized);
  if (!isNaN(d.getTime())) return d;

  // 最后尝试原始字符串
  d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/** 分批写入 sheet */
function _pc_writeChunked(sheet, data, colCount) {
  var CHUNK = 50000;
  for (var r = 0; r < data.length; r += CHUNK) {
    var chunk = data.slice(r, r + CHUNK);
    sheet.getRange(r + 1, 1, chunk.length, colCount).setValues(chunk);
  }
}
