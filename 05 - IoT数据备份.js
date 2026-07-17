// V20260525.1 — IoT 数据备份模块

// ========== IoT 数据备份配置 ==========
const IOT_SPREADSHEET_ID = "1nZkZv1voGKyCzAJTxBZn8kKGVhXvGqlbo8ygJrq-cZA";
const IOT_SOURCE_SHEET_NAME = "IoT_CT_Detail";
const IOT_DEST_FOLDER_ID = "193URIPvHupVuxv0aN1kYC9ePYVkArKKs";
const OPERA_DEST_FOLDER_ID = "1zrhKMMyEI_9HCKhpCVEygOKrcrvzVXxz";

/** 将 IoT_CT_Detail 数据复制到目标文件夹，每日一个 Spreadsheet（幂等）*/
function copyIoTDetailToFolder(e) {
  const today = Utilities.formatDate(new Date(), currentTimeZone, "yyyy-MM-dd");
  const fileName = IOT_SOURCE_SHEET_NAME + "_" + today;
  const destFolder = DriveApp.getFolderById(IOT_DEST_FOLDER_ID);

  let destSS;
  const existing = destFolder.getFilesByName(fileName);
  if (existing.hasNext()) {
    destSS = SpreadsheetApp.open(existing.next());
  } else {
    destSS = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(destSS.getId()).moveTo(destFolder);
  }

  const sourceSheet = SpreadsheetApp.openById(IOT_SPREADSHEET_ID).getSheetByName(IOT_SOURCE_SHEET_NAME);
  const sourceData = sourceSheet.getDataRange().getValues();

  // A列日期格式化为 yyyy-MM-dd HH:mm:ss
  const enrichedData = sourceData.map(function(row, idx) {
    if (idx === 0) return row.concat(["HourSlot"]);
    const rawDate = row[0];
    const d = rawDate instanceof Date ? rawDate : (rawDate ? new Date(rawDate) : null);
    let formattedDate = "";
    let hourSlot = "";
    if (d && !isNaN(d.getTime())) {
      const y = d.getFullYear();
      const mo = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      const h = String(d.getHours()).padStart(2, "0");
      const mi = String(d.getMinutes()).padStart(2, "0");
      const s = String(d.getSeconds()).padStart(2, "0");
      formattedDate = y + "-" + mo + "-" + day + " " + h + ":" + mi + ":" + s;
      hourSlot = y + "-" + mo + "-" + day + "_" + h;
    }
    const newRow = row.slice(); // 浅拷贝，避免修改原数组
    newRow[0] = formattedDate;
    newRow.push(hourSlot);
    return newRow;
  });

  let destSheet = destSS.getSheetByName(IOT_SOURCE_SHEET_NAME);
  if (!destSheet) {
    destSheet = destSS.insertSheet(IOT_SOURCE_SHEET_NAME);
    const defaultSheet = destSS.getSheetByName("Sheet1");
    if (defaultSheet && destSS.getSheets().length > 1) destSS.deleteSheet(defaultSheet);
  }
  destSheet.clearContents();
  const CHUNK = 50000;
  for (let i = 0; i < enrichedData.length; i += CHUNK) {
    const chunk = enrichedData.slice(i, i + CHUNK);
    destSheet.getRange(i + 1, 1, chunk.length, enrichedData[0].length).setValues(chunk);
  }

  try { writeLog("copyIoTDetailToFolder", "成功", fileName + " 已写入 " + sourceData.length + " 行", e ? "定时" : "手动", ""); } catch (err) {}
}

/** 将 Opera_CT_Detail 数据复制到 Full_OPERA_Data 文件夹，每日一个 Spreadsheet（幂等）*/
function copyOperaDetailToFolder(e) {
  const OPERA_SOURCE_SHEET_NAME = "Opera_CT_Detail";
  const today = Utilities.formatDate(new Date(), currentTimeZone, "yyyy-MM-dd");
  const fileName = OPERA_SOURCE_SHEET_NAME + "_" + today;
  const destFolder = DriveApp.getFolderById(OPERA_DEST_FOLDER_ID);

  let destSS;
  const existing = destFolder.getFilesByName(fileName);
  if (existing.hasNext()) {
    destSS = SpreadsheetApp.open(existing.next());
  } else {
    destSS = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(destSS.getId()).moveTo(destFolder);
  }

  const sourceSheet = SpreadsheetApp.openById(IOT_SPREADSHEET_ID).getSheetByName(OPERA_SOURCE_SHEET_NAME);
  const sourceData = sourceSheet.getDataRange().getValues();

  // 表头映射: Line→机台号, TagDate→日期/时间, TagValue→周期
  // TagDate 格式化: yyyy-MM-dd HH:mm:ss; TagValue 四舍五入保留2位小数
  const enrichedData = sourceData.map(function(row, idx) {
    if (idx === 0) return ["机台号", "日期/时间", "周期", "HourSlot"];

    // 日期/时间列 (原 TagDate, index 1)
    const tagDateRaw = row[1];
    var formattedDate = "";
    var hourSlot = "";
    if (tagDateRaw) {
      var d = tagDateRaw instanceof Date ? tagDateRaw : new Date(String(tagDateRaw).replace(" ", "T"));
      if (d && !isNaN(d.getTime())) {
        var y = d.getFullYear();
        var mo = String(d.getMonth() + 1).padStart(2, "0");
        var day = String(d.getDate()).padStart(2, "0");
        var h = String(d.getHours()).padStart(2, "0");
        var mi = String(d.getMinutes()).padStart(2, "0");
        var s = String(d.getSeconds()).padStart(2, "0");
        formattedDate = y + "-" + mo + "-" + day + " " + h + ":" + mi + ":" + s;
        hourSlot = y + "-" + mo + "-" + day + "_" + h;
      }
    }

    // TagValue (index 2): 四舍五入保留2位小数
    var tagValueRaw = row[2];
    var roundedValue = "";
    if (tagValueRaw !== null && tagValueRaw !== undefined && tagValueRaw !== "") {
      var num = parseFloat(tagValueRaw);
      if (!isNaN(num)) {
        roundedValue = Math.round(num * 100) / 100;
      } else {
        roundedValue = tagValueRaw;
      }
    }

    return [row[0], formattedDate, roundedValue, hourSlot];
  });

  let destSheet = destSS.getSheetByName(OPERA_SOURCE_SHEET_NAME);
  if (!destSheet) {
    destSheet = destSS.insertSheet(OPERA_SOURCE_SHEET_NAME);
    const defaultSheet = destSS.getSheetByName("Sheet1");
    if (defaultSheet && destSS.getSheets().length > 1) destSS.deleteSheet(defaultSheet);
  }
  destSheet.clearContents();
  const CHUNK = 50000;
  for (let i = 0; i < enrichedData.length; i += CHUNK) {
    const chunk = enrichedData.slice(i, i + CHUNK);
    destSheet.getRange(i + 1, 1, chunk.length, enrichedData[0].length).setValues(chunk);
  }

  try { writeLog("copyOperaDetailToFolder", "成功", fileName + " 已写入 " + sourceData.length + " 行", e ? "定时" : "手动", ""); } catch (err) {}
}
