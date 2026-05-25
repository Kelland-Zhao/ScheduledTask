// V20260525.1 — IoT 数据备份模块

// ========== IoT 数据备份配置 ==========
const IOT_SPREADSHEET_ID = "1nZkZv1voGKyCzAJTxBZn8kKGVhXvGqlbo8ygJrq-cZA";
const IOT_SOURCE_SHEET_NAME = "IoT_CT_Detail";
const IOT_DEST_FOLDER_ID = "193URIPvHupVuxv0aN1kYC9ePYVkArKKs";

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

  let destSheet = destSS.getSheetByName(IOT_SOURCE_SHEET_NAME);
  if (!destSheet) {
    destSheet = destSS.insertSheet(IOT_SOURCE_SHEET_NAME);
    const defaultSheet = destSS.getSheetByName("Sheet1");
    if (defaultSheet && destSS.getSheets().length > 1) destSS.deleteSheet(defaultSheet);
  }
  destSheet.clearContents();
  if (sourceData.length > 0) {
    destSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
  }

  try { writeLog("copyIoTDetailToFolder", "成功", fileName + " 已写入 " + sourceData.length + " 行", e ? "定时" : "手动", ""); } catch (err) {}
}
