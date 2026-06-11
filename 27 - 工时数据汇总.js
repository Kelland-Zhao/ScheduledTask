// 工时数据汇总（TB1+TB2 → MasterData，合并实际工时）
// 入口：aggregateTB1TB2ToMasterData（每日 09:20 定时 or 手动）
// 数据源：_pd_TARGET_SHEET_ID Sheets: TB1, TB2
//         _pd_ATTENDANCE_SHEET_ID Sheet: 跟班考勤SUM
// 目标表：_pd_TARGET_SHEET_ID Sheet: MasterData

const _pd_ATTENDANCE_SHEET_ID = "1ZYh71zxJnBj8v5FlEghAPHZyJ96vyD6hB7UVDj0ebo8";

// ========== 主入口 ==========
function aggregateTB1TB2ToMasterData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const startTime = new Date();
    const logDetails = [];

    const targetSpreadsheet = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    const tb1Sheet = targetSpreadsheet.getSheetByName("TB1");
    const tb2Sheet = targetSpreadsheet.getSheetByName("TB2");
    const masterDataSheet = targetSpreadsheet.getSheetByName("MasterData");

    const tb1Data = _pd_convertSheetToMasterData(tb1Sheet, "TB1");
    const tb2Data = _pd_convertSheetToMasterData(tb2Sheet, "TB2");
    logDetails.push(`TB1: ${tb1Data.length}条, TB2: ${tb2Data.length}条`);

    const allData = tb1Data.concat(tb2Data);
    if (allData.length === 0) throw new Error("没有有效数据需要同步");

    const actualManhourMap = _pd_getActualManhourData();
    logDetails.push(`实际工时记录: ${actualManhourMap.size}条`);

    const mergedData = _pd_mergeActualManhour(allData, actualManhourMap);
    const matchedCount = mergedData.filter(r => r[7] !== "").length;
    logDetails.push(`实际工时匹配: ${matchedCount}条`);

    _pd_incrementalUpdateMasterData(masterDataSheet, mergedData);
    logDetails.push(`写入MasterData: ${mergedData.length}条`);

    const duration = ((new Date()) - startTime) / 1000;
    writeLog("aggregateTB1TB2ToMasterData", "成功", logDetails.join("; "), trigger, `耗时${duration}s`);

  } catch (err) {
    try { writeLog("aggregateTB1TB2ToMasterData", "失败", err.message, trigger, ""); } catch (e2) {}
    throw err;
  }
}

// ========== 辅助函数 ==========
function _pd_convertSheetToMasterData(sheet, workshop) {
  const result = [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return result;

  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const dateShiftHeaders = allData[0];
  const operatingQtyRow = allData[1];

  for (let colIndex = 1; colIndex < dateShiftHeaders.length; colIndex++) {
    const dateShift = dateShiftHeaders[colIndex];
    const operatingQty = operatingQtyRow[colIndex];
    if (!dateShift || !operatingQty || Number(operatingQty) <= 0) continue;

    const monthStr = _pd_extractMonth(String(dateShift));
    const weekNum = _pd_calculateWeek(String(dateShift));

    for (let rowIndex = 2; rowIndex < allData.length; rowIndex++) {
      const personName = allData[rowIndex][0];
      const manhour = allData[rowIndex][colIndex];
      if (!personName || !manhour || Number(manhour) <= 0) continue;
      result.push([dateShift, workshop, operatingQty, personName, manhour, monthStr, weekNum, ""]);
    }
  }
  return result;
}

function _pd_extractMonth(dateShiftStr) {
  if (!dateShiftStr) return "";
  const segs = String(dateShiftStr).split("_")[0].split(".");
  return segs.length >= 2 ? `${segs[0]}.${segs[1]}` : "";
}

function _pd_calculateWeek(dateShiftStr) {
  if (!dateShiftStr) return "";
  try {
    const segs = String(dateShiftStr).split("_")[0].split(".");
    if (segs.length >= 3) {
      const date = new Date(parseInt(segs[0]), parseInt(segs[1]) - 1, parseInt(segs[2]));
      const firstDay = new Date(date.getFullYear(), 0, 1);
      return Math.ceil((Math.floor((date - firstDay) / 86400000) + firstDay.getDay() + 1) / 7);
    }
  } catch (e) {}
  return "";
}

function _pd_convertShiftType(shift1) {
  switch (String(shift1).trim()) {
    case "1": return "1夜";
    case "2": return "2早";
    case "3": case "4": return "3中";
    default: return null;
  }
}

function _pd_getCurrentMonth() {
  return Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy.MM");
}

function _pd_buildAttendanceDateShift(timeStr, shift1) {
  if (!timeStr || !shift1) return null;
  const shiftMapping = _pd_convertShiftType(shift1);
  if (!shiftMapping) return null;
  const clean = String(timeStr).replace("日期-", "");
  if (!clean || isNaN(clean)) return null;
  return `${_pd_getCurrentMonth()}.${clean.padStart(2, '0')}_${shiftMapping}`;
}

function _pd_getActualManhourData() {
  try {
    const sheet = SpreadsheetApp.openById(_pd_ATTENDANCE_SHEET_ID).getSheetByName("跟班考勤SUM");
    if (!sheet) { Logger.log("未找到'跟班考勤SUM'工作表"); return new Map(); }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return new Map();
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const map = new Map();
    data.forEach(row => {
      const name = row[2], timeStr = row[6], actualManhour = row[7], shift1 = row[8];
      if (!name || !timeStr || !shift1) return;
      const dateShift = _pd_buildAttendanceDateShift(timeStr, shift1);
      if (!dateShift) return;
      map.set(`${name}_${dateShift}`, actualManhour || "");
    });
    return map;
  } catch (err) {
    Logger.log(`读取实际工时数据失败: ${err.message}`);
    return new Map();
  }
}

function _pd_mergeActualManhour(masterData, actualManhourMap) {
  return masterData.map(record => [
    record[0], record[1], record[2], record[3], record[4], record[5], record[6],
    actualManhourMap.get(`${record[3]}_${record[0]}`) || ""
  ]);
}

function _pd_incrementalUpdateMasterData(sheet, data) {
  if (data.length === 0) return;
  const lastRow = sheet.getLastRow();
  const existingData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  const existingMap = new Map();
  existingData.forEach((row, i) => existingMap.set(`${row[0]}_${row[3]}`, { rowIndex: i + 2, data: row }));

  const updateOps = [], newRecords = [];
  data.forEach(newRow => {
    const key = `${newRow[0]}_${newRow[3]}`;
    if (existingMap.has(key)) {
      updateOps.push({ row: existingMap.get(key).rowIndex, data: newRow });
      existingMap.delete(key);
    } else {
      newRecords.push(newRow);
    }
  });

  updateOps.forEach(op => sheet.getRange(op.row, 1, 1, 8).setValues([op.data]));

  if (existingMap.size > 0) {
    Array.from(existingMap.values())
      .sort((a, b) => b.rowIndex - a.rowIndex)
      .forEach(item => sheet.deleteRow(item.rowIndex));
  }

  if (newRecords.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRecords.length, 8).setValues(newRecords);
  }

  sheet.getRange(1, 1, 1, 8).setValues([[
    "日期班次 / Date & Shift", "车间 / Workshop", "开机数 / Operating Machine Qty",
    "姓名 / Name", "安排工时 / Scheduling Manhour", "月份 / Month", "周 / Week",
    "实际工时 / Actual Manhour"
  ]]);
}
