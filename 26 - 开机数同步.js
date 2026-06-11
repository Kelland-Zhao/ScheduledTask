// 开机数同步（从 Line日计划 → TB1/TB2 开机数更新）
// 入口：syncProductionData（每日 09:15 定时 or 手动）
// 数据源：_pd_SOURCE_SHEET_ID Sheet: 2.6_Line日计划
// 目标表：_pd_TARGET_SHEET_ID Sheets: TB1, TB2

const _pd_SOURCE_SHEET_ID = "1zNOM35TeOmaZTtaAR4izQdA37cLGSgPhJAw40KkYioQ";
const _pd_TARGET_SHEET_ID = "1dyS5C7r4pqYIeRT0p1zYzngt0EDCYR4hsswurAsEBYg";

// ========== 主入口 ==========
function syncProductionData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const startTime = new Date();
    const logDetails = [];

    const targetSpreadsheet = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    const tb1Sheet = targetSpreadsheet.getSheetByName("TB1");
    const tb2Sheet = targetSpreadsheet.getSheetByName("TB2");

    const targetDates = _pd_calculateTargetDates();
    const linePlanData = _pd_readLinePlanData(_pd_SOURCE_SHEET_ID);
    logDetails.push(`读取Line日计划: ${linePlanData.length}条`);

    const machineCountData = _pd_calculateMachineCount(linePlanData, targetDates);
    logDetails.push(`计算开机数: TB1(${Object.keys(machineCountData.TB1).length}), TB2(${Object.keys(machineCountData.TB2).length})`);

    const tb1DateDetails = [];
    const tb2DateDetails = [];
    targetDates.forEach(d => {
      tb1DateDetails.push(`${d.dateShift}(${machineCountData.TB1[d.dateShift] || 0}台)`);
      tb2DateDetails.push(`${d.dateShift}(${machineCountData.TB2[d.dateShift] || 0}台)`);
    });
    logDetails.push(`TB1: ${tb1DateDetails.join(', ')}`);
    logDetails.push(`TB2: ${tb2DateDetails.join(', ')}`);

    const tb1ExistingData = _pd_getExistingDataWithValues(tb1Sheet);
    const tb2ExistingData = _pd_getExistingDataWithValues(tb2Sheet);

    const tb1Results = _pd_updateMachineCount(tb1Sheet, machineCountData.TB1, tb1ExistingData);
    const tb2Results = _pd_updateMachineCount(tb2Sheet, machineCountData.TB2, tb2ExistingData);

    const duration = ((new Date()) - startTime) / 1000;
    logDetails.push(`TB1: 新增${tb1Results.created} 更新${tb1Results.updated} 跳过${tb1Results.skipped}`);
    logDetails.push(`TB2: 新增${tb2Results.created} 更新${tb2Results.updated} 跳过${tb2Results.skipped}`);

    writeLog("syncProductionData", "成功", logDetails.join("; "), trigger, `耗时${duration}s`);

  } catch (err) {
    try { writeLog("syncProductionData", "失败", err.message, trigger, ""); } catch (e2) {}
    throw err;
  }
}

// ========== 辅助函数 ==========
function _pd_calculateTargetDates() {
  const today = new Date();
  const dates = [];
  for (let i = 1; i <= 2; i++) {
    const targetDate = new Date(today.getTime() + i * 24 * 60 * 60 * 1000);
    const dateStr = Utilities.formatDate(targetDate, "Asia/Shanghai", "yyyy-MM-dd");
    dates.push(
      { date: dateStr, shift: '夜', dateShift: `${dateStr.replace(/-/g, '.')}_1夜` },
      { date: dateStr, shift: '早', dateShift: `${dateStr.replace(/-/g, '.')}_2早` },
      { date: dateStr, shift: '中', dateShift: `${dateStr.replace(/-/g, '.')}_3中` }
    );
  }
  return dates;
}

function _pd_readLinePlanData(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("2.6_Line日计划");
  if (!sheet) throw new Error("未找到'2.6_Line日计划'工作表");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 21).getValues();
}

function _pd_buildDateShift(dateStr, shiftStr) {
  const shiftMapping = { '早': '2早', '中': '3中', '夜': '1夜' };
  if (!dateStr || !shiftStr) return null;
  const date = dateStr instanceof Date ? dateStr : new Date(dateStr);
  if (isNaN(date.getTime())) return null;
  try {
    const dateFormatted = Utilities.formatDate(date, "Asia/Shanghai", "yyyy.MM.dd");
    return `${dateFormatted}_${shiftMapping[shiftStr] || shiftStr}`;
  } catch (e) {
    return null;
  }
}

function _pd_calculateMachineCount(data, targetDates) {
  const targetDateSet = new Set(targetDates.map(d => d.date));
  const machineCountMap = new Map();

  data.forEach(row => {
    const workshop = String(row[0]).trim();
    const dateValue = row[1];
    const aem = String(row[4]).trim();
    const machineStatus = String(row[6] || '');
    const shift = String(row[9]).trim();

    let dateStr = '';
    if (dateValue instanceof Date) {
      dateStr = Utilities.formatDate(dateValue, "Asia/Shanghai", "yyyy-MM-dd");
    } else {
      const raw = String(dateValue).trim();
      const sep = raw.includes('-') ? '-' : raw.includes('/') ? '/' : null;
      if (sep) {
        const parts = raw.split(sep);
        if (parts.length === 3) dateStr = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
      }
    }

    if (!targetDateSet.has(dateStr)) return;
    if (!workshop || !aem || !shift || !['TB1', 'TB2'].includes(workshop)) return;
    if (!machineStatus.includes('生产') && !machineStatus.includes('PQ')) return;
    if (!['早', '中', '夜'].includes(shift)) return;

    const dateShift = _pd_buildDateShift(dateValue, shift);
    if (!dateShift) return;

    const key = `${workshop}|${dateShift}`;
    if (!machineCountMap.has(key)) machineCountMap.set(key, new Set());
    machineCountMap.get(key).add(aem);
  });

  const result = { TB1: {}, TB2: {} };
  machineCountMap.forEach((aemSet, key) => {
    const [workshop, dateShift] = key.split('|');
    if (workshop && dateShift) result[workshop][dateShift] = aemSet.size;
  });
  return result;
}

function _pd_getExistingDataWithValues(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return new Map();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const existingData = new Map();
  headers.forEach((header, index) => {
    const h = String(header).trim();
    if (h && h !== "日期班次 / Date & Shift" && h !== "开机数 / Operating Machine Qty") {
      existingData.set(h, values[index]);
    }
  });
  return existingData;
}

function _pd_updateMachineCount(sheet, newData, existingData) {
  const results = { created: 0, updated: 0, skipped: 0, total: 0 };
  const lastCol = sheet.getLastColumn();
  let headers = lastCol >= 1 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  Object.entries(newData).forEach(([dateShift, newCount]) => {
    results.total++;
    const existingCount = existingData.get(dateShift);
    if (existingCount === undefined || existingCount === null || existingCount === '') {
      sheet.getRange(1, headers.length + 1).setValue(dateShift);
      sheet.getRange(2, headers.length + 1).setValue(newCount);
      headers.push(dateShift);
      results.created++;
    } else if (Number(existingCount) !== Number(newCount)) {
      const colIndex = headers.indexOf(dateShift);
      if (colIndex !== -1) { sheet.getRange(2, colIndex + 1).setValue(newCount); results.updated++; }
    } else {
      results.skipped++;
    }
  });
  return results;
}

// ========== 测试函数 ==========
function testSyncProductionData() {
  try {
    Logger.log("=== 开始测试 ===");
    const targetDates = _pd_calculateTargetDates();
    Logger.log(`目标日期: ${JSON.stringify(targetDates)}`);
    const linePlanData = _pd_readLinePlanData(_pd_SOURCE_SHEET_ID);
    Logger.log(`读取数据: ${linePlanData.length}条`);
    if (linePlanData.length > 0) {
      Logger.log(`示例: 车间=${linePlanData[0][0]}, 日期=${linePlanData[0][1]}, AEM=${linePlanData[0][4]}, 机器状况=${linePlanData[0][6]}, 班次=${linePlanData[0][9]}`);
    }
    const machineCountData = _pd_calculateMachineCount(linePlanData, targetDates);
    Object.entries(machineCountData.TB1).forEach(([ds, c]) => Logger.log(`TB1 ${ds}: ${c}台`));
    Object.entries(machineCountData.TB2).forEach(([ds, c]) => Logger.log(`TB2 ${ds}: ${c}台`));
    Logger.log("=== 测试完成 ===");
  } catch (err) {
    Logger.log(`测试失败: ${err.message}`);
  }
}
