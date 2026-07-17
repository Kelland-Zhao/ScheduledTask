// V20260713.01 — 开机数同步（从 Line日计划 → 开机数 sheet 纵表更新）
// 入口：syncProductionData（每日 09:15 定时 or 手动）
// 数据源：_pd_SOURCE_SHEET_ID Sheet: 2.6_Line日计划
// 目标表：_pd_TARGET_SHEET_ID Sheet: 开机数

const _pd_SOURCE_SHEET_ID = "1zNOM35TeOmaZTtaAR4izQdA37cLGSgPhJAw40KkYioQ";
const _pd_TARGET_SHEET_ID = "1dyS5C7r4pqYIeRT0p1zYzngt0EDCYR4hsswurAsEBYg";

// ========== 主入口 ==========
function syncProductionData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const startTime = new Date();
    const logDetails = [];

    const targetSpreadsheet = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    const mcSheet = targetSpreadsheet.getSheetByName("开机数");
    if (!mcSheet) throw new Error("未找到'开机数'工作表");

    const targetDates = _pd_calculateTargetDates();
    const linePlanData = _pd_readLinePlanData(_pd_SOURCE_SHEET_ID);
    logDetails.push(`读取Line日计划: ${linePlanData.length}条`);

    const machineCountData = _pd_calculateMachineCount(linePlanData, targetDates);
    logDetails.push(`计算开机数: TB1(${Object.keys(machineCountData.TB1).length}), TB2(${Object.keys(machineCountData.TB2).length})`);

    const results = _pd_updateMachineCountSheet(mcSheet, machineCountData);

    const duration = ((new Date()) - startTime) / 1000;
    logDetails.push(`新增${results.created} 更新${results.updated} 跳过${results.skipped}`);

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

/** 纵表增量更新：读取现有数据 → 按 dateShift 匹配 → 更新或追加 */
function _pd_updateMachineCountSheet(sheet, machineCountData) {
  const results = { created: 0, updated: 0, skipped: 0 };
  const lastRow = sheet.getLastRow();

  // 读取现有数据，构建 dateShift → {row, tb1, tb2} Map
  const existingMap = new Map();
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    existing.forEach((row, i) => {
      const ds = String(row[0]).trim();
      if (ds) existingMap.set(ds, { row: i + 2, tb1: row[1], tb2: row[2] });
    });
  }

  // 收集所有 dateShift
  const allDateShifts = new Set();
  Object.keys(machineCountData.TB1).forEach(k => allDateShifts.add(k));
  Object.keys(machineCountData.TB2).forEach(k => allDateShifts.add(k));

  const newRows = [];
  allDateShifts.forEach(dateShift => {
    const newTb1 = machineCountData.TB1[dateShift] || 0;
    const newTb2 = machineCountData.TB2[dateShift] || 0;
    const exist = existingMap.get(dateShift);

    if (exist) {
      const oldTb1 = Number(exist.tb1) || 0;
      const oldTb2 = Number(exist.tb2) || 0;
      if (oldTb1 !== newTb1 || oldTb2 !== newTb2) {
        sheet.getRange(exist.row, 2, 1, 2).setValues([[newTb1, newTb2]]);
        results.updated++;
      } else {
        results.skipped++;
      }
    } else {
      newRows.push([dateShift, newTb1, newTb2]);
      results.created++;
    }
  });

  if (newRows.length > 0) {
    // 按 dateShift 排序后追加
    newRows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
    sheet.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }

  return results;
}

// ========== 测试函数 ==========
function testSyncProductionData() {
  try {
    Logger.log("=== 开始测试 ===");
    const targetDates = _pd_calculateTargetDates();
    Logger.log(`目标日期: ${targetDates.length}个班次`);
    const linePlanData = _pd_readLinePlanData(_pd_SOURCE_SHEET_ID);
    Logger.log(`读取Line日计划: ${linePlanData.length}条`);
    if (linePlanData.length > 0) {
      Logger.log(`示例: 车间=${linePlanData[0][0]}, 日期=${linePlanData[0][1]}, AEM=${linePlanData[0][4]}, 机器状况=${linePlanData[0][6]}, 班次=${linePlanData[0][9]}`);
    }
    const machineCountData = _pd_calculateMachineCount(linePlanData, targetDates);
    const allDateShifts = new Set();
    Object.keys(machineCountData.TB1).forEach(k => allDateShifts.add(k));
    Object.keys(machineCountData.TB2).forEach(k => allDateShifts.add(k));
    Array.from(allDateShifts).sort().forEach(ds => {
      Logger.log(`${ds}: TB1=${machineCountData.TB1[ds] || 0}台, TB2=${machineCountData.TB2[ds] || 0}台`);
    });
    Logger.log("=== 测试完成 ===");
  } catch (err) {
    Logger.log(`测试失败: ${err.message}`);
  }
}
