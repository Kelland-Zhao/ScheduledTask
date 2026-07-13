// V20260713.01 — 工时数据汇总（开机数 + AttendanceSync → MasterData）
// 入口：aggregateTB1TB2ToMasterData（每日 09:20 定时 or 手动）
// 数据源：_pd_TARGET_SHEET_ID Sheet: 开机数（开机数）
//         _pd_ATTENDANCE_SYNC_ID Sheet: AttendanceSync（人员出勤）
// 目标表：_pd_TARGET_SHEET_ID Sheet: MasterData
// 过滤：车间=ALL 跳过，工序≠INJ 跳过

const _pd_ATTENDANCE_SYNC_ID = "1UBg1Ake18cFp6gj0jKRX1Y9GJ0VL1pY5aXK-UoCeAY0";

// ========== 主入口 ==========
function aggregateTB1TB2ToMasterData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const startTime = new Date();
    const logDetails = [];

    const targetSpreadsheet = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    const mcSheet = targetSpreadsheet.getSheetByName("开机数");
    const masterDataSheet = targetSpreadsheet.getSheetByName("MasterData");
    if (!mcSheet) throw new Error("未找到'开机数'工作表");

    const allData = _pd_readAttendanceData(mcSheet);
    logDetails.push(`出勤记录: ${allData.length}条`);

    if (allData.length === 0) {
      throw new Error("没有有效数据需要同步");
    }

    const stats = _pd_incrementalUpdateMasterData(masterDataSheet, allData);
    if (stats.added + stats.updated === 0) {
      logDetails.push(`MasterData无变化（跳过${stats.skipped}条）`);
    } else {
      logDetails.push(`MasterData: 新增${stats.added}条, 更新${stats.updated}条, 跳过${stats.skipped}条`);
    }

    const duration = ((new Date()) - startTime) / 1000;
    writeLog("aggregateTB1TB2ToMasterData", "成功", logDetails.join("; "), trigger, `耗时${duration}s`);

  } catch (err) {
    try { writeLog("aggregateTB1TB2ToMasterData", "失败", err.message, trigger, ""); } catch (e2) {}
    throw err;
  }
}

// ========== 辅助函数 ==========

/** 班次映射：AttendanceSync → dateShift 后缀 */
const _pd_SHIFT_MAP = { "早班": "2早", "中班": "3中", "夜班": "1夜" };

/** 从 开机数 sheet + AttendanceSync 读取并合并数据 */
function _pd_readAttendanceData(mcSheet) {
  const result = [];

  // 1. 读取开机数 → Map<dateShift, {tb1, tb2}>
  const mcLastRow = mcSheet.getLastRow();
  const mcMap = new Map();
  if (mcLastRow > 1) {
    const mcData = mcSheet.getRange(2, 1, mcLastRow - 1, 3).getValues();
    mcData.forEach(row => {
      const ds = String(row[0]).trim();
      if (ds) mcMap.set(ds, { tb1: Number(row[1]) || 0, tb2: Number(row[2]) || 0 });
    });
  }

  // 2. 读取 AttendanceSync
  const attSheet = SpreadsheetApp.openById(_pd_ATTENDANCE_SYNC_ID).getSheetByName("AttendanceSync");
  if (!attSheet) { Logger.log("未找到'AttendanceSync'工作表"); return result; }
  const attLastRow = attSheet.getLastRow();
  if (attLastRow < 2) return result;
  const attData = attSheet.getRange(2, 1, attLastRow - 1, 11).getValues();

  attData.forEach(row => {
    let dateStr = "";                                    // A: 日期（Date对象→字符串）
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], "Asia/Shanghai", "yyyy-MM-dd");
    } else {
      dateStr = String(row[0] || "").trim();
    }
    const name = String(row[2] || "").trim();            // C: 姓名
    const process = String(row[3] || "").trim();         // D: 工序
    const workshop = String(row[5] || "").trim();        // F: 车间
    const shift = String(row[6] || "").trim();           // G: 班次
    const hours = Number(row[7]) || 0;                   // H: 工时

    if (!dateStr || !name || hours <= 0) return;
    if (process !== "INJ") return;
    if (workshop !== "TB1" && workshop !== "TB2") return;

    const shiftSuffix = _pd_SHIFT_MAP[shift];
    if (!shiftSuffix) return;

    const dateShift = dateStr.replace(/-/g, ".") + "_" + shiftSuffix;
    const mc = mcMap.get(dateShift);
    const operatingQty = mc ? (workshop === "TB1" ? mc.tb1 : mc.tb2) : 0;
    if (operatingQty <= 0) return;

    const monthStr = _pd_extractMonth(dateShift);
    const weekNum = _pd_calculateWeek(dateShift);
    result.push([dateShift, workshop, operatingQty, name, hours, monthStr, weekNum]);
  });

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

function _pd_incrementalUpdateMasterData(sheet, data) {
  const stats = { updated: 0, added: 0, skipped: 0 };
  if (data.length === 0) return stats;

  const lastRow = sheet.getLastRow();
  const existingData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];
  const existingMap = new Map();
  existingData.forEach((row, i) => existingMap.set(`${row[0]}_${row[3]}`, { rowIndex: i + 2, data: row }));

  const updateOps = [], newRecords = [];
  data.forEach(newRow => {
    const key = `${newRow[0]}_${newRow[3]}`;
    if (existingMap.has(key)) {
      const oldRow = existingMap.get(key).data;
      // 比较7列是否完全一致
      const changed = oldRow.some((val, i) => String(val) !== String(newRow[i]));
      if (changed) {
        updateOps.push({ row: existingMap.get(key).rowIndex, data: newRow });
      } else {
        stats.skipped++;
      }
      existingMap.delete(key);
    } else {
      newRecords.push(newRow);
    }
  });

  updateOps.forEach(op => sheet.getRange(op.row, 1, 1, 7).setValues([op.data]));
  stats.updated = updateOps.length;

  if (newRecords.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRecords.length, 7).setValues(newRecords);
  }
  stats.added = newRecords.length;

  sheet.getRange(1, 1, 1, 7).setValues([[
    "日期班次 / Date & Shift", "车间 / Workshop", "开机数 / Operating Machine Qty",
    "姓名 / Name", "安排工时 / Scheduling Manhour", "月份 / Month", "周 / Week"
  ]]);

  return stats;
}
