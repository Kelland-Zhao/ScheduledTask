// V20260613.02 — 工时数据汇总（TB1+TB2 → MasterData，合并实际工时）
// 入口：aggregateTB1TB2ToMasterData（每日 09:20 定时 or 手动）
// 数据源：_pd_TARGET_SHEET_ID Sheets: TB1, TB2
//         _pd_ATTENDANCE_SHEET_ID Sheet: 跟班考勤SUM
// 目标表：_pd_TARGET_SHEET_ID Sheet: MasterData
// 新增：开机数>0但人员排班缺失时，邮件提醒李华、丁志伟（CC直线上级+上级的上级）

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

    const gaps = [];
    const tb1Data = _pd_convertSheetToMasterData(tb1Sheet, "TB1", gaps);
    const tb2Data = _pd_convertSheetToMasterData(tb2Sheet, "TB2", gaps);
    logDetails.push(`TB1: ${tb1Data.length}条, TB2: ${tb2Data.length}条`);

    // 开机数>0 但人员排班缺失 → 发送提醒邮件
    if (gaps.length > 0) {
      _pd_sendPersonnelGapAlert(gaps, trigger);
      logDetails.push(`排班缺失提醒: ${gaps.length}个班次`);
    }

    const allData = tb1Data.concat(tb2Data);
    if (allData.length === 0) {
      if (gaps.length > 0) {
        writeLog("aggregateTB1TB2ToMasterData", "跳过", "开机数>0但无人员排班（已发送提醒）", trigger, "");
        return;
      }
      throw new Error("没有有效数据需要同步");
    }

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
function _pd_convertSheetToMasterData(sheet, workshop, gaps) {
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

    let hasPersonnel = false;
    for (let rowIndex = 2; rowIndex < allData.length; rowIndex++) {
      const personName = allData[rowIndex][0];
      const manhour = allData[rowIndex][colIndex];
      if (!personName || !manhour || Number(manhour) <= 0) continue;
      hasPersonnel = true;
      result.push([dateShift, workshop, operatingQty, personName, manhour, monthStr, weekNum, ""]);
    }

    // 开机数>0 但无任何人员填写工时 → 记录排班缺失
    if (!hasPersonnel && gaps) {
      gaps.push({ workshop: workshop, dateShift: String(dateShift) });
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

// ========== 排班缺失提醒 ==========
/** 构建 userID 查找表：name→{email,lineMgr} 和 email→{name,lineMgr} */
function _pd_buildUserLookup() {
  const map = { nameMap: {}, emailMap: {} };
  try {
    const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet) return map;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return map;

    const data = sheet.getRange(1, 1, lastRow, 61).getValues();
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const name = String(row[1] || "").trim();
      const email = String(row[9] || "").trim().toLowerCase();
      const lineMgr = String(row[60] || "").trim().toLowerCase();
      if (name && email) {
        map.nameMap[name] = { email: email, lineMgr: lineMgr };
      }
      if (email) {
        map.emailMap[email] = { name: name, lineMgr: lineMgr };
      }
    }
  } catch (err) {
    console.error("_pd_buildUserLookup 失败: " + err.message);
  }
  return map;
}

/** 提取日期部分 "2026.06.14_1夜" → "2026.06.14" */
function _pd_extractDateFromShift(dateShift) {
  var idx = String(dateShift).lastIndexOf("_");
  return idx >= 0 ? String(dateShift).substring(0, idx) : String(dateShift);
}

/** 提取班次中文显示 "1夜"→"夜班" "2早"→"早班" "3中"→"中班" */
function _pd_shiftDisplay(shiftPart) {
  var map = { "1夜": "夜班", "2早": "早班", "3中": "中班" };
  return map[shiftPart] || shiftPart;
}

/** 发送排班缺失提醒邮件 */
function _pd_sendPersonnelGapAlert(gaps, trigger) {
  try {
    const { nameMap, emailMap } = _pd_buildUserLookup();

    // 目标收件人
    const targetNames = ["李华", "丁志伟"];
    const toEmails = [];
    const ccSet = {};

    targetNames.forEach(function(name) {
      const user = nameMap[name];
      if (!user || !user.email) {
        console.warn("userID 中未找到: " + name);
        return;
      }
      toEmails.push(user.email);

      // 直线上级 (BI列)
      if (user.lineMgr && !ccSet[user.lineMgr]) {
        ccSet[user.lineMgr] = true;
      }

      // 直线上级的直线上级
      const supData = emailMap[user.lineMgr];
      if (supData && supData.lineMgr && !ccSet[supData.lineMgr]) {
        ccSet[supData.lineMgr] = true;
      }
    });

    if (toEmails.length === 0) {
      console.warn("未找到李华或丁志伟的邮箱，跳过排班缺失提醒");
      return;
    }

    // 去掉和 TO 重复的 CC
    toEmails.forEach(function(e) { delete ccSet[e]; });
    var ccEmails = Object.keys(ccSet);

    // 构建 gap 行（日期 | 车间 | 班次），按日期+车间排序
    var gapRows = "";
    gaps.sort(function(a, b) {
      var d = String(a.dateShift).localeCompare(String(b.dateShift));
      if (d !== 0) return d;
      return a.workshop.localeCompare(b.workshop);
    });
    gaps.forEach(function(g) {
      var dateStr = _pd_extractDateFromShift(g.dateShift);
      var shiftPart = _pd_splitShift(g.dateShift);
      gapRows += "<tr><td style='padding:6px 12px;border:1px solid #ddd'>" + dateStr + "</td>" +
        "<td style='padding:6px 12px;border:1px solid #ddd'>" + g.workshop + "</td>" +
        "<td style='padding:6px 12px;border:1px solid #ddd'>" + _pd_shiftDisplay(shiftPart) + "</td></tr>";
    });

    var sheetUrl = "https://docs.google.com/spreadsheets/d/" + _pd_TARGET_SHEET_ID;

    var html = '<div style="font-family:Arial,\'Microsoft YaHei\',sans-serif;max-width:600px">' +
      '<h3 style="color:#E60012">⚠ 注塑工序人员排班缺失提醒</h3>' +
      '<p>以下日期班次已安排<strong>开机</strong>，但<strong>人员排班尚未填写</strong>，请及时处理：</p>' +
      '<p style="font-size:13px">📊 <a href="' + sheetUrl + '" style="color:#E60012">打开排班表</a></p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
      '<tr style="background:#E60012;color:white"><th style="padding:8px;text-align:left">日期</th><th style="padding:8px;text-align:left">车间</th><th style="padding:8px;text-align:left">班次</th></tr>' +
      gapRows +
      '</table>' +
      '<p style="color:#888;font-size:12px;margin-top:16px">此邮件由 工时数据汇总系统 自动发送</p>' +
      '</div>';

    var subject = "【排班缺失提醒】 注塑工序人员排班未填写";
    var options = { htmlBody: html, name: "工时数据汇总系统" };
    if (ccEmails.length > 0) options.cc = ccEmails.join(",");

    GmailApp.sendEmail(toEmails.join(","), subject, "", options);
    console.log("排班缺失提醒已发送 → TO: " + toEmails.join(",") + " CC: " + ccEmails.join(","));

  } catch (err) {
    console.error("发送排班缺失提醒失败: " + err.message);
    try { writeLog("_pd_sendPersonnelGapAlert", "失败", err.message, trigger, ""); } catch (e2) {}
  }
}

/** 从日期班次提取班次部分 */
function _pd_splitShift(dateShift) {
  var idx = String(dateShift).lastIndexOf("_");
  return idx >= 0 ? String(dateShift).substring(idx + 1) : "";
}
