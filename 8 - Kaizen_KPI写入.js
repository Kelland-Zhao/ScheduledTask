// V20260530.03 — Kaizen KPI 数据写入（upsert：已有行更新 C/D，新组合追加，保留历史）

const KAIZEN_SOURCE_SS_ID = "1R0O_GfCQyQWXJ6ZR5AfE1xgCOuPrMP3FFOZ6u3I9rLg";
const KAIZEN_DEST_SS_ID   = "1eeG8wA9YbCZfNv6qndDKa3vs7VKs2MT2fFG12Tk-At4";
const KAIZEN_SOURCE_SHEET = "Kaizen_Year";
const KAIZEN_DEST_SHEET   = "Master_Data";
const KAIZEN_STAFF_SHEET  = "保养组人员名单";

/** 读取 Kaizen_Year，过滤 EQU+保养组人员，按月份+工序分组统计后 upsert 到 Master_Data */
function writeKaizenKPI(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const srcSS  = SpreadsheetApp.openById(KAIZEN_SOURCE_SS_ID);
    const destSS = SpreadsheetApp.openById(KAIZEN_DEST_SS_ID);

    // 1. 读取保养组人员名单，建立工号白名单
    const staffRows = destSS.getSheetByName(KAIZEN_STAFF_SHEET).getDataRange().getValues();
    const staffSet = {};
    for (let i = 1; i < staffRows.length; i++) {
      const id = String(staffRows[i][0] || "").trim();
      if (id.length >= 5) staffSet[id.slice(-5)] = true;
    }

    // 2. 读取 Kaizen_Year，过滤并分组统计
    //    列索引：0=优秀备注, 2=月份, 9=部门, 11=工序, 14=项目负责人工号
    const kaizenRows = srcSS.getSheetByName(KAIZEN_SOURCE_SHEET).getDataRange().getValues();
    const groups = {};

    for (let i = 1; i < kaizenRows.length; i++) {
      const r         = kaizenRows[i];
      const excellent = String(r[0]  || "").trim();
      const month     = String(r[2]  || "").trim();
      const dept      = String(r[9]  || "").trim();
      const process   = String(r[11] || "").trim();
      const empId     = String(r[14] || "").trim();

      if (!month || !process) continue;
      if (dept !== "EQU") continue;
      if (!staffSet[empId]) continue;

      const key = month + "_" + process;
      if (!groups[key]) groups[key] = { month: month, process: process, total: 0, excellent: 0 };
      groups[key].total++;
      if (excellent === "优秀") groups[key].excellent++;
    }

    if (Object.keys(groups).length === 0) {
      try { writeLog("writeKaizenKPI", "跳过", "无符合条件的数据", trigger, ""); } catch (err) {}
      return;
    }

    // 3. 读取 Master_Data 现有行，建立 "月份_工序" → 1-indexed 行号 map
    const destSheet = destSS.getSheetByName(KAIZEN_DEST_SHEET);
    const existingVals = destSheet.getRange("A:B").getValues();
    const existingMap = {};
    for (let i = 1; i < existingVals.length; i++) {
      const a = String(existingVals[i][0] || "").trim();
      const b = String(existingVals[i][1] || "").trim();
      if (a && b) existingMap[a + "_" + b] = i + 1; // 1-indexed
    }

    // 4. Upsert：已有行更新 C/D，新组合追加整行
    let updateCount = 0;
    let appendCount = 0;
    const appendRows = [];

    Object.keys(groups).forEach(function(k) {
      const g = groups[k];
      if (existingMap[k]) {
        destSheet.getRange(existingMap[k], 3, 1, 2).setValues([[g.total, g.excellent]]);
        updateCount++;
      } else {
        appendRows.push([g.month, g.process, g.total, g.excellent, ""]);
        appendCount++;
      }
    });

    if (appendRows.length > 0) {
      destSheet.getRange(destSheet.getLastRow() + 1, 1, appendRows.length, 5).setValues(appendRows);
    }

    try { writeLog("writeKaizenKPI", "成功", "更新 " + updateCount + " 行，追加 " + appendCount + " 行到 " + KAIZEN_DEST_SHEET, trigger, ""); } catch (err) {}
  } catch (err) {
    try { writeLog("writeKaizenKPI", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
