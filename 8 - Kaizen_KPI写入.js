// V20260530.01 — Kaizen KPI 数据写入（清空后按 EQU 部门+保养组人员 分组写入）

const KAIZEN_SOURCE_SS_ID = "1R0O_GfCQyQWXJ6ZR5AfE1xgCOuPrMP3FFOZ6u3I9rLg";
const KAIZEN_DEST_SS_ID   = "1eeG8wA9YbCZfNv6qndDKa3vs7VKs2MT2fFG12Tk-At4";
const KAIZEN_SOURCE_SHEET = "Kaizen_Year";
const KAIZEN_DEST_SHEET   = "Master_Data";
const KAIZEN_STAFF_SHEET  = "保养组人员名单";

/** 读取 Kaizen_Year，过滤 EQU+保养组人员，按月份+工序分组统计后清空重写 Master_Data */
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

    // 3. 清空 Master_Data 数据区（保留表头第1行），重新写入
    const destSheet = destSS.getSheetByName(KAIZEN_DEST_SHEET);
    const lastRow = destSheet.getLastRow();
    if (lastRow > 1) {
      destSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }

    const keys = Object.keys(groups).sort();
    if (keys.length === 0) {
      try { writeLog("writeKaizenKPI", "跳过", "无符合条件的数据", trigger, ""); } catch (err) {}
      return;
    }

    const output = keys.map(function(k) {
      const g = groups[k];
      return [g.month, g.process, g.total, g.excellent, ""];
    });
    destSheet.getRange(2, 1, output.length, 5).setValues(output);

    try { writeLog("writeKaizenKPI", "成功", "写入 " + output.length + " 行到 " + KAIZEN_DEST_SHEET, trigger, ""); } catch (err) {}
  } catch (err) {
    try { writeLog("writeKaizenKPI", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
