// V20260609.01 — 季度优秀Kaizen同步（源表→Master_Data，含工号和Sub_Process匹配）

const KQ_SOURCE_SS_ID = "1TQV8QUz3PWWyfEGGcWydRc7HTkw7WpJwqM7sLEATF7A";
const KQ_DEST_SS_ID   = "1eeG8wA9YbCZfNv6qndDKa3vs7VKs2MT2fFG12Tk-At4";
const KQ_SOURCE_SHEET = "Kaizen季度评选获奖";
const KQ_DEST_SHEET   = "季度优秀Kaizen记录";
const KQ_STAFF_SHEET  = "保养组人员名单";

/** 从「月度优秀Kaizen验证与推广系统」同步季度评选获奖数据到 Kaizen_Master_Data，补充工号和Sub_Process */
function writeKaizenQuarterly(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const srcSS  = SpreadsheetApp.openById(KQ_SOURCE_SS_ID);
    const destSS = SpreadsheetApp.openById(KQ_DEST_SS_ID);

    // 1. 读取保养组人员名单，建立 姓名 → {工号, Sub_Process} map
    const staffRows = destSS.getSheetByName(KQ_STAFF_SHEET).getDataRange().getValues();
    const staffMap  = {};
    for (let i = 1; i < staffRows.length; i++) {
      const id   = String(staffRows[i][0] || "").trim();
      const name = String(staffRows[i][1] || "").trim();
      const sub  = String(staffRows[i][3] || "").trim();
      if (name) staffMap[name] = { id: id, subProcess: sub };
    }

    // 2. 读取源表（跳过表头），列：时间(0),序号(1),提名部门(2),工序(3),项目名称(4),项目负责人(5)
    const srcRows = srcSS.getSheetByName(KQ_SOURCE_SHEET).getDataRange().getValues();

    // 3. 读取目标表现有行，建立 "时间_序号" → 1-indexed 行号 map
    const destSheet    = destSS.getSheetByName(KQ_DEST_SHEET);
    const destLastRow  = Math.max(destSheet.getLastRow(), 1);
    const existingVals = destSheet.getRange(1, 1, destLastRow, 2).getValues();
    const existingMap  = {};
    for (let i = 1; i < existingVals.length; i++) {
      const key = String(existingVals[i][0]).trim() + "_" + String(existingVals[i][1]).trim();
      if (key !== "_") existingMap[key] = i + 1;
    }

    // 4. Upsert
    let updateCount  = 0;
    let appendCount  = 0;
    const appendRows = [];
    const unmatched  = [];

    for (let i = 1; i < srcRows.length; i++) {
      const r      = srcRows[i];
      const 时间   = String(r[0] || "").trim();
      const 序号   = String(r[1] || "").trim();
      const 部门   = String(r[2] || "").trim();
      const 工序   = String(r[3] || "").trim();
      const 名称   = String(r[4] || "").trim();
      const 负责人 = String(r[5] || "").trim();

      if (!时间 || !序号) continue;

      const staff      = staffMap[负责人];
      const 工号       = staff ? staff.id        : "";
      const subProcess = staff ? staff.subProcess : "";

      if (!staff && 负责人) unmatched.push(负责人);

      const row = [时间, 序号, 部门, 工序, 名称, 负责人, 工号, subProcess];
      const key = 时间 + "_" + 序号;

      if (existingMap[key]) {
        destSheet.getRange(existingMap[key], 1, 1, 8).setValues([row]);
        updateCount++;
      } else {
        appendRows.push(row);
        appendCount++;
      }
    }

    if (appendRows.length > 0) {
      const startRow = destSheet.getLastRow() + 1;
      destSheet.getRange(startRow, 1, appendRows.length, 8).setValues(appendRows);
    }

    try { writeLog("writeKaizenQuarterly", "成功", "更新 " + updateCount + " 行，追加 " + appendCount + " 行", trigger, ""); } catch (err) {}

    if (unmatched.length > 0) {
      const seen = {};
      const uniqueUnmatched = [];
      unmatched.forEach(function(n) { if (!seen[n]) { seen[n] = true; uniqueUnmatched.push(n); } });
      try { writeLog("writeKaizenQuarterly", "警告", "以下人员在保养组名单中未找到，工号/Sub_Process 留空：" + uniqueUnmatched.join("、"), trigger, ""); } catch (err) {}
    }

  } catch (err) {
    try { writeLog("writeKaizenQuarterly", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
