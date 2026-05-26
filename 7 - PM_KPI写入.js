// V20260526.3 — PM KPI 数据写入（按 A/B 匹配后写入 C/D，E:J 为 ArrayFormula）

const PM_SOURCE_SS_ID = "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24";
const PM_DEST_SS_ID   = "1KYRwzlYZz9OK8NPAkndhc9MxGlLEILYsRbsMkSvcV_U";
const PM_KPI_SHEET    = "3.PM_KPI";

/** 从 Master Data 计算后，按 3.PM_KPI A/B 列匹配将结果写入 C/D 列（跳过已有值的行） */
function writePMKPI(e) {
  const trigger = e ? "定时" : "手动";
  try {
    const srcSS  = SpreadsheetApp.openById(PM_SOURCE_SS_ID);
    const destSS = SpreadsheetApp.openById(PM_DEST_SS_ID);

    // 1. KPI排除条件：状态=启用 的 Workcenter 黑名单
    const excludeRows = srcSS.getSheetByName("KPI排除条件").getDataRange().getValues();
    const excluded = {};
    for (let i = 1; i < excludeRows.length; i++) {
      if (String(excludeRows[i][3]).trim() === "启用") {
        excluded[String(excludeRows[i][0]).trim()] = true;
      }
    }

    // 2. TF Machine Type：Workcenter → Machine Type_修正（SubProcess="TF" 时补充分类）
    const tfRows = srcSS.getSheetByName("TF Machine Type").getDataRange().getValues();
    const tfTypeMap = {};
    for (let i = 1; i < tfRows.length; i++) {
      tfTypeMap[String(tfRows[i][0]).trim()] = String(tfRows[i][2]).trim();
    }

    // 3. 读取 3.PM_KPI A:D，找到 A/B 有值但 C 为空的待写入行
    //    pendingRows["YYYYMM_Process"] = 1-indexed 行号
    const kpiSheet  = destSS.getSheetByName(PM_KPI_SHEET);
    const abcdVals  = kpiSheet.getRange("A:D").getValues();
    const pendingRows = {};
    for (let i = 0; i < abcdVals.length; i++) {
      const a = String(abcdVals[i][0] || "").trim();
      const b = String(abcdVals[i][1] || "").trim();
      const c = abcdVals[i][2];
      if (a && b && !c) pendingRows[a + "_" + b] = i + 1; // 1-indexed
    }

    if (Object.keys(pendingRows).length === 0) {
      try { writeLog("writePMKPI", "跳过", "无待写入行（C列均已有值）", trigger, ""); } catch (err) {}
      return;
    }

    // 4. SubProcess → Process 映射
    function mapProcess(sub, key, wc) {
      if (key === "IM" || sub === "IM") return "IM";
      if (key === "PK" || sub === "PK") return "PK";
      if (sub === "AFT & USW")          return "AFT";
      if (sub === "HSL & Bou")          return "HSL+BOU";
      if (sub === "Zar")                return "ZAH";
      if (sub === "TF") {
        const t = tfTypeMap[wc] || "";
        if (t.indexOf("AFT") >= 0 || t.indexOf("USW") >= 0) return "AFT";
        if (t.indexOf("HSL") >= 0 || t.indexOf("Bou") >= 0) return "HSL+BOU";
        if (t.indexOf("Zar") >= 0) return "ZAH";
        return "_TF_OTHER";
      }
      return null;
    }

    // 5. 遍历 Master Data，按 (YearMonth, Process) 分组统计
    const masterRows = srcSS.getSheetByName("Master Data").getDataRange().getValues();
    const groups = {};

    for (let i = 1; i < masterRows.length; i++) {
      const r = masterRows[i];
      if (!r[0]) continue;
      const key = String(r[1]  || "").trim();  // B: KeyProcess
      const wc  = String(r[3]  || "").trim();  // D: Workcenter
      if (excluded[wc]) continue;
      const ym  = String(r[10] || "").trim();  // K: YearMonth
      if (!ym) continue;
      const sub  = String(r[11] || "").trim(); // L: SubProcess
      const perf = String(r[13] || "").trim(); // N: 绩效状态

      const proc = mapProcess(sub, key, wc);
      if (!proc) continue;

      const targets = [proc];
      if (proc === "AFT" || proc === "HSL+BOU" || proc === "ZAH" || proc === "_TF_OTHER") {
        targets.push("TF");
      }

      targets.forEach(function(p) {
        if (!groups[ym])    groups[ym]    = {};
        if (!groups[ym][p]) groups[ym][p] = { total: 0, valid: 0 };
        groups[ym][p].total++;
        if (perf === "好/Good") groups[ym][p].valid++;
      });
    }

    // 6. 按匹配行号写入 C/D（YearMonth 固定6位，_后为 Process）
    let writeCount = 0;
    Object.keys(pendingRows).forEach(function(k) {
      const ym   = k.substring(0, 6);
      const proc = k.substring(7);
      const g = groups[ym] && groups[ym][proc];
      if (!g) return;
      kpiSheet.getRange(pendingRows[k], 3, 1, 2).setValues([[g.total, g.valid]]);
      writeCount++;
    });

    if (writeCount === 0) {
      try { writeLog("writePMKPI", "跳过", "Master Data 中无对应数据", trigger, ""); } catch (err) {}
      return;
    }

    try { writeLog("writePMKPI", "成功", "写入 " + writeCount + " 行到 " + PM_KPI_SHEET, trigger, ""); } catch (err) {}
  } catch (err) {
    try { writeLog("writePMKPI", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
