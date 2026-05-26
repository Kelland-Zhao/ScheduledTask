// V20260526.2 — PM KPI 数据写入（仅写 A:D 列，E:J 由 ArrayFormula 自动计算）

const PM_SOURCE_SS_ID = "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24";
const PM_DEST_SS_ID   = "1KYRwzlYZz9OK8NPAkndhc9MxGlLEILYsRbsMkSvcV_U";
const PM_KPI_SHEET    = "3.PM_KPI";

/** 从 Master Data 计算并追加 PM KPI 数据到 3.PM_KPI（幂等，仅写 A:D） */
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

    // 3. 已存在的 Month+Process（幂等去重）；扫描 A:B 列定位最后数据行
    //    注：E:J 为 ArrayFormula，不能用 getLastRow() 定位，需扫描原始数据列
    const kpiSheet = destSS.getSheetByName(PM_KPI_SHEET);
    const abVals   = kpiSheet.getRange("A:B").getValues();
    const existing = {};
    let lastDataRow = 3; // 默认：header 在第3行，数据从第4行开始
    for (let i = 3; i < abVals.length; i++) { // 从第4行(index=3)开始，跳过 header
      const a = String(abVals[i][0] || "").trim();
      const b = String(abVals[i][1] || "").trim();
      if (a && b) {
        existing[a + "_" + b] = true;
        lastDataRow = i + 1; // 转为 1-indexed 行号
      }
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
    //    groups[ym][proc] = { total, valid }
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

      // 子工序同时累加到 TF 汇总行
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

    // 6. 构造待写入行（仅 A:D，跳过已存在的 Month+Process）
    const ORDER = ["IM", "AFT", "HSL+BOU", "ZAH", "TF", "PK"];
    const newRows = [];

    Object.keys(groups).sort().forEach(function(ym) {
      ORDER.forEach(function(proc) {
        const g = groups[ym][proc];
        if (!g) return;
        if (existing[ym + "_" + proc]) return;
        newRows.push([Number(ym), proc, g.total, g.valid]);
      });
    });

    if (newRows.length === 0) {
      try { writeLog("writePMKPI", "跳过", "无新数据需要写入", trigger, ""); } catch (err) {}
      return;
    }

    // 7. 追加写入 A:D 列（E:J 由 ArrayFormula 自动填充）
    kpiSheet.getRange(lastDataRow + 1, 1, newRows.length, 4).setValues(newRows);

    try { writeLog("writePMKPI", "成功", "写入 " + newRows.length + " 行到 " + PM_KPI_SHEET, trigger, ""); } catch (err) {}
  } catch (err) {
    try { writeLog("writePMKPI", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
