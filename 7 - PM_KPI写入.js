// V20260526.1 — PM KPI 数据写入

const PM_SOURCE_SS_ID = "1C7z16Ww22dfGxeqKRrBw27J-zmXXStAt2xoEYwL-V24";
const PM_DEST_SS_ID   = "1KYRwzlYZz9OK8NPAkndhc9MxGlLEILYsRbsMkSvcV_U";
const PM_KPI_SHEET    = "3.PM_KPI";

/** 从 Master Data 计算并追加 PM KPI 数据到 3.PM_KPI（幂等） */
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

    // 3. PM_Target：Year_PM_Target_ID → Target（G列）
    const ptRows = destSS.getSheetByName("PM_Target").getDataRange().getValues();
    const ptMap = {};
    for (let i = 1; i < ptRows.length; i++) {
      const id = String(ptRows[i][0]).trim();
      if (id) ptMap[id] = ptRows[i][6];
    }

    // 4. 已存在的 Monthly_PM_ID（幂等去重）
    const kpiSheet = destSS.getSheetByName(PM_KPI_SHEET);
    const kpiRows  = kpiSheet.getDataRange().getValues();
    const existing = {};
    for (let i = 0; i < kpiRows.length; i++) {
      const id = String(kpiRows[i][8] || "").trim();
      if (id) existing[id] = true;
    }

    // 5. SubProcess → Process 映射
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

    // 6. 遍历 Master Data，按 (YearMonth, Process) 分组统计
    //    groups[ym][proc] = { total, valid, sb(sumBefore), sa(sumAfter), cn(countAvg) }
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
      const ub   = r[8];                        // I: UPDT_BEFORE PM
      const ua   = r[9];                        // J: UPDT_AFTER PM

      const proc = mapProcess(sub, key, wc);
      if (!proc) continue;

      const ubN = typeof ub === "number" ? ub : parseFloat(ub);
      const uaN = typeof ua === "number" ? ua : parseFloat(ua);

      // 子工序同时累加到 TF 汇总行
      const targets = [proc];
      if (proc === "AFT" || proc === "HSL+BOU" || proc === "ZAH" || proc === "_TF_OTHER") {
        targets.push("TF");
      }

      targets.forEach(function(p) {
        if (!groups[ym])    groups[ym]    = {};
        if (!groups[ym][p]) groups[ym][p] = { total: 0, valid: 0, sb: 0, sa: 0, cn: 0 };
        const g = groups[ym][p];
        g.total++;
        if (perf === "好/Good") g.valid++;
        if (!isNaN(ubN) && !isNaN(uaN) && ubN > 0 && uaN > 0) {
          g.sb += ubN; g.sa += uaN; g.cn++;
        }
      });
    }

    // 7. 构造待写入行（跳过已存在的 Monthly_PM_ID）
    const ORDER = ["IM", "AFT", "HSL+BOU", "ZAH", "TF", "PK"];
    const newRows = [];

    Object.keys(groups).sort().forEach(function(ym) {
      ORDER.forEach(function(proc) {
        const g = groups[ym][proc];
        if (!g) return;
        const mid   = ym + "_" + proc;
        if (existing[mid]) return;
        const year  = String(ym).substring(0, 4);
        const ytid  = year + "_" + proc;
        const kpiPct = g.total > 0 ? g.valid / g.total : 0;
        const pmPerf = g.cn > 0 ? g.sb / g.sa : "";
        newRows.push([ym, proc, g.total, g.valid, kpiPct, pmPerf, year, ptMap[ytid] || "", mid, ytid]);
      });
    });

    if (newRows.length === 0) {
      try { writeLog("writePMKPI", "跳过", "无新数据需要写入", trigger, ""); } catch (err) {}
      return;
    }

    // 8. 追加写入 + 设置百分比格式（E/F/H列）
    const lastRow = kpiSheet.getLastRow();
    kpiSheet.getRange(lastRow + 1, 1, newRows.length, 10).setValues(newRows);
    kpiSheet.getRange(lastRow + 1, 5, newRows.length, 1).setNumberFormat("0%");
    kpiSheet.getRange(lastRow + 1, 6, newRows.length, 1).setNumberFormat("0.00%");
    kpiSheet.getRange(lastRow + 1, 8, newRows.length, 1).setNumberFormat("0.00%");

    try { writeLog("writePMKPI", "成功", "写入 " + newRows.length + " 行到 " + PM_KPI_SHEET, trigger, ""); } catch (err) {}
  } catch (err) {
    try { writeLog("writePMKPI", "失败", err.message, trigger, err.stack || ""); } catch (e) {}
    console.error(err.stack || err.message);
  }
}
