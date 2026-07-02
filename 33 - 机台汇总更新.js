// V20260702.01 — 机台汇总更新（从 IM/TF/PK 三个数据源汇总活跃机台到 Database_MasterData）
// 入口：workcenter_update（每周四 10:31 定时 or 手动）
// 逻辑：读取三个工序的 Workcenter + Workshop → 合并 → 全量写入 "Workcenter_V202602"
// 迁移自：Database_MasterData code.gs（scriptId: 168P7bXzcwBS9b9Fx9u6e-m52zAusUwMED4aOPeCGZLbsuHVIlKT9rpT1）

// ========== 数据源配置 ==========
const WC_SOURCES = [
  {
    id: "11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc",
    name: "IM",
    process: "IM",
    sheetName: "1. Line Database",
    workcenterCol: 4,   // D列
    workshopCol: 2,     // B列
    startRow: 2,
  },
  {
    id: "1Wm_-6j60ZD8KrUMa1gjzuUlRzky1zmc3hrX_u0ZFxT0",
    name: "TF",
    process: "TF",
    sheetName: "1. Active Machine",
    workcenterCol: 4,   // D列
    workshopCol: 2,     // B列
    startRow: 3,
  },
  {
    id: "1736u5O21vH6Qhw6Uesqlg-uyZ955blUbSRST3xtQGZ8",
    name: "PK",
    process: "PK",
    sheetName: "1. Active Machine",
    workcenterCol: 3,   // C列
    workshopCol: 1,     // A列
    startRow: 3,
  },
];

const WC_TARGET = {
  ID: "1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U",
  SHEET_NAME: "Workcenter_V202602",
};

const WC_FUNC_NAME = "workcenter_update";

// ========== 主入口 ==========
function workcenter_update(e) {
  const trigger = e ? "定时" : "手动";
  const startTime = new Date();
  let totalSourceRecords = 0;
  let totalProcessedRecords = 0;

  try {
    writeLog(WC_FUNC_NAME, "成功", "开始执行机台汇总更新", trigger, "");
    const mergedData = [];

    for (let i = 0; i < WC_SOURCES.length; i++) {
      const src = WC_SOURCES[i];
      writeLog(WC_FUNC_NAME, "成功", "正在处理数据源: " + src.name, trigger, "");

      try {
        const ss = SpreadsheetApp.openById(src.id);
        const ws = ss.getSheetByName(src.sheetName);
        if (!ws) {
          throw new Error("无法找到工作表 \"" + src.sheetName + "\" 在数据源 " + src.name);
        }

        const lastRow = ws.getLastRow();
        if (lastRow < src.startRow) {
          writeLog(WC_FUNC_NAME, "成功", "数据源 " + src.name + " 无数据", trigger, "");
          continue;
        }

        const maxCol = Math.max(src.workcenterCol, src.workshopCol);
        const data = ws.getRange(src.startRow, 1, lastRow - src.startRow + 1, maxCol).getValues();
        writeLog(WC_FUNC_NAME, "成功", "数据源 " + src.name + " 读取到 " + data.length + " 条记录", trigger, "");
        totalSourceRecords += data.length;

        let validCount = 0;
        for (let j = 0; j < data.length; j++) {
          const row = data[j];
          const workcenter = (row[src.workcenterCol - 1] || "").toString().trim();
          const workshop = (row[src.workshopCol - 1] || "").toString().trim();

          if (workcenter === "" && workshop === "") continue;

          mergedData.push([workcenter, workshop, src.process]);
          validCount++;
        }
        totalProcessedRecords += validCount;
        writeLog(WC_FUNC_NAME, "成功", "数据源 " + src.name + " 处理完成，有效: " + validCount, trigger, "");

      } catch (srcErr) {
        writeLog(WC_FUNC_NAME, "失败", "处理数据源 " + src.name + " 时出错: " + srcErr.message, trigger, "");
      }
    }

    if (mergedData.length === 0) {
      writeLog(WC_FUNC_NAME, "跳过", "没有有效数据需要写入", trigger, "");
      return;
    }

    // 写入目标工作表
    const targetSS = SpreadsheetApp.openById(WC_TARGET.ID);
    let targetWS = targetSS.getSheetByName(WC_TARGET.SHEET_NAME);
    if (!targetWS) {
      targetWS = targetSS.insertSheet(WC_TARGET.SHEET_NAME);
      writeLog(WC_FUNC_NAME, "成功", "创建新工作表 \"" + WC_TARGET.SHEET_NAME + "\"", trigger, "");
    }

    const lastRow = targetWS.getLastRow();
    if (lastRow > 1) targetWS.getRange(2, 1, lastRow - 1, 3).clear();
    if (lastRow === 0) {
      targetWS.getRange(1, 1, 1, 3).setValues([["Workcenter", "Workshop", "Process"]]);
    }
    targetWS.getRange(2, 1, mergedData.length, 3).setValues(mergedData);

    const endTime = new Date();
    const duration = ((endTime.getTime() - startTime.getTime()) / 1000).toFixed(1);
    writeLog(WC_FUNC_NAME, "成功",
      "机台汇总更新完成，耗时: " + duration + "秒，源: " + totalSourceRecords + "，处理: " + totalProcessedRecords + "，输出: " + mergedData.length,
      trigger, "");

  } catch (err) {
    writeLog(WC_FUNC_NAME, "失败", err.message, trigger, err.stack || "");
  }
}
