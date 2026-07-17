// V20260702.01 — 故障代码合并更新（从 IM/TF/PK 三个数据源合并到 Database_MasterData）
// 入口：errorCode_update（每周四 10:30 定时 or 手动）
// 逻辑：读取三个工序的 "4.Line_UPDT" → 清洗去重 → 追加转规格记录 → 全量写入 "errorCode 2.0"
// 迁移自：Database_MasterData code.gs（scriptId: 168P7bXzcwBS9b9Fx9u6e-m52zAusUwMED4aOPeCGZLbsuHVIlKT9rpT1）

// ========== 数据源配置 ==========
const EC_SOURCES = [
  { id: "1uIEoS1T1Evw0rXa3IUrew91mCffgwf-X_yQ4K4dKDYA", name: "IM", process: "IM" },
  { id: "1VrXakMRBVcsDO9fAN9JlqLyWq3Sr61dNLqQK8x8OjaM", name: "TF", process: "TF" },
  { id: "13IWm6K59Q86Yhw_X84ekTsa8i_fvrR8g6H0QuJtJvkE", name: "PK", process: "PK" },
];

const EC_TARGET = {
  ID: "1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U",
  SHEET_NAME: "errorCode 2.0",
};

const EC_SOURCE_SHEET = "4.Line_UPDT";
const EC_FUNC_NAME = "errorCode_update";

// 手动追加的转规格记录（每个工序一条）
const EC_MANUAL_RECORDS = [
  ["IM", "", "", "转规格 / Changeover", ""],
  ["TF", "", "", "转规格 / Changeover", ""],
  ["PK", "", "", "转规格 / Changeover", ""],
];

// ========== 主入口 ==========
function errorCode_update(e) {
  const trigger = e ? "定时" : "手动";
  const startTime = new Date();
  let totalSourceRecords = 0;
  let totalProcessedRecords = 0;

  try {
    writeLog(EC_FUNC_NAME, "成功", "开始执行故障代码合并更新", trigger, "");
    const mergedData = [];

    // 遍历每个数据源
    for (let i = 0; i < EC_SOURCES.length; i++) {
      const src = EC_SOURCES[i];
      writeLog(EC_FUNC_NAME, "成功", "正在处理数据源: " + src.name, trigger, "");

      try {
        const ss = SpreadsheetApp.openById(src.id);
        const ws = ss.getSheetByName(EC_SOURCE_SHEET);
        if (!ws) {
          throw new Error("无法找到工作表 \"" + EC_SOURCE_SHEET + "\" 在数据源 " + src.name);
        }

        const lastRow = ws.getLastRow();
        if (lastRow <= 1) {
          writeLog(EC_FUNC_NAME, "成功", "数据源 " + src.name + " 无数据", trigger, "");
          continue;
        }

        const data = ws.getRange(2, 1, lastRow - 1, 4).getValues();
        writeLog(EC_FUNC_NAME, "成功", "数据源 " + src.name + " 读取到 " + data.length + " 条记录", trigger, "");
        totalSourceRecords += data.length;

        let validCount = 0;
        for (let j = 0; j < data.length; j++) {
          const row = data[j];
          const machineType = row[0] || "";
          const updtEnglish = row[1] || "";
          const updtLocal = row[2] || "";
          const downtimeDriver = row[3] || "";

          // 跳过空行
          if (machineType === "" && updtEnglish === "" && updtLocal === "" && downtimeDriver === "") continue;
          // 跳过已删除记录
          if (downtimeDriver.indexOf("已删除") !== -1) continue;

          // 推断 Process
          let process = src.process;
          if (machineType.indexOf("IM") !== -1) process = "IM";
          else if (machineType.indexOf("TF") !== -1) process = "TF";
          else if (machineType.indexOf("PKG") !== -1) process = "PK";

          mergedData.push([process, machineType, updtEnglish, updtLocal, downtimeDriver]);
          validCount++;
        }
        totalProcessedRecords += validCount;
        writeLog(EC_FUNC_NAME, "成功", "数据源 " + src.name + " 处理完成，有效记录: " + validCount, trigger, "");

      } catch (srcErr) {
        writeLog(EC_FUNC_NAME, "失败", "处理数据源 " + src.name + " 时出错: " + srcErr.message, trigger, "");
      }
    }

    // 追加手动转规格记录
    mergedData.push.apply(mergedData, EC_MANUAL_RECORDS);
    totalProcessedRecords += EC_MANUAL_RECORDS.length;

    if (mergedData.length === 0) {
      writeLog(EC_FUNC_NAME, "跳过", "没有有效数据需要写入", trigger, "");
      return;
    }

    // 写入目标工作表
    const targetSS = SpreadsheetApp.openById(EC_TARGET.ID);
    let targetWS = targetSS.getSheetByName(EC_TARGET.SHEET_NAME);
    if (!targetWS) {
      targetWS = targetSS.insertSheet(EC_TARGET.SHEET_NAME);
      writeLog(EC_FUNC_NAME, "成功", "创建新工作表 \"" + EC_TARGET.SHEET_NAME + "\"", trigger, "");
    }

    const lastRow = targetWS.getLastRow();
    if (lastRow > 1) targetWS.getRange(2, 1, lastRow - 1, 5).clear();
    if (lastRow === 0) {
      targetWS.getRange(1, 1, 1, 5).setValues([["Process", "Machine_Type", "UPDT (English)", "UPDT (Local Language)", "Downtime Driver"]]);
    }
    targetWS.getRange(2, 1, mergedData.length, 5).setValues(mergedData);

    const endTime = new Date();
    const duration = ((endTime.getTime() - startTime.getTime()) / 1000).toFixed(1);
    writeLog(EC_FUNC_NAME, "成功",
      "故障代码合并更新完成，耗时: " + duration + "秒，源记录: " + totalSourceRecords + "，处理: " + totalProcessedRecords + "，输出: " + mergedData.length,
      trigger, "");

  } catch (err) {
    writeLog(EC_FUNC_NAME, "失败", err.message, trigger, err.stack || "");
  }
}
