// 注塑工艺周检查 2.0
// 定时触发入口: runWeeklyInspection()

// ========== 数据源配置 ==========
const INSPECTION_CONFIG = {
  MACHINE_SOURCE: {
    ID: "1BeoCokGiWAdkfFTSVOkxNr4Gr9O6FlnwTvnOmLYNY_U",
    SHEET_NAME: "Database",
  },
  INSPECTION_SOURCE: {
    ID: "18hoqx_pnoRrjqijiDCOWz_hnyu3EUSWS-_sWKAwh868",
    SHEET_NAME: "Injection",
  },
  MASTER_DATA: {
    ID: "1afvNifotG_Ik36NQ7aptPjKT4ebAyeSBCc4hJ9WL7v4",
    SHEET_NAME: "Master Data",
  },
};

// ========== 定时触发入口 ==========
function runWeeklyInspection() {
  _weeklyInspectionMain("定时触发");
}

// ========== 手动触发入口 ==========
function runWeeklyInspectionManual() {
  _weeklyInspectionMain("手动");
}

// ========== 主流程 ==========
function _weeklyInspectionMain(trigger) {
  try {
    console.log("开始执行注塑工艺周检查2.0...");

    const currentWeek = _wiGetCurrentWeek();
    console.log(`当前周次: ${currentWeek}`);

    const machineData = _wiGetMachineData();
    console.log(`获取到机台数据: ${machineData.length} 条`);

    const inspectionData = _wiGetInspectionData();
    console.log(`获取到检查记录: ${inspectionData.length} 条`);

    const masterData = _wiGetMasterData();
    console.log(`获取到主数据: ${masterData.length} 条`);

    // 保留历史数据，重建当前周次
    const historicalData = masterData.filter(r => r["周次"] !== currentWeek);
    console.log(`保留历史数据: ${historicalData.length} 条`);

    const currentWeekBaseData = machineData.map(machine => ({
      周次: currentWeek,
      Workcenter: machine["Workcenter"],
      责任人: machine["责任人"],
      备份责任人: machine["备份责任人"],
      实际检查人: "",
      工号: "",
      车间: "",
      物料: "",
      时间戳: "",
      工艺参数: "",
      通过备注: "",
      豁免备注: "",
      豁免原因: "",
      "日期&班次": "",
    }));
    console.log(`创建当前周次基础数据: ${currentWeekBaseData.length} 条`);

    const updatedCurrentWeekData = _wiMatchInspectionRecords(currentWeekBaseData, inspectionData, currentWeek);
    console.log(`检查记录匹配完成，当前周次数据: ${updatedCurrentWeekData.length} 条`);

    const finalMasterData = [...historicalData, ...updatedCurrentWeekData];
    console.log(`最终主数据: ${finalMasterData.length} 条`);

    // 补录上周漏匹配的空记录
    const prevWeek = _wiGetPreviousWeek(currentWeek);
    const backfilledFinalData = finalMasterData.map(record => {
      if (record["周次"] === prevWeek && !record["实际检查人"]) {
        const matchingInspection = inspectionData.find(insp =>
          insp["机台"] === record["Workcenter"] &&
          _wiGetWeekFromTimestamp(insp["时间戳"]) === prevWeek
        );
        if (matchingInspection) {
          console.log(`🔄 补录上周记录: ${record["Workcenter"]} (${prevWeek})`);
          return {
            ...record,
            实际检查人: matchingInspection["姓名"],
            工号: matchingInspection["工号"],
            车间: matchingInspection["车间"],
            物料: matchingInspection["物料"],
            时间戳: matchingInspection["时间戳"],
            工艺参数: matchingInspection["工艺参数"],
            通过备注: matchingInspection["通过备注"],
            豁免备注: matchingInspection["豁免备注"],
            豁免原因: matchingInspection["豁免原因"],
            "日期&班次": matchingInspection["日期&班次"],
          };
        }
      }
      return record;
    });
    console.log(`上周补录完成 (${prevWeek})`);

    _wiWriteMasterData(backfilledFinalData);
    console.log("注塑工艺周检查2.0执行完成！");
    try { writeLog("runWeeklyInspection", "成功", `周次 ${currentWeek}，写入 ${backfilledFinalData.length} 条记录`, trigger, ""); } catch (err) {}
  } catch (error) {
    console.error("执行过程中发生错误:", error);
    try { writeLog("runWeeklyInspection", "失败", error.message, trigger, error.stack || ""); } catch (err) {}
    throw error;
  }
}

// ========== 数据读取 ==========
function _wiGetMachineData() {
  try {
    const sheet = SpreadsheetApp.openById(INSPECTION_CONFIG.MACHINE_SOURCE.ID)
      .getSheetByName(INSPECTION_CONFIG.MACHINE_SOURCE.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const machineData = [];
    for (let i = 1; i < data.length; i++) {
      const rec = {};
      headers.forEach((h, idx) => { rec[h] = data[i][idx]; });
      if (!rec["无需检查Y/N"] || rec["无需检查Y/N"] !== "Y") {
        machineData.push(rec);
      }
    }
    return machineData;
  } catch (e) {
    console.error("获取机台数据失败:", e);
    throw e;
  }
}

function _wiGetInspectionData() {
  try {
    const sheet = SpreadsheetApp.openById(INSPECTION_CONFIG.INSPECTION_SOURCE.ID)
      .getSheetByName(INSPECTION_CONFIG.INSPECTION_SOURCE.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const rec = {};
      headers.forEach((h, idx) => { rec[h] = data[i][idx]; });
      result.push(rec);
    }
    return result;
  } catch (e) {
    console.error("获取检查记录失败:", e);
    throw e;
  }
}

function _wiGetMasterData() {
  try {
    const sheet = SpreadsheetApp.openById(INSPECTION_CONFIG.MASTER_DATA.ID)
      .getSheetByName(INSPECTION_CONFIG.MASTER_DATA.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const rec = {};
      headers.forEach((h, idx) => { rec[h] = data[i][idx]; });
      result.push(rec);
    }
    return result;
  } catch (e) {
    console.error("获取主数据失败:", e);
    throw e;
  }
}

// ========== 匹配逻辑 ==========
function _wiMatchInspectionRecords(masterData, inspectionData, currentWeek) {
  const finalMasterData = [];
  const unmatchedInspections = [];

  console.log(`开始匹配检查记录，当前周次: ${currentWeek}`);
  console.log(`检查记录总数: ${inspectionData.length}`);

  if (inspectionData.length > 0) {
    console.log("检查记录字段结构:", Object.keys(inspectionData[0]));
  }

  for (const masterRecord of masterData) {
    if (masterRecord["周次"] === currentWeek) {
      const matchingInspection = inspectionData.find(inspection => {
        const machineMatch = inspection["机台"] === masterRecord["Workcenter"];
        const inspectionWeek = _wiGetWeekFromTimestamp(inspection["时间戳"]);
        return machineMatch && inspectionWeek === currentWeek;
      });

      if (matchingInspection) {
        const inspectionWeek = _wiGetWeekFromTimestamp(matchingInspection["时间戳"]);
        finalMasterData.push({
          ...masterRecord,
          实际检查人: matchingInspection["姓名"],
          工号: matchingInspection["工号"],
          车间: matchingInspection["车间"],
          物料: matchingInspection["物料"],
          时间戳: matchingInspection["时间戳"],
          工艺参数: matchingInspection["工艺参数"],
          通过备注: matchingInspection["通过备注"],
          豁免备注: matchingInspection["豁免备注"],
          豁免原因: matchingInspection["豁免原因"],
          "日期&班次": matchingInspection["日期&班次"],
        });
        console.log(`✅ 匹配: ${masterRecord["Workcenter"]} - ${matchingInspection["姓名"]} (${inspectionWeek})`);
      } else {
        finalMasterData.push(masterRecord);
        console.log(`❌ 未匹配: ${masterRecord["Workcenter"]}`);
      }
    } else {
      finalMasterData.push(masterRecord);
    }
  }

  // 检查未匹配的检查记录
  for (const inspection of inspectionData) {
    const inspectionWeek = _wiGetWeekFromTimestamp(inspection["时间戳"]);
    const machineExists = masterData.some(r =>
      r["Workcenter"] === inspection["机台"] && r["周次"] === currentWeek
    );
    if (inspectionWeek === currentWeek && machineExists) {
      const alreadyMatched = finalMasterData.some(r =>
        r["Workcenter"] === inspection["机台"] && r["实际检查人"] === inspection["姓名"]
      );
      if (!alreadyMatched) {
        unmatchedInspections.push({
          机台: inspection["机台"],
          姓名: inspection["姓名"],
          时间戳: inspection["时间戳"],
          转换周次: inspectionWeek,
        });
      }
    }
  }

  if (unmatchedInspections.length > 0) {
    console.log(`⚠️  发现 ${unmatchedInspections.length} 条未匹配的检查记录:`);
    unmatchedInspections.forEach(i => {
      console.log(`   - 机台: ${i.机台}, 姓名: ${i.姓名}, 时间戳: ${i.时间戳}, 转换周次: ${i.转换周次}`);
    });
  }

  return finalMasterData;
}

// ========== 写入主数据 ==========
function _wiWriteMasterData(masterData) {
  try {
    const sheet = SpreadsheetApp.openById(INSPECTION_CONFIG.MASTER_DATA.ID)
      .getSheetByName(INSPECTION_CONFIG.MASTER_DATA.SHEET_NAME);

    const headers = [
      "周次", "Workcenter", "责任人", "备份责任人",
      "实际检查人", "工号", "车间", "物料",
      "时间戳", "工艺参数", "通过备注", "豁免备注", "豁免原因", "日期&班次",
    ];

    const existingData = sheet.getDataRange().getValues();
    if (existingData.length === 0 || existingData[0].length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clear();
    }

    if (masterData.length > 0) {
      const dataRows = masterData.map(record =>
        headers.map(header => {
          const val = record[header];
          if (val instanceof Date) {
            return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
          return val || "";
        })
      );
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }

    console.log(`成功写入 ${masterData.length} 条主数据记录`);
  } catch (e) {
    console.error("写入主数据失败:", e);
    throw e;
  }
}

// ========== 工具函数 ==========
function _wiGetCurrentWeek() {
  const now = new Date();
  return `${now.getFullYear()}-W${_wiGetWeekNumber(now).toString().padStart(2, "0")}`;
}

function _wiGetPreviousWeek(weekStr) {
  const year = parseInt(weekStr.split("-W")[0]);
  const week = parseInt(weekStr.split("-W")[1]);
  if (week === 1) {
    const lastWeek = _wiGetWeekNumber(new Date(year - 1, 11, 28));
    return `${year - 1}-W${lastWeek.toString().padStart(2, "0")}`;
  }
  return `${year}-W${(week - 1).toString().padStart(2, "0")}`;
}

function _wiGetWeekNumber(date) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  const thursday = new Date(target.valueOf());
  thursday.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(thursday.getFullYear(), 0, 1);
  const dayNr2 = (firstThursday.getDay() + 6) % 7;
  const firstThursdayDate = new Date(firstThursday.valueOf());
  firstThursdayDate.setDate(firstThursday.getDate() - dayNr2 + 3);
  return Math.floor((thursday - firstThursdayDate) / (7 * 24 * 60 * 60 * 1000)) + 1;
}

function _wiGetWeekFromTimestamp(timestamp) {
  if (!timestamp) return "";
  try {
    let date;
    if (typeof timestamp === "string") {
      date = new Date(timestamp.includes(" ") ? timestamp.replace(" ", "T") : timestamp);
    } else {
      date = new Date(timestamp);
    }
    if (isNaN(date.getTime())) {
      console.error("无效的时间戳:", timestamp);
      return "";
    }
    return `${date.getFullYear()}-W${_wiGetWeekNumber(date).toString().padStart(2, "0")}`;
  } catch (e) {
    console.error("解析时间戳失败:", timestamp, e);
    return "";
  }
}
