// V20260604.03 — Workcenter 数据同步（从 Line Database → Workcenter，Equipment_Number_EAM 字典匹配）
// 入口：syncWorkcenterData（每日 08:20 定时 or 手动）
// 逻辑：读取 Line Database，过滤有效 Individual Machine，通过 Equipment_Number_EAM 的 K列匹配设备编号，
//       智能转换 Final Machine Type，判断主设备标识，全量写入 Workcenter 的 A-F 列 + K列（New Formed Cell）

// ========== 数据源配置 ==========
const _ws_ID_PLAN = "11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc";
const _ws_SHEET_PLAN = "1. Line Database";

const _ws_ID_EQU = "12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM";
const _ws_SHEET_EQU = "Workcenter";
const _ws_SHEET_EQUIPMENT_NUMBER = "Equipment_Number_EAM";

// ========== 主入口 ==========
function syncWorkcenterData(e) {
  const trigger = e ? "定时" : "手动";
  try {
    console.log("开始执行 Workcenter 数据同步...");

    // 1. 读取 Line Database 数据
    console.log("正在读取 Line Database 数据...");
    const ssPlan = SpreadsheetApp.openById(_ws_ID_PLAN);
    const wsPlan = ssPlan.getSheetByName(_ws_SHEET_PLAN);
    const dataPlan = wsPlan.getDataRange().getValues();

    if (dataPlan.length <= 1) {
      try { writeLog("syncWorkcenterData", "跳过", "Line Database 表格为空或只有表头", trigger, ""); } catch (e2) {}
      return;
    }

    const objArray = _ws_getObjArray(dataPlan);
    console.log("成功读取 " + objArray.length + " 行数据");

    // 2. 读取 Equipment_Number_EAM，构建 K列→A列 映射
    console.log("正在读取 Equipment_Number_EAM 工作表数据...");
    const ssEQU = SpreadsheetApp.openById(_ws_ID_EQU);
    const wsEquipmentNumber = ssEQU.getSheetByName(_ws_SHEET_EQUIPMENT_NUMBER);
    const dataEquipmentNumber = wsEquipmentNumber.getDataRange().getValues();

    const equipmentNumberMap = {};
    if (dataEquipmentNumber.length > 1) {
      for (let i = 1; i < dataEquipmentNumber.length; i++) {
        const kColumnValue = dataEquipmentNumber[i][10]; // K列（索引10）
        const aColumnValue = dataEquipmentNumber[i][0];  // A列（索引0）

        if (kColumnValue && kColumnValue.toString().trim() !== "") {
          equipmentNumberMap[kColumnValue.toString().trim()] = aColumnValue || "";
        }
      }
      console.log("成功构建设备编号映射，共 " + Object.keys(equipmentNumberMap).length + " 条记录");
    } else {
      console.log("警告: Equipment_Number_EAM 表格为空或只有表头");
    }

    // 3. 过滤有效的 Individual Machine 记录
    const objArrayNoNull = objArray.filter(function (r) {
      return r["Individual Machine"] != "";
    });
    console.log("过滤后有效记录: " + objArrayNoNull.length + " 行");

    // 4. 构建 A-F 同步数据 + K列 New Formed Cell
    const syncData = [];       // A-F 六列
    const newFormedCellCol = []; // K列

    objArrayNoNull.forEach(function (item) {
      // D列 Final Machine Type：机器性能不为空则用机器性能，否则用 Machine Type
      let finalMachineType = "";
      if (item["机器性能"] && item["机器性能"].toString().trim() !== "") {
        finalMachineType = item["机器性能"];
      } else {
        finalMachineType = item["Machine Type"];
      }

      // 智能转换 Final Machine Type
      finalMachineType = _ws_convertFinalMachineType(finalMachineType);

      // E列 主设备标识
      const isMainEquipment = _ws_isMainEquipmentType(finalMachineType, item["Machine Type"]);

      // F列 设备编号：基于 A列 Individual Machine 在 Equipment_Number_EAM 的 K列匹配，取对应 A列值
      const workcenterValue = item["Individual Machine"] ? item["Individual Machine"].toString().trim() : "";
      let equipmentNumber = "";
      if (workcenterValue && equipmentNumberMap[workcenterValue]) {
        equipmentNumber = equipmentNumberMap[workcenterValue];
      }

      syncData.push([item["Individual Machine"], item["Machine Type"], item["机器性能"], finalMachineType, isMainEquipment, equipmentNumber]);

      // K列 New Formed Cell：机组号，来自 Line Database E列
      newFormedCellCol.push([item["New Formed Cell"] || ""]);
    });

    // 5. 执行数据同步
    console.log("开始同步数据到 Workcenter 表格...");
    const resultDataWritten = _ws_dataWritten(_ws_ID_EQU, _ws_SHEET_EQU, syncData, newFormedCellCol);

    if (resultDataWritten === true) {
      console.log("✅ 数据同步成功完成！共同步 " + syncData.length + " 条记录");
      try { writeLog("syncWorkcenterData", "成功", "同步 " + syncData.length + " 条记录到 Workcenter", trigger, ""); } catch (e2) {}
    } else {
      console.log("❌ 数据同步失败: " + resultDataWritten);
      try { writeLog("syncWorkcenterData", "失败", String(resultDataWritten), trigger, ""); } catch (e2) {}
    }

  } catch (err) {
    console.log("❌ 主函数执行错误: " + err.toString());
    console.log("错误堆栈: " + (err.stack || ""));
    try { writeLog("syncWorkcenterData", "失败", err.message, trigger, err.stack || ""); } catch (e2) {}
  }
}

// ========== 数据读取 ==========
function _ws_getObjArray(data) {
  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const rowObject = {};

    for (let j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = rowData[j];
    }

    rows.push(rowObject);
  }

  return rows;
}

// ========== 写入 Workcenter ==========
function _ws_clearSheetContent(ss, sheetName, rowCount) {
  try {
    console.log("正在清空 " + sheetName + " 表格 A-F 列及 K列内容...");
    const ws = ss.getSheetByName(sheetName);
    const lastRow = ws.getLastRow();

    if (lastRow > 1) {
      // 清空 A-F 列
      ws.getRange(2, 1, lastRow - 1, 6).clearContent();
      // 清空 K列
      ws.getRange(2, 11, lastRow - 1, 1).clearContent();
      console.log("✅ 成功清空 A-F 列及 K列 " + (lastRow - 1) + " 行数据");
    } else {
      console.log("表格只有表头，无需清空");
    }
    return true;
  } catch (e) {
    console.log("❌ 清空表格内容失败: " + e.toString());
    throw e;
  }
}

function _ws_dataWritten(id, sheetName, syncData, newFormedCellCol) {
  try {
    console.log("正在打开目标表格: " + sheetName);
    const ss = SpreadsheetApp.openById(id);
    const ws = ss.getSheetByName(sheetName);

    // 1. 清空表格内容（保留表头）
    const clearResult = _ws_clearSheetContent(ss, sheetName, syncData.length);
    if (!clearResult) {
      return "清空表格失败";
    }

    // 2. 校验数据
    const validationResult = _ws_validateData(syncData);
    if (validationResult !== true) {
      return validationResult;
    }

    // 3. 写入 A-F 列
    console.log("正在写入 " + syncData.length + " 行数据到 A-F 列...");
    if (syncData.length > 0) {
      ws.getRange(2, 1, syncData.length, 6).setValues(syncData);
      console.log("✅ A-F 列数据写入成功");
    }

    // 4. 写入 K列 New Formed Cell
    if (newFormedCellCol.length > 0) {
      ws.getRange(2, 11, newFormedCellCol.length, 1).setValues(newFormedCellCol);
      console.log("✅ K列 New Formed Cell 写入成功");
    }

    console.log("✅ 数据写入完成");
    return true;

  } catch (e) {
    console.log("❌ 数据写入失败: " + e.toString());
    return e.toString();
  }
}

// ========== 数据验证 ==========
function _ws_validateData(data) {
  if (!Array.isArray(data) || data.length === 0) {
    return "数据格式无效或为空";
  }

  for (let i = 0; i < data.length; i++) {
    if (!Array.isArray(data[i]) || data[i].length !== 6) {
      return "第 " + (i + 1) + " 行数据格式不正确，期望6个字段，实际" + data[i].length + "个字段";
    }
  }

  return true;
}

// ========== Final Machine Type 智能转换 ==========
function _ws_convertFinalMachineType(machineType) {
  if (!machineType || typeof machineType !== "string") {
    return machineType;
  }

  const typeStr = machineType.toString().trim();

  // E 开头 → ENG
  if (typeStr.startsWith("E")) {
    return "ENG";
  }

  // F 开头 + F数字(数字) 格式（排除 FT400）→ FCS
  if (typeStr.startsWith("F")) {
    if (typeStr === "FT400") {
      return typeStr; // 保持原值
    }

    if (/^F\d+\(\d+\)$/.test(typeStr)) {
      return "FCS";
    }
  }

  return typeStr;
}

// ========== 主设备判断 ==========
function _ws_isMainEquipmentType(finalMachineType, machineType) {
  // B列 Machine Type 以 HT 开头 → Y
  if (machineType && typeof machineType === "string") {
    const machineTypeStr = machineType.toString().trim();
    if (machineTypeStr.startsWith("HT")) {
      return "Y";
    }
  }

  // D列 Final Machine Type 属于主设备类型列表 → Y
  if (finalMachineType && typeof finalMachineType === "string") {
    const typeStr = finalMachineType.toString().trim();
    const mainEquipmentTypes = ["HS", "DP", "ENG", "FCS", "H Auto"];

    if (mainEquipmentTypes.includes(typeStr)) {
      return "Y";
    }
  }

  return "N";
}
