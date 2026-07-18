// V20260719.01 — 机台周期监控
// 入口: monitorMachineCycle（每日 8:10 定时 or 手动）
// Step 1: Workcenter(E=Y,D=6AX) → 同步至机台周期标准
// Step 2: 检查标准周期为空 → 邮件提醒 S&C 维护
// Step 3: IoT_Data 最新文件 → 按机台+班别取平均 → 写机台周期实际值
// Step 4: 平均 vs 标准 > 0.1s → 邮件报警 IDL+S&C

// ========== 常量 ==========
var _mc_WORKCENTER_SS_ID = "12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM";
var _mc_WORKCENTER_SHEET = "Workcenter";
var _mc_TARGET_SS_ID = "1cfJBxEKnNcwt1xH_tSRjKpD6Dv1JqOEzJxi2p7mZiZM";
var _mc_STANDARD_SHEET = "机台周期标准";
var _mc_ACTUAL_SHEET = "机台周期实际值";
var _mc_IOT_FOLDER_ID = "193URIPvHupVuxv0aN1kYC9ePYVkArKKs";
var _mc_IOT_FILE_PREFIX = "IoT_CT_Detail_";
var _mc_USER_SS_ID = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
var _mc_USER_SHEET = "userID";
var _mc_TARGET_MACHINE_TYPE = "6AX";
var _mc_ALARM_THRESHOLD = 0.1;
var _mc_CT_TAG_SUFFIX = ":CT";
var _mc_SENDER_NAME = "机台周期监控系统";
var _mc_TIMEZONE = "Asia/Shanghai";

var _mc_TEST_MODE = true;
var _mc_TEST_EMAIL = "kelland_zhao@colpal.com";

// ========== 主入口 ==========
/**
 * 机台周期监控主流程
 * @param {Object} e - 定时触发事件 { triggerType: "scheduled" }，手动调用可不传
 */
function monitorMachineCycle(e) {
  var trigger = e && e.triggerType ? "定时" : "手动";
  var startTime = new Date();

  try {
    console.log("=== 机台周期监控开始 ===");

    // Step 1: 机台清单同步
    var syncResult = _mc_syncMachines();
    console.log("Step1 同步: 新增" + syncResult.newCount + "台");

    // Step 2: 标准周期缺失检查
    var stdResult = _mc_checkStandards();
    console.log("Step2 标准检查: 缺失" + stdResult.missingCount + "台");

    // 合并 Step 1+2 邮件
    if (syncResult.newCount > 0 || stdResult.missingCount > 0) {
      _mc_sendMaintenanceEmail(syncResult.newMachines, stdResult.missingList, trigger);
    }

    // Step 3: 计算实际平均周期并写入
    var calcResult = _mc_calcAndWriteAverages();
    console.log("Step3 实际周期: " + calcResult.recordCount + " 条记录");

    // Step 4: 超周期报警
    if (calcResult.recordCount > 0) {
      var alarmResult = _mc_checkAndAlarm(calcResult.records, trigger);
      console.log("Step4 报警: " + alarmResult.alarmCount + " 条超标");
    }

    var duration = ((new Date()) - startTime) / 1000;
    var logMsg = "新增" + syncResult.newCount + "台/标准缺" + stdResult.missingCount +
      "/实际" + calcResult.recordCount + "条/报警" +
      (calcResult.recordCount > 0 ? (typeof alarmResult !== "undefined" ? alarmResult.alarmCount : 0) : 0);
    console.log("机台周期监控完成，耗时" + duration + "s: " + logMsg);
    try { writeLog("monitorMachineCycle", "成功", logMsg, trigger, ""); } catch (e2) {}

  } catch (err) {
    console.error("机台周期监控失败: " + err.message);
    try { writeLog("monitorMachineCycle", "失败", err.message, trigger, ""); } catch (e3) {}
  }
}

// ========== 通用工具 ==========

/** HTML 转义 */
function _mc_escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** 格式化日期 yyyy-MM-dd */
function _mc_formatDate(d) {
  return Utilities.formatDate(d, _mc_TIMEZONE, "yyyy-MM-dd");
}

/** 格式化日期时间 yyyy-MM-dd HH:mm */
function _mc_formatDateTime(d) {
  return Utilities.formatDate(d, _mc_TIMEZONE, "yyyy-MM-dd HH:mm");
}

/** 班别映射: shift值 → 中文名 */
function _mc_shiftName(shiftVal) {
  var s = String(shiftVal || "").trim();
  if (s === "1") return "夜班";
  if (s === "2") return "早班";
  if (s === "3") return "中班";
  return s;
}

// ========== Step 1: 机台清单同步 ==========

/**
 * 从 Workcenter 筛选机台，对比标准表，自动补入缺失
 * @returns {{newCount: number, newMachines: Array<{workcenter: string, machineType: string}>}}
 */
function _mc_syncMachines() {
  var ws = SpreadsheetApp.openById(_mc_WORKCENTER_SS_ID).getSheetByName(_mc_WORKCENTER_SHEET);
  if (!ws) {
    console.warn("Workcenter sheet 未找到");
    return { newCount: 0, newMachines: [] };
  }

  // 读 Workcenter
  var data = ws.getDataRange().getValues();
  // Row 1: header, data from row 2
  var workcenterMachines = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var eVal = String(row[4] || "").trim();  // E列(0-indexed:4) 是否主设备
    var dVal = String(row[3] || "").trim();  // D列(0-indexed:3) Final Machine Type
    var aVal = String(row[0] || "").trim();  // A列(0-indexed:0) Workcenter 机台号

    if (eVal === "Y" && dVal === _mc_TARGET_MACHINE_TYPE && aVal) {
      workcenterMachines[aVal] = dVal;
    }
  }
  console.log("Workcenter " + _mc_TARGET_MACHINE_TYPE + " 机台: " + Object.keys(workcenterMachines).length + " 台");

  // 读标准表现有机台
  var existing = _mc_getExistingMachines();

  // 找出缺失的
  var newMachines = [];
  Object.keys(workcenterMachines).forEach(function (mc) {
    if (!(mc in existing)) {
      newMachines.push({ workcenter: mc, machineType: workcenterMachines[mc] });
    }
  });

  // 自动补入标准表
  if (newMachines.length > 0) {
    var targetSS = SpreadsheetApp.openById(_mc_TARGET_SS_ID);
    var stdSheet = targetSS.getSheetByName(_mc_STANDARD_SHEET);
    if (stdSheet) {
      var lastRow = stdSheet.getLastRow();
      var nowStr = _mc_formatDateTime(new Date());
      newMachines.forEach(function (nm, idx) {
        var row = lastRow + 1 + idx;
        stdSheet.getRange(row, 1).setValue(nm.workcenter);
        stdSheet.getRange(row, 2).setValue(nm.machineType);
        stdSheet.getRange(row, 4).setValue(nowStr);     // 最后更新时间
        stdSheet.getRange(row, 5).setValue("自动补入");   // 备注
      });
    }
    console.log("自动补入机台: " + newMachines.length + " 台");
  }

  return { newCount: newMachines.length, newMachines: newMachines };
}

/**
 * 读取标准表中已有的机台号集合
 * @returns {Object<string, boolean>} 机台号 → true
 */
function _mc_getExistingMachines() {
  var ss = SpreadsheetApp.openById(_mc_TARGET_SS_ID);
  var sheet = ss.getSheetByName(_mc_STANDARD_SHEET);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  var existing = {};
  // Row 1: header, data from row 2
  for (var r = 1; r < data.length; r++) {
    var aVal = String(data[r][0] || "").trim();
    if (aVal) existing[aVal] = true;
  }
  return existing;
}

// ========== Step 2: 标准周期缺失检查 & 收件人 ==========

/**
 * 扫描标准表 C 列为空的行
 * @returns {{missingCount: number, missingList: Array<{workcenter: string, machineType: string}>}}
 */
function _mc_checkStandards() {
  var ss = SpreadsheetApp.openById(_mc_TARGET_SS_ID);
  var sheet = ss.getSheetByName(_mc_STANDARD_SHEET);
  if (!sheet) return { missingCount: 0, missingList: [] };

  var data = sheet.getDataRange().getValues();
  var missingList = [];
  for (var r = 1; r < data.length; r++) {
    var aVal = String(data[r][0] || "").trim();
    var cVal = String(data[r][2] || "").trim();
    var bVal = String(data[r][1] || "").trim();
    if (aVal && !cVal) {
      missingList.push({ workcenter: aVal, machineType: bVal });
    }
  }
  return { missingCount: missingList.length, missingList: missingList };
}

/**
 * 获取 S&C 收件人列表
 * userID Row2 表头: A=SAPID, B=NAME, J=GMail, O=工序, P=职位
 * 筛选: O=INJ, P 含 S&C, J 不为空
 * @returns {string[]} 邮箱列表
 */
function _mc_getSCRecipients() {
  var ws = SpreadsheetApp.openById(_mc_USER_SS_ID).getSheetByName(_mc_USER_SHEET);
  if (!ws) return [];

  var data = ws.getDataRange().getValues();
  var emails = [];
  for (var r = 2; r < data.length; r++) {
    var oVal = String(data[r][14] || "").trim();  // O列(0-indexed:14) 工序
    var pVal = String(data[r][15] || "").trim().toUpperCase();  // P列(0-indexed:15) 职位
    var jVal = String(data[r][9] || "").trim();   // J列(0-indexed:9) GMail

    if (oVal === "INJ" && pVal.indexOf("S&C") >= 0 && jVal) {
      if (emails.indexOf(jVal) < 0) emails.push(jVal);
    }
  }
  console.log("S&C 收件人: " + emails.length + " 人");
  return emails;
}

/**
 * 获取 IDL+S&C 收件人列表（报警用）
 * @returns {string[]} 邮箱列表
 */
function _mc_getAlarmRecipients() {
  var ws = SpreadsheetApp.openById(_mc_USER_SS_ID).getSheetByName(_mc_USER_SHEET);
  if (!ws) return [];

  var data = ws.getDataRange().getValues();
  var emails = [];
  for (var r = 2; r < data.length; r++) {
    var oVal = String(data[r][14] || "").trim();
    var pVal = String(data[r][15] || "").trim().toUpperCase();
    var jVal = String(data[r][9] || "").trim();

    if (oVal === "INJ" && (pVal.indexOf("IDL") >= 0 || pVal.indexOf("S&C") >= 0) && jVal) {
      if (emails.indexOf(jVal) < 0) emails.push(jVal);
    }
  }
  console.log("报警收件人(IDL+S&C): " + emails.length + " 人");
  return emails;
}

/**
 * 发送维护提醒邮件（Step 1+2 合并）
 * @param {Array} newMachines - 新增机台列表
 * @param {Array} missingList - 标准周期缺失列表
 * @param {string} trigger - "定时" or "手动"
 */
function _mc_sendMaintenanceEmail(newMachines, missingList, trigger) {
  var recipients = _mc_getSCRecipients();
  if (recipients.length === 0) {
    console.warn("无 S&C 收件人，跳过维护邮件");
    return;
  }

  var to = _mc_TEST_MODE ? _mc_TEST_EMAIL : recipients.join(",");
  var nowStr = _mc_formatDateTime(new Date());
  var todayStr = _mc_formatDate(new Date());
  var subject = "【机台周期标准维护】 " + todayStr;
  var html = _mc_buildMaintenanceEmailHtml(newMachines, missingList, nowStr);

  try {
    GmailApp.sendEmail(to, subject, "", { htmlBody: html, name: _mc_SENDER_NAME });
    console.log("维护邮件已发送: " + (newMachines.length + missingList.length) + " 项待处理");
  } catch (err) {
    console.error("维护邮件发送失败: " + err.message);
  }
}

/**
 * 构建维护提醒邮件 HTML
 */
function _mc_buildMaintenanceEmailHtml(newMachines, missingList, nowStr) {
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto">';

  // 头部 模式A 不含Logo
  html += '<div style="background:#E60012;color:white;padding:14px 28px">';
  html += '<h2 style="margin:0;font-size:20px">机台周期标准维护</h2>';
  html += '<p style="margin:4px 0 0;opacity:0.85;font-size:12px">发送时间：' + nowStr + '</p>';
  html += '</div>';

  html += '<div style="padding:20px 28px">';

  var totalItems = newMachines.length + missingList.length;
  if (totalItems === 0) {
    html += '<p style="font-size:14px;color:#27ae60">✅ 所有机台标准周期已维护，无需处理。</p>';
  } else {
    html += '<p style="font-size:14px;color:#34495e">以下机台需要维护标准周期：</p>';

    // 新增机台
    if (newMachines.length > 0) {
      html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:20px">新增机台（' + newMachines.length + ' 台）</h3>';
      html += '<table style="width:100%;font-size:13px;border-collapse:collapse">';
      html += '<tr style="background:#E60012;color:white"><th style="padding:10px;text-align:left">机台号</th><th style="padding:10px;text-align:left">机型</th></tr>';
      for (var i = 0; i < newMachines.length; i++) {
        var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
        html += '<tr style="background:' + bg + '">';
        html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _mc_escapeHtml(newMachines[i].workcenter) + '</td>';
        html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _mc_escapeHtml(newMachines[i].machineType) + '</td>';
        html += '</tr>';
      }
      html += '</table>';
    }

    // 标准缺失
    if (missingList.length > 0) {
      html += '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:20px">标准周期未维护（' + missingList.length + ' 台）</h3>';
      html += '<table style="width:100%;font-size:13px;border-collapse:collapse">';
      html += '<tr style="background:#E60012;color:white"><th style="padding:10px;text-align:left">机台号</th><th style="padding:10px;text-align:left">机型</th></tr>';
      for (var j = 0; j < missingList.length; j++) {
        var bg2 = j % 2 === 0 ? '#ffffff' : '#f8f9fa';
        html += '<tr style="background:' + bg2 + '">';
        html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _mc_escapeHtml(missingList[j].workcenter) + '</td>';
        html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _mc_escapeHtml(missingList[j].machineType) + '</td>';
        html += '</tr>';
      }
      html += '</table>';
    }

    html += '<p style="color:#e74c3c;font-size:13px;margin-top:12px">请在 ';
    html += '<a href="https://docs.google.com/spreadsheets/d/' + _mc_TARGET_SS_ID + '/edit" style="color:#E60012;text-decoration:underline">机台周期标准</a>';
    html += ' 中补录以上机台的<b>标准周期</b>。</p>';
  }

  html += '</div>';

  // 尾部
  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px;padding:0 24px 16px">此邮件由 ' + _mc_SENDER_NAME + ' 自动发送，请勿回复。</p>';
  html += '</div></body></html>';
  return html;
}

// ========== Step 3: 实际周期计算与写入 ==========

/**
 * 从 IoT_Data 文件夹获取最新日期的文件
 * 文件名格式: IoT_CT_Detail_YYYY-MM-DD
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function _mc_findLatestIoTFile() {
  var folder = DriveApp.getFolderById(_mc_IOT_FOLDER_ID);
  var files = folder.getFiles();
  var latestFileName = "";
  var latestFile = null;

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (name.indexOf(_mc_IOT_FILE_PREFIX) === 0 && name.indexOf(".gs") < 0) {
      if (name > latestFileName) {
        latestFileName = name;
        latestFile = file;
      }
    }
  }

  if (latestFile) {
    console.log("最新 IoT 文件: " + latestFileName);
    return SpreadsheetApp.open(latestFile);
  }
  console.warn("IoT_Data 文件夹中未找到文件");
  return null;
}

/**
 * 从最新 IoT 文件计算每机台每班别平均 TagValue
 * 写入机台周期实际值表（覆盖写入）
 * @returns {{recordCount: number, records: Array}}
 */
function _mc_calcAndWriteAverages() {
  var iotSS = _mc_findLatestIoTFile();
  if (!iotSS) {
    return { recordCount: 0, records: [] };
  }

  var iotSheet = iotSS.getSheetByName("IoT_CT_Detail");
  if (!iotSheet) {
    // 备选：取第一个 sheet
    var sheets = iotSS.getSheets();
    if (sheets.length === 0) return { recordCount: 0, records: [] };
    iotSheet = sheets[0];
  }

  // 取表头，确定列位置
  var header = iotSheet.getRange(1, 1, 1, iotSheet.getLastColumn()).getValues()[0];
  var colIdx = {};
  for (var h = 0; h < header.length; h++) {
    var hName = String(header[h] || "").trim();
    if (hName === "shift" || hName === "TagName" || hName === "Line" || hName === "TagValue") {
      colIdx[hName] = h;
    }
  }
  if (!("shift" in colIdx) || !("TagName" in colIdx) || !("Line" in colIdx) || !("TagValue" in colIdx)) {
    console.warn("IoT 表头不完整，缺少必要列");
    return { recordCount: 0, records: [] };
  }

  // 读数据行
  var lastRow = iotSheet.getLastRow();
  if (lastRow <= 1) return { recordCount: 0, records: [] };
  var rawData = iotSheet.getRange(2, 1, lastRow - 1, iotSheet.getLastColumn()).getValues();

  // 按 Line + shift 分组累计
  var groups = {};  // key: "Line|shift" -> { sum, count }
  for (var r = 0; r < rawData.length; r++) {
    var tagName = String(rawData[r][colIdx.TagName] || "");
    if (tagName.indexOf(_mc_CT_TAG_SUFFIX) < 0) continue;

    var line = String(rawData[r][colIdx.Line] || "").trim();
    var shift = String(rawData[r][colIdx.shift] || "").trim();
    var tagValue = parseFloat(rawData[r][colIdx.TagValue]);

    if (!line || !shift || isNaN(tagValue)) continue;

    var key = line + "|" + shift;
    if (!groups[key]) groups[key] = { sum: 0, count: 0 };
    groups[key].sum += tagValue;
    groups[key].count++;
  }

  // 获取标准周期 lookup
  var stdLookup = _mc_getStandardLookup();

  // 构建结果数组
  var records = [];
  Object.keys(groups).forEach(function (key) {
    var parts = key.split("|");
    var line = parts[0];
    var shift = parts[1];
    var avg = Math.round((groups[key].sum / groups[key].count) * 100) / 100;
    var machineType = stdLookup[line] ? stdLookup[line].machineType : "";
    var stdCycle = stdLookup[line] ? stdLookup[line].stdCycle : "";

    records.push({
      workcenter: line,
      machineType: machineType,
      shift: shift,
      avgCycle: avg,
      stdCycle: stdCycle
    });
  });

  // 按机台号+班别排序（夜早中）
  var shiftOrder = { "1": 0, "2": 1, "3": 2 };
  records.sort(function (a, b) {
    if (a.workcenter !== b.workcenter) return a.workcenter.localeCompare(b.workcenter);
    return (shiftOrder[a.shift] || 9) - (shiftOrder[b.shift] || 9);
  });

  // 写入机台周期实际值
  var targetSS = SpreadsheetApp.openById(_mc_TARGET_SS_ID);
  var actualSheet = targetSS.getSheetByName(_mc_ACTUAL_SHEET);
  if (!actualSheet) {
    console.warn("机台周期实际值 sheet 不存在");
    return { recordCount: 0, records: records };
  }

  // 清空旧数据（保留表头）
  var lastDataRow = actualSheet.getLastRow();
  if (lastDataRow > 1) {
    actualSheet.getRange(2, 1, lastDataRow - 1, actualSheet.getLastColumn()).clearContent();
  }

  // 写入新数据
  var dataDate = _mc_formatDate(new Date());
  var updateTime = _mc_formatDateTime(new Date());
  if (records.length > 0) {
    var outputRows = records.map(function (rec) {
      return [rec.workcenter, rec.machineType, _mc_shiftName(rec.shift), rec.avgCycle, rec.stdCycle, dataDate, updateTime];
    });
    actualSheet.getRange(2, 1, outputRows.length, 7).setValues(outputRows);
  }

  console.log("写入实际值: " + records.length + " 行");
  return { recordCount: records.length, records: records };
}

/**
 * 从标准表读取 lookup: 机台号 -> { machineType, stdCycle }
 * @returns {Object}
 */
function _mc_getStandardLookup() {
  var ss = SpreadsheetApp.openById(_mc_TARGET_SS_ID);
  var sheet = ss.getSheetByName(_mc_STANDARD_SHEET);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  var lookup = {};
  for (var r = 1; r < data.length; r++) {
    var aVal = String(data[r][0] || "").trim();  // 机台号
    var bVal = String(data[r][1] || "").trim();  // 机型
    var cVal = data[r][2];                         // 标准周期（数字）
    if (aVal) {
      lookup[aVal] = {
        machineType: bVal,
        stdCycle: (cVal !== "" && cVal !== null && cVal !== undefined) ? parseFloat(cVal) : null
      };
    }
  }
  return lookup;
}
