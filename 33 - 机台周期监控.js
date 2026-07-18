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
