// V20260814.01 — E&E 考勤数据同步 + 班次顺序自动更新（修复日期单元格误判为休息；同步窗口扩展为昨天~今天+5）
// 入口：syncAttendanceFromEE（每日定时）、updateShiftSchedule（每周六定时）
// 数据源：E&E电子考勤记录 (1dMON_DEcAUH9xRsfOkEF37fIN7DuyVHfNwOoUyd-V-0)
// 映射表：userID (workshop) + 班次顺序 (shift)
// 目标表：EDS 任务安排主数据 → AttendanceSync sheet + 班次顺序 sheet

// ========== 常量 ==========
const _ee_SOURCE_SS_ID = "1dMON_DEcAUH9xRsfOkEF37fIN7DuyVHfNwOoUyd-V-0";
const _ee_TASK_SS_ID = "1UBg1Ake18cFp6gj0jKRX1Y9GJ0VL1pY5aXK-UoCeAY0";
const _ee_USER_SS_ID = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const _ee_ATTENDANCE_SHEET = "AttendanceSync";

// EDS 三工序代码
const _ee_TARGET_PROCESSES = ['INJ', 'TF', 'PK'];

// 同步天数范围：今天 + 未来 N 天（含今天）
const _ee_SYNC_DAYS = 5;
// 回看天数：含昨天，保证前一天补录/修正的出勤数据能被同步
const _ee_SYNC_LOOKBACK_DAYS = 1;

// ========== 主入口 ==========
/**
 * 从 E&E 考勤表同步 昨天~今天+5天 的出勤数据到 AttendanceSync
 * @param {Object} e - 定时触发传入的事件对象 { triggerType: "scheduled" }，手动调用可不传
 */
function syncAttendanceFromEE(e) {
  const trigger = e && e.triggerType ? "定时" : "手动";

  // 防止并发执行（如上次还没跑完下一次就触发了）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.warn("考勤同步跳过：上次执行尚未完成");
    return { success: false, error: "locked" };
  }

  const startTime = new Date();

  try {
    // 生成同步日期列表：今天 + 未来 N 天
    var dates = _ee_getTargetDates();
    console.log("考勤同步日期范围: " + dates.join(", "));

    // 1. 构建映射表（全局，只读一次）
    var sapWorkshopMap = _ee_buildSapWorkshopMap();
    console.log("车间映射表大小: " + Object.keys(sapWorkshopMap).length);

    // 预读各日期的班次映射
    var dateShiftMaps = {};
    dates.forEach(function (d) {
      dateShiftMaps[d] = _ee_buildDateShiftMap(d);
    });

    // 2. 逐日读取 E&E 数据并组装行（缓存 sheet 引用避免重复查找）
    var allRows = [];
    var sheetCache = {};   // monthKey → sheet | null
    var skippedMonths = {}; // monthKey → [dates]

    dates.forEach(function (targetDate) {
      var monthKey = targetDate.substring(0, 7); // "2026-07"
      if (!(monthKey in sheetCache)) {
        sheetCache[monthKey] = _ee_findMonthSheet(targetDate) || null;
      }
      var eeSheet = sheetCache[monthKey];
      if (!eeSheet) {
        if (!skippedMonths[monthKey]) skippedMonths[monthKey] = [];
        skippedMonths[monthKey].push(targetDate);
        return;
      }

      var colIndex = _ee_findDayColumn(eeSheet, targetDate);
      if (colIndex < 0) {
        console.warn("跳过 " + targetDate + "：未找到对应列");
        return;
      }

      var eeData = _ee_readDailyData(eeSheet, colIndex);
      var dateShiftMap = dateShiftMaps[targetDate] || {};
      var seen = {};

      eeData.forEach(function (person) {
        if (_ee_TARGET_PROCESSES.indexOf(person.process) === -1) return;

        var sapID = person.sapID;
        var key = targetDate + "|" + sapID;
        if (seen[key]) return;
        seen[key] = true;

        // E&E SAP ID 为8位(如32867920)，userID 为5位(如67920)，取后5位匹配
        var shortSapID = sapID.length > 5 ? sapID.slice(-5) : sapID;
        var workshop = sapWorkshopMap[shortSapID] || "";
        var shift = "";
        if (person.team === "D") {
          shift = "早班";
        } else {
          shift = dateShiftMap[person.team] || "";
        }
        var status = person.status || "休息";

        allRows.push([
          targetDate,
          shortSapID,
          person.name,
          person.process,
          person.team,
          workshop,
          shift,
          person.hours,
          status,
          "E&E考勤记录",
          Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm")
        ]);
      });

      console.log(targetDate + ": " + eeData.length + " 人 → 三工序 " + Object.keys(seen).length + " 条");
    });

    // 汇总未找到月度 sheet 的日期
    var skippedMonthKeys = Object.keys(skippedMonths);
    if (skippedMonthKeys.length > 0) {
      skippedMonthKeys.forEach(function (mk) {
        var skippedDates = skippedMonths[mk];
        console.warn("未找到 " + mk + " 月度 sheet，跳过以下日期: " + skippedDates.join(", "));
      });
    }

    // 3. 覆盖写入（原地覆写目标日期数据，保留历史）
    _ee_writeToAttendanceSync(dates, allRows);

    var duration = ((new Date()) - startTime) / 1000;
    var logDetail = dates[0] + "~" + dates[dates.length - 1] + " 同步 " + allRows.length + " 条，耗时 " + duration + "s";
    console.log(logDetail);

    try {
      writeLog("syncAttendanceFromEE", "成功", logDetail, trigger, "");
    } catch (e2) {
      console.warn("writeLog 失败: " + e2.message);
    }

    return { success: true, dates: dates, count: allRows.length, duration: duration };

  } catch (err) {
    console.error("考勤同步失败: " + err.message);
    try {
      writeLog("syncAttendanceFromEE", "失败", err.message, trigger, "");
    } catch (e2) {
      console.warn("writeLog 失败: " + e2.message);
    }
    return { success: false, error: err.message };
  }
}

// ========== 日期工具 ==========

/** 获取同步日期列表：昨天 ~ 今天+未来 N 天 */
function _ee_getTargetDates() {
  var dates = [];
  var now = new Date();
  for (var i = -_ee_SYNC_LOOKBACK_DAYS; i <= _ee_SYNC_DAYS; i++) {
    var d = new Date(now.getTime() + i * 24 * 60 * 60 * 1000);
    dates.push(Utilities.formatDate(d, "Asia/Shanghai", "yyyy-MM-dd"));
  }
  return dates;
}

/** 获取当天日期（用于测试等单日场景） */
function _ee_getTargetDate() {
  return Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd");
}

// ========== E&E Sheet 定位 ==========

/**
 * 根据日期定位 E&E 考勤表中的月度 sheet
 * 尝试多种命名模式：YYYY.MM → YYYY.MM月 → YYYY.M → YYYY.M月
 */
function _ee_findMonthSheet(targetDate) {
  const parts = targetDate.split("-");
  const year = parts[0];
  const month = parts[1];         // "07"
  const monthNoPad = String(parseInt(month, 10)); // "7"

  const ss = SpreadsheetApp.openById(_ee_SOURCE_SS_ID);

  // 尝试模式（按优先级）
  const patterns = [
    year + "." + month,            // 2026.07
    year + "." + month + "月",     // 2026.07月
    year + "." + monthNoPad,       // 2026.7
    year + "." + monthNoPad + "月" // 2026.7月
  ];

  for (var i = 0; i < patterns.length; i++) {
    var sheet = ss.getSheetByName(patterns[i]);
    if (sheet) {
      console.log("匹配 sheet 命名: " + patterns[i]);
      return sheet;
    }
  }

  // 模糊搜索：遍历所有 sheet，匹配年份+月份
  var allSheets = ss.getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    var name = allSheets[j].getSheetName();
    // 匹配 "2026.07" 或 "2026.07月" 或 "2026.7" 或 "2026.7月"
    if (name.indexOf(year + "." + month) === 0 || name.indexOf(year + "." + monthNoPad) === 0) {
      console.log("模糊匹配 sheet: " + name);
      return allSheets[j];
    }
  }

  return null;
}

/**
 * 在 E&E sheet 中找到目标日期对应的列索引（0-indexed）
 * Row 4 是列头行，包含 "1"~"31" 等日期数字
 */
function _ee_findDayColumn(eeSheet, targetDate) {
  var targetDay = parseInt(targetDate.split("-")[2], 10); // 12
  var targetDayStr = String(targetDay);

  var headerRow = eeSheet.getRange(4, 1, 1, eeSheet.getLastColumn()).getValues()[0];

  // 先找当月列（列 G=6 到 AK=36，对应 day 1-31）
  for (var c = 6; c <= 36 && c < headerRow.length; c++) {
    if (String(headerRow[c] || "").trim() === targetDayStr) {
      return c;
    }
  }

  // 目标日<=31但没找到，可能是跨月（如7月31日后是8月1-7日）
  // AL=37 开始是次月溢出
  for (var c2 = 37; c2 < headerRow.length; c2++) {
    if (String(headerRow[c2] || "").trim() === targetDayStr) {
      return c2;
    }
  }

  return -1;
}

// ========== E&E 数据读取 ==========

/**
 * 读取 E&E sheet 指定列的人员出勤数据
 * @returns {Array<{sapID, name, process, team, hours}>}
 */
function _ee_readDailyData(eeSheet, colIndex) {
  var lastRow = eeSheet.getLastRow();
  if (lastRow < 5) return [];

  // 读取全部数据列：B(工号), C(姓名), D(工序), E(班别), 目标日列
  var allData = eeSheet.getRange(5, 2, lastRow - 4, 4).getValues();
  var dayValues = eeSheet.getRange(5, colIndex + 1, lastRow - 4, 1).getValues();

  var result = [];
  for (var i = 0; i < allData.length; i++) {
    var sapID = String(allData[i][0] || "").trim();
    var name = String(allData[i][1] || "").trim();
    var process = String(allData[i][2] || "").trim();
    var team = String(allData[i][3] || "").trim();

    if (!name) continue;

    var cellValue = _ee_cellValueToText(dayValues[i][0]);
    var parsed = _ee_parseCellValue(cellValue);

    result.push({
      sapID: sapID,
      name: name,
      process: process,
      team: team,
      hours: parsed.hours,
      status: parsed.status,
      cellRaw: cellValue
    });
  }

  return result;
}

/**
 * 单元格值 → 可解析文本
 * Sheets 会把 "11/3" 这类录入自动识别为日期（11月3日），getValues() 返回 Date 对象，
 * String(Date) 是长日期字符串，parseFloat 得到 NaN → 被误判为"休息"。
 * 这里把日期单元格还原为录入形态 "月/日"，例如 2026-11-03 → "11/3"。
 */
function _ee_cellValueToText(v) {
  if (v instanceof Date) {
    return (v.getMonth() + 1) + "/" + v.getDate();
  }
  return String(v || "").trim();
}

// E&E 休假/出勤代码映射（来源：E&E Settings sheet F-G列）
var _ee_CODE_MAP = {
  'Y01': '年休假',
  'Y09': '三八假',
  'Y04': '公司假',
  'N01': '病假',
  'Y07': '恩恤假'
};

/**
 * 解析出勤单元格值 → { hours, status }
 * "7.5/2" → {hours:7.5, status:"在岗"}, "N01" → {hours:0, status:"病假"}
 * "" → {hours:0, status:"休息"}, "4/2\\nY01" → {hours:4, status:"在岗"}
 */
function _ee_parseCellValue(cellValue) {
  if (!cellValue) return { hours: 0, status: "休息" };

  // 多行单元格（如 "4/2\nY01"），取第一行解析工时
  var firstLine = cellValue.split("\n")[0].trim();

  // 检查第二行是否有休假代码（如 "4/2\nY01"）
  var lines = cellValue.split("\n");
  var codeLine = lines.length > 1 ? lines[lines.length - 1].trim() : "";

  if (!firstLine) return { hours: 0, status: "休息" };

  // 纯字母代码（如 N01, Y01, V02）— 整日休假
  if (/^[A-Z]\d+/.test(firstLine)) {
    var desc = _ee_CODE_MAP[firstLine] || firstLine;
    return { hours: 0, status: desc };
  }

  // 数字格式（含 "/" 如 "7.5/2"、"-" 如 "11/3-1"）
  var slashIdx = firstLine.indexOf("/");
  var numStr = slashIdx > 0 ? firstLine.substring(0, slashIdx) : firstLine;
  // 去掉可能的后缀如 "-1"
  var dashIdx = numStr.indexOf("-");
  if (dashIdx > 0) numStr = numStr.substring(0, dashIdx);

  var num = parseFloat(numStr);
  if (!isNaN(num) && num > 0) {
    // 如果有第二行代码，附加说明（如 "在岗(Y01年休假)"）
    if (codeLine && _ee_CODE_MAP[codeLine]) {
      return { hours: num, status: "在岗" };
    }
    return { hours: num, status: "在岗" };
  }

  return { hours: 0, status: "休息" };
}

// ========== 映射表构建 ==========

/**
 * 从 userID 表构建 SAP ID → Workshop 映射
 * userID 表: A列=SAPID(0), N列=车间(13)
 */
function _ee_buildSapWorkshopMap() {
  var ws = SpreadsheetApp.openById(_ee_USER_SS_ID).getSheetByName("userID");
  if (!ws) return {};

  var values = ws.getDataRange().getValues();
  var map = {};
  for (var i = 2; i < values.length; i++) {
    var sapID = String(values[i][0] || "").trim();
    var workshop = String(values[i][13] || "").trim();
    if (sapID && workshop) {
      map[sapID] = workshop;
    }
  }
  return map;
}

/**
 * 从班次顺序 sheet 构建 组别(A/B/C) → 班次 映射
 * D组固定返回空（外部处理为早班）
 * 班次顺序: A列=日期, B列=班次, C列=组别
 */
function _ee_buildDateShiftMap(targetDate) {
  var ss = SpreadsheetApp.openById(_ee_TASK_SS_ID);
  var ws = ss.getSheetByName("班次顺序");
  if (!ws) return {};

  var data = ws.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var rowDate = _ee_normalizeDate(data[i][0]);
    var shift = String(data[i][1] || "").trim();
    var team = String(data[i][2] || "").trim();

    if (rowDate === targetDate && team && shift) {
      map[team] = shift;
    }
  }

  return map;
}

/**
 * 日期标准化 → yyyy-MM-dd
 */
function _ee_normalizeDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Shanghai", "yyyy-MM-dd");
  }
  return String(val).trim();
}

// ========== AttendanceSync 写入 ==========

/**
 * 覆盖写入：先原地覆写新数据，再清除末尾冗余行
 * 避免 clearContents 裸奔导致中途崩溃数据全丢
 * @param {string[]} dates - 要覆盖的日期列表
 * @param {Array[]} rows  - 新数据行
 */
function _ee_writeToAttendanceSync(dates, rows) {
  var ss = SpreadsheetApp.openById(_ee_TASK_SS_ID);
  var ws = ss.getSheetByName(_ee_ATTENDANCE_SHEET);
  if (!ws) {
    throw new Error("未找到 " + _ee_ATTENDANCE_SHEET + " sheet");
  }

  var dateSet = {};
  dates.forEach(function (d) { dateSet[d] = true; });

  var lastRow = ws.getLastRow();
  var headerRow = 1;

  if (lastRow > headerRow) {
    // 读取全部现有数据
    var allData = ws.getRange(headerRow, 1, lastRow, 11).getValues(); // A~K 共11列
    var keepRows = [allData[0]]; // 保留表头
    var deletedCount = 0;

    for (var r = 1; r < allData.length; r++) {
      var rowDate = _ee_normalizeDate(allData[r][0]);
      if (dateSet[rowDate]) {
        deletedCount++;
      } else {
        keepRows.push(allData[r]);
      }
    }

    var finalRows = keepRows.concat(rows);
    var newTotal = finalRows.length; // 含表头

    if (newTotal > 1) {
      // 原地覆写：setValues 是原子的，不会留空窗口
      ws.getRange(1, 1, newTotal, 11).setValues(finalRows);
      // 新数据比旧数据短 → 清除末尾冗余行
      if (newTotal < lastRow) {
        ws.getRange(newTotal + 1, 1, lastRow - newTotal, 11).clearContent();
      }
      console.log("覆盖旧数据: " + deletedCount + " 行，保留: " + (keepRows.length - 1) + " 行，写入新数据: " + rows.length + " 行");
    }
  } else {
    // 空表，直接写入（含表头）
    if (rows.length > 0) {
      ws.getRange(1, 1, 1, rows[0].length).setValues([["日期<br>Date", "SAP ID", "姓名<br>Name", "工序<br>Process", "班别<br>Team", "车间<br>Workshop", "班次<br>Shift", "工时<br>Hours", "出勤状态<br>Status", "数据来源<br>Source", "同步时间<br>Sync Time"]]);
      ws.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      console.log("初始化写入: 表头 + " + rows.length + " 行");
    }
  }
}

// ========== 班次顺序自动更新 ==========

/** 轮换表：早→夜→中→早（每周一更换） */
var _ee_SHIFT_ROTATION = {
  '早班': '夜班',
  '夜班': '中班',
  '中班': '早班'
};

/** 参与轮换的组别（D组固定早班，不在sheet中体现） */
var _ee_ROTATION_GROUPS = ['A', 'B', 'C'];

/**
 * 自动更新班次顺序 sheet，生成下3周排班
 * - 轮换规则：每周一更换，早→夜→中→早，D组永远早班
 * - 预定每周六自动执行（通过定时设置配置），也可手动触发
 * - 生成范围：下周一 ~ 下下下周日（3周 × 7天 × 3组 = 63行）
 * @param {Object} e - 定时触发传入的事件对象 { triggerType: "scheduled" }
 */
function updateShiftSchedule(e) {
  var trigger = e && e.triggerType ? "定时" : "手动";
  var startTime = new Date();

  try {
    var ss = SpreadsheetApp.openById(_ee_TASK_SS_ID);
    var ws = ss.getSheetByName("班次顺序");
    if (!ws) throw new Error("未找到班次顺序 sheet");

    // 1. 读取现有数据
    var data = ws.getDataRange().getValues();

    // 2. 找到最近一个周一三组齐全的排班状态
    var latestState = _findLatestShiftState(data);
    if (!latestState) throw new Error("无法确定当前排班状态，sheet数据不足（需至少一个完整周）");

    console.log("基准周一: " + Utilities.formatDate(latestState.monday, "Asia/Shanghai", "yyyy-MM-dd") +
                " | A=" + latestState.state.A + " B=" + latestState.state.B + " C=" + latestState.state.C);

    // 3. 计算目标起始日期（下周一）
    var nextMonday = _getNextMonday();
    console.log("目标起始周一: " + Utilities.formatDate(nextMonday, "Asia/Shanghai", "yyyy-MM-dd"));

    // 4. 从基准周一到目标周一，逐周应用轮换
    var weekDiff = Math.round((nextMonday.getTime() - latestState.monday.getTime()) / (7 * 24 * 60 * 60 * 1000));
    var state = { A: latestState.state.A, B: latestState.state.B, C: latestState.state.C };
    for (var r = 0; r < weekDiff; r++) {
      state = {
        A: _ee_SHIFT_ROTATION[state.A],
        B: _ee_SHIFT_ROTATION[state.B],
        C: _ee_SHIFT_ROTATION[state.C]
      };
    }
    if (weekDiff > 0) {
      console.log("经" + weekDiff + "周轮换后: A=" + state.A + " B=" + state.B + " C=" + state.C);
    }

    // 5. 生成3周排班
    var newRows = _generateShiftRows(state, nextMonday, 3);

    // 6. 收集已有日期+组别，去重避免重复写入
    var existingSet = {};
    for (var i = 1; i < data.length; i++) {
      var ed = _ee_normalizeDate(data[i][0]);
      var eg = String(data[i][2] || "").trim();
      if (ed && eg) existingSet[ed + "|" + eg] = true;
    }

    var rowsToAdd = newRows.filter(function(row) {
      return !existingSet[row[0] + "|" + row[2]];
    });

    // 7. 追加写入
    if (rowsToAdd.length > 0) {
      ws.getRange(data.length + 1, 1, rowsToAdd.length, 3).setValues(rowsToAdd);
      var expectedTotal = 21 * 3; // 3组 × 7天 × 3周
      console.log("已添加 " + rowsToAdd.length + " 行（目标" + expectedTotal + "行，跳过重复" + (newRows.length - rowsToAdd.length) + "行）");
    } else {
      console.log("无需更新，目标日期已有完整排班数据");
    }

    var duration = ((new Date()) - startTime) / 1000;
    try {
      writeLog("updateShiftSchedule", "成功",
               "添加" + rowsToAdd.length + "行，起始周一" + Utilities.formatDate(nextMonday, "Asia/Shanghai", "yyyy-MM-dd") + "，耗时" + duration + "s",
               trigger, "");
    } catch (e2) {}

    return {
      success: true,
      count: rowsToAdd.length,
      startDate: Utilities.formatDate(nextMonday, "Asia/Shanghai", "yyyy-MM-dd"),
      duration: duration
    };

  } catch (err) {
    console.error("更新班次顺序失败: " + err.message);
    try { writeLog("updateShiftSchedule", "失败", err.message, trigger, ""); } catch (e2) {}
    return { success: false, error: err.message };
  }
}

/**
 * 从现有数据中找到最近一个周一的三组排班状态
 * @param {Array<Array>} data - 含表头的全部行
 * @returns {{ monday: Date, state: {A: string, B: string, C: string} } | null}
 */
function _findLatestShiftState(data) {
  // dateStr → { group → shift }
  var dateShiftMap = {};
  for (var i = 1; i < data.length; i++) {
    var dateStr = _ee_normalizeDate(data[i][0]);
    var shift = String(data[i][1] || "").trim();
    var group = String(data[i][2] || "").trim();
    if (!dateStr || !shift || !group) continue;
    if (!dateShiftMap[dateStr]) dateShiftMap[dateStr] = {};
    dateShiftMap[dateStr][group] = shift;
  }

  // 按日期倒序找最新一个三组齐全的周一
  var dates = Object.keys(dateShiftMap).sort(); // 升序
  for (var j = dates.length - 1; j >= 0; j--) {
    var ds = dates[j];
    var parts = ds.split("-");
    var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    if (d.getDay() !== 1) continue; // 非周一跳过

    var gm = dateShiftMap[ds];
    if (gm['A'] && gm['B'] && gm['C']) {
      return {
        monday: d,
        state: { A: gm['A'], B: gm['B'], C: gm['C'] }
      };
    }
  }

  return null;
}

/** 获取下周一日期 */
function _getNextMonday() {
  var now = new Date();
  var dayOfWeek = now.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  // 周日(0)→+1天, 周一(1)→+7天, 周二-六(2-6)→+(8-dayOfWeek)天
  var daysUntilMonday = dayOfWeek === 0 ? 1 : (8 - dayOfWeek);
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysUntilMonday);
}

/**
 * 从起始周一开始生成 N 周的排班行
 * @param {{A: string, B: string, C: string}} startState - 起始周一的排班
 * @param {Date} startMonday - 起始周一日期
 * @param {number} weeks - 生成周数
 * @returns {Array<Array<string>>} [[日期, 班别, 组别], ...]
 */
function _generateShiftRows(startState, startMonday, weeks) {
  var rows = [];
  var state = { A: startState.A, B: startState.B, C: startState.C };

  for (var w = 0; w < weeks; w++) {
    // 第0周不轮换（直接使用 startState），后续每周一轮换
    if (w > 0) {
      state = {
        A: _ee_SHIFT_ROTATION[state.A],
        B: _ee_SHIFT_ROTATION[state.B],
        C: _ee_SHIFT_ROTATION[state.C]
      };
    }

    var weekStart = new Date(startMonday.getFullYear(), startMonday.getMonth(), startMonday.getDate() + w * 7);

    _ee_ROTATION_GROUPS.forEach(function(group) {
      for (var d = 0; d < 7; d++) {
        var date = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + d);
        var dateStr = Utilities.formatDate(date, "Asia/Shanghai", "yyyy-MM-dd");
        rows.push([dateStr, state[group], group]);
      }
    });
  }

  return rows;
}

// ========== 测试函数 ==========

/** 手动测试：同步今天的考勤数据 */
function testSyncAttendanceToday() {
  console.log("=== 考勤同步测试（今天） ===");
  var result = syncAttendanceFromEE();
  console.log("结果: " + JSON.stringify(result));
}

/** 手动测试：同步指定日期的考勤数据 */
function testSyncAttendanceDate(dateStr) {
  if (!dateStr) dateStr = _ee_getTargetDate();
  console.log("=== 考勤同步测试（" + dateStr + "） ===");
  // 临时覆盖目标日期
  var originalTarget = _ee_getTargetDate;
  _ee_getTargetDate = function () { return dateStr; };
  var result = syncAttendanceFromEE();
  _ee_getTargetDate = originalTarget;
  console.log("结果: " + JSON.stringify(result));
}

/** 诊断：列出 E&E 考勤表所有 sheet 名称和结构信息 */
function testDiagnoseEESheets() {
  var ss = SpreadsheetApp.openById(_ee_SOURCE_SS_ID);
  var sheets = ss.getSheets();
  console.log("=== E&E 考勤表 Sheet 诊断 ===");
  console.log("总数: " + sheets.length);
  sheets.forEach(function (s) {
    var name = s.getSheetName();
    var rows = s.getLastRow();
    var cols = s.getLastColumn();
    // 检查是否是月度考勤 sheet (名称含数字)
    var isMonthSheet = /\d{4}\.\d{1,2}/.test(name);
    var marker = isMonthSheet ? " [月度考勤]" : "";
    console.log(name + " | " + rows + "行×" + cols + "列" + marker);
  });
}

/** 诊断：查看指定日期在 E&E 中的原始数据 */
function testDiagnoseEEDate(dateStr) {
  if (!dateStr) dateStr = _ee_getTargetDate();
  console.log("=== E&E 日期诊断: " + dateStr + " ===");

  var eeSheet = _ee_findMonthSheet(dateStr);
  if (!eeSheet) { console.log("未找到对应 sheet"); return; }
  console.log("Sheet: " + eeSheet.getSheetName());

  var colIndex = _ee_findDayColumn(eeSheet, dateStr);
  console.log("列索引: " + colIndex);

  if (colIndex >= 0) {
    var data = _ee_readDailyData(eeSheet, colIndex);
    console.log("总人数: " + data.length);

    // 按工序统计
    var processStats = {};
    data.forEach(function (p) {
      if (!processStats[p.process]) processStats[p.process] = { count: 0, inDuty: 0 };
      processStats[p.process].count++;
      if (p.hours > 0) processStats[p.process].inDuty++;
    });
    console.log("工序分布: " + JSON.stringify(processStats));

    // 显示前10条
    console.log("前10条数据:");
    data.slice(0, 10).forEach(function (p) {
      console.log("  " + p.name + " | " + p.process + " | 组" + p.team + " | " + p.hours + "h | 原始值:" + p.cellRaw);
    });
  }
}

/** 手动测试：更新班次顺序（生成下3周） */
function testUpdateShiftSchedule() {
  console.log("=== 班次顺序更新测试 ===");
  var result = updateShiftSchedule();
  console.log("结果: " + JSON.stringify(result));
}
