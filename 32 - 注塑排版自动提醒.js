// V20260713.01 — 注塑排班自动提醒
// 入口：checkAttendanceSchedule（周一~周五 9:00 定时 or 手动）
// 数据源：E&E电子考勤记录 → 仅 INJ 工序 → 整周未打则邮件提醒考勤员
// 邮件：To=考勤员, CC=直线上级+line manager, 含E&E超链接

// ========== 常量 ==========
var _sc_SOURCE_SS_ID = "1dMON_DEcAUH9xRsfOkEF37fIN7DuyVHfNwOoUyd-V-0";
var _sc_USER_SS_ID = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
var _sc_LINE_MANAGER = "kelland_zhao@colpal.com";
var _sc_TARGET_PROCESS = "INJ";
var _sc_SENDER_NAME = "考勤排班提醒系统";
var _sc_EE_URL = "https://docs.google.com/spreadsheets/d/1dMON_DEcAUH9xRsfOkEF37fIN7DuyVHfNwOoUyd-V-0/edit";

var _sc_TEST_MODE = true;
var _sc_TEST_EMAIL = "kelland_zhao@colpal.com";

// ========== 主入口 ==========
/**
 * 检查当周 INJ 工序考勤填写情况，整周未打则邮件提醒考勤员
 * @param {Object} e - 定时触发传入的事件对象 { triggerType: "scheduled" }，手动调用可不传
 */
function checkAttendanceSchedule(e) {
  var trigger = e && e.triggerType ? "定时" : "手动";
  var startTime = new Date();

  try {
    // 1. 获取当前周日期范围（周一~周日）
    var weekInfo = _sc_getCurrentWeek();
    console.log("检查周: " + weekInfo.monday + " ~ " + weekInfo.sunday);

    // 2. 预读 userID 映射（考勤员姓名 → 邮箱 + 上级邮箱）
    var userMap = _sc_buildUserMap();
    console.log("userID 映射: " + Object.keys(userMap).length + " 人");

    // 3. 逐日读取 E&E 数据，累积到 personMap（跨月自动切换sheet）
    var personMap = {};     // sapID → { name, process, team, clerk, dailyValues: { dateStr: value } }
    var sheetCache = {};    // monthKey → sheet | null
    var skippedMonths = {}; // monthKey → [dates]

    weekInfo.dates.forEach(function (dateStr) {
      var monthKey = dateStr.substring(0, 7);
      if (!(monthKey in sheetCache)) {
        sheetCache[monthKey] = _sc_findMonthSheet(dateStr) || null;
      }
      var eeSheet = sheetCache[monthKey];
      if (!eeSheet) {
        if (!skippedMonths[monthKey]) skippedMonths[monthKey] = [];
        skippedMonths[monthKey].push(dateStr);
        return;
      }

      var colIndex = _sc_findDayColumn(eeSheet, dateStr);
      if (colIndex < 0) {
        console.warn("跳过 " + dateStr + "：未找到日期列");
        return;
      }

      var dayData = _sc_readDayData(eeSheet, colIndex);
      dayData.forEach(function (p) {
        var sid = p.sapID;
        if (!personMap[sid]) {
          personMap[sid] = {
            name: p.name,
            process: p.process,
            team: p.team,
            clerk: p.clerk,
            dailyValues: {}
          };
        }
        personMap[sid].dailyValues[dateStr] = p.cellValue;
      });

      console.log(dateStr + ": " + dayData.length + " 人");
    });

    // 汇总未找到月度 sheet 的日期
    Object.keys(skippedMonths).forEach(function (mk) {
      console.warn("未找到 " + mk + " 月度 sheet，跳过: " + skippedMonths[mk].join(", "));
    });

    // 4. 筛选 INJ 工序 + 整周7天全空 → 按考勤员归类
    var clerkGroups = {}; // clerkName → [{ name, sapID, team }]
    var injCount = 0;

    Object.keys(personMap).forEach(function (sid) {
      var p = personMap[sid];
      if (p.process !== _sc_TARGET_PROCESS) return;
      injCount++;

      // 检查7天是否全部为空
      var allEmpty = weekInfo.dates.every(function (ds) {
        var v = p.dailyValues[ds];
        return v === undefined || String(v).trim() === "";
      });

      if (allEmpty) {
        var clerkName = p.clerk || "未知考勤员";
        if (!clerkGroups[clerkName]) clerkGroups[clerkName] = [];
        clerkGroups[clerkName].push({
          name: p.name,
          sapID: sid,
          team: p.team
        });
      }
    });

    console.log("INJ 总人数: " + injCount + ", 未打考勤员数: " + Object.keys(clerkGroups).length);

    // 5. 逐考勤员发送邮件
    var sentCount = 0;
    var skipCount = 0;
    var clerkNames = Object.keys(clerkGroups);

    if (clerkNames.length === 0) {
      console.log("无需提醒，INJ 工序考勤均已填写");
    }

    clerkNames.forEach(function (clerkName) {
      var undaList = clerkGroups[clerkName];
      var userInfo = userMap[clerkName];

      if (!userInfo || !userInfo.email) {
        console.warn("考勤员 " + clerkName + " 在 userID 中无邮箱，跳过发送（" + undaList.length + " 人未打）");
        skipCount++;
        return;
      }

      var to = _sc_TEST_MODE ? _sc_TEST_EMAIL : userInfo.email;
      var ccList = [];
      if (!_sc_TEST_MODE) {
        if (userInfo.supervisor) ccList.push(userInfo.supervisor);
        ccList.push(_sc_LINE_MANAGER);
      }

      var html = _sc_buildEmailHtml(clerkName, undaList, weekInfo);
      var subject = "【考勤安排提醒】 " + weekInfo.monday + " 周";

      try {
        var options = { htmlBody: html, name: _sc_SENDER_NAME };
        if (ccList.length > 0) options.cc = ccList.join(",");
        GmailApp.sendEmail(to, subject, "", options);
        sentCount++;
        console.log("已发送: " + clerkName + " → " + to + (ccList.length > 0 ? " CC:" + ccList.join(",") : ""));
      } catch (mailErr) {
        console.error("邮件发送失败: " + clerkName + " - " + mailErr.message);
        try { writeLog("checkAttendanceSchedule", "异常", clerkName + " 邮件发送失败: " + mailErr.message, trigger, ""); } catch (e3) {}
      }
    });

    var duration = ((new Date()) - startTime) / 1000;
    var logMsg = "INJ未打" + clerkNames.length + "组/发送" + sentCount + "封" + (skipCount > 0 ? "/跳过" + skipCount : "") + "，耗时" + duration + "s";
    console.log(logMsg);

    try {
      writeLog("checkAttendanceSchedule", clerkNames.length > 0 ? "成功" : "无需提醒", logMsg, trigger, "");
    } catch (e2) {}

    return {
      success: true,
      clerksTotal: clerkNames.length,
      sent: sentCount,
      skipped: skipCount,
      duration: duration
    };

  } catch (err) {
    console.error("考勤排班检查失败: " + err.message);
    try { writeLog("checkAttendanceSchedule", "失败", err.message, trigger, ""); } catch (e2) {}
    return { success: false, error: err.message };
  }
}

// ========== 日期工具 ==========

/** 获取当前周（周一~周日）的日期列表 */
function _sc_getCurrentWeek() {
  var now = new Date();
  var dayOfWeek = now.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  var daysFromMonday = dayOfWeek === 0 ? 6 : (dayOfWeek - 1);
  var monday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysFromMonday);

  var dates = [];
  for (var i = 0; i < 7; i++) {
    var d = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i);
    dates.push(Utilities.formatDate(d, "Asia/Shanghai", "yyyy-MM-dd"));
  }

  return {
    monday: dates[0],
    sunday: dates[6],
    dates: dates
  };
}

// ========== E&E Sheet 定位（参考 31 号脚本 _ee_findMonthSheet / _ee_findDayColumn） ==========

/**
 * 根据日期定位 E&E 考勤表中的月度 sheet
 * 尝试多种命名模式：YYYY.MM → YYYY.MM月 → YYYY.M → YYYY.M月
 */
function _sc_findMonthSheet(targetDate) {
  var parts = targetDate.split("-");
  var year = parts[0];
  var month = parts[1];
  var monthNoPad = String(parseInt(month, 10));

  var ss = SpreadsheetApp.openById(_sc_SOURCE_SS_ID);
  var patterns = [
    year + "." + month,
    year + "." + month + "月",
    year + "." + monthNoPad,
    year + "." + monthNoPad + "月"
  ];

  for (var i = 0; i < patterns.length; i++) {
    var sheet = ss.getSheetByName(patterns[i]);
    if (sheet) return sheet;
  }

  // 模糊搜索
  var allSheets = ss.getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    var name = allSheets[j].getSheetName();
    if (name.indexOf(year + "." + month) === 0 || name.indexOf(year + "." + monthNoPad) === 0) {
      return allSheets[j];
    }
  }

  return null;
}

/**
 * 在 E&E sheet 中找到目标日期对应的列索引（0-indexed）
 * Row 4 是日期表头，G列(6)~AK列(36)对应 day 1~31，AL列(37)起为次月溢出
 */
function _sc_findDayColumn(eeSheet, targetDate) {
  var targetDay = parseInt(targetDate.split("-")[2], 10);
  var targetDayStr = String(targetDay);

  var headerRow = eeSheet.getRange(4, 1, 1, eeSheet.getLastColumn()).getValues()[0];

  // 当月列 G(6)~AK(36)
  for (var c = 6; c <= 36 && c < headerRow.length; c++) {
    if (String(headerRow[c] || "").trim() === targetDayStr) return c;
  }

  // 次月溢出列 AL(37)+
  for (var c2 = 37; c2 < headerRow.length; c2++) {
    if (String(headerRow[c2] || "").trim() === targetDayStr) return c2;
  }

  return -1;
}

// ========== E&E 数据读取 ==========

/**
 * 读取 E&E sheet 指定列的当日人员数据
 * 列结构：B=工号(1), C=姓名(2), D=工序(3), E=班别(4), F=考勤员(5)
 * @returns {Array<{sapID, name, process, team, clerk, cellValue}>}
 */
function _sc_readDayData(eeSheet, colIndex) {
  var lastRow = eeSheet.getLastRow();
  if (lastRow < 5) return [];

  // 读取固定列 B~F（0-indexed: 1~5） + 目标日列
  var fixedData = eeSheet.getRange(5, 2, lastRow - 4, 5).getValues();
  var dayValues = eeSheet.getRange(5, colIndex + 1, lastRow - 4, 1).getValues();

  var result = [];
  for (var i = 0; i < fixedData.length; i++) {
    var sapID = String(fixedData[i][0] || "").trim();
    var name = String(fixedData[i][1] || "").trim();
    var process = String(fixedData[i][2] || "").trim();
    var team = String(fixedData[i][3] || "").trim();
    var clerk = String(fixedData[i][4] || "").trim();

    if (!name) continue;

    result.push({
      sapID: sapID,
      name: name,
      process: process,
      team: team,
      clerk: clerk,
      cellValue: String(dayValues[i][0] || "").trim()
    });
  }

  return result;
}

// ========== userID 映射 ==========

/**
 * 构建 userID 姓名 → 邮箱 + 直线上级 映射
 * userID 列: B=NAME(1), J=GMail(9), BI=直线上级邮箱(61, 0-indexed)
 * @returns {Object} { "曹悦": { email: "...", supervisor: "..." } }
 */
function _sc_buildUserMap() {
  var ws = SpreadsheetApp.openById(_sc_USER_SS_ID).getSheetByName("userID");
  if (!ws) return {};

  var values = ws.getDataRange().getValues();
  var map = {};

  for (var i = 2; i < values.length; i++) {
    var name = String(values[i][1] || "").trim();
    var email = String(values[i][9] || "").trim();
    var supervisor = String(values[i][61] || "").trim();

    if (name && email) {
      map[name] = {
        email: email,
        supervisor: supervisor
      };
    }
  }

  return map;
}

// ========== 邮件 HTML 构建（参考 邮件UI规范 §3 模式A + §6.7 通用表格 + 22号脚本样式） ==========

/**
 * 构建提醒邮件 HTML
 * @param {string} clerkName - 考勤员姓名
 * @param {Array<{name, sapID, team}>} undaList - 未打员工列表
 * @param {{monday, sunday, dates}} weekInfo - 周信息
 */
function _sc_buildEmailHtml(clerkName, undaList, weekInfo) {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm");
  var weekRange = weekInfo.monday + " ~ " + weekInfo.sunday;

  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto">';

  // ===== 头部（模式A 含Logo，参考 22号脚本） =====
  html += '<div style="background:#E60012;color:white;padding:14px 28px">';
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%"><tr>';
  html += '<td style="vertical-align:middle">';
  html += '<h2 style="margin:0;font-size:20px">考勤安排提醒</h2>';
  html += '<p style="margin:4px 0 0;opacity:0.85;font-size:12px">发送时间：' + nowStr + '</p>';
  html += '</td>';
  html += '<td style="width:50px;vertical-align:middle;text-align:right;padding-left:16px">';
  html += '<img src="' + _dr_COLGATE_LOGO + '" style="height:36px;display:block" alt="Colgate">';
  html += '</td></tr></table>';
  html += '</div>';

  // ===== 正文 =====
  html += '<div style="padding:20px 28px">';

  // 考勤员信息
  html += '<p style="color:#888;font-size:13px;margin:0 0 4px">日期：' + weekInfo.monday + ' 周（' + weekRange + '）</p>';
  html += '<p style="font-size:14px;color:#34495e;margin:0 0 16px">';
  html += '考勤员：<b>' + _sc_escapeHtml(clerkName) + '</b>，以下员工当周考勤尚未填写：';
  html += '</p>';

  // 未打员工表格
  html += '<table style="width:100%;font-size:13px;border-collapse:collapse">';
  html += '<tr style="background:#E60012;color:white">';
  html += '<th style="padding:10px;text-align:left">姓名</th>';
  html += '<th style="padding:10px;text-align:left">工号</th>';
  html += '<th style="padding:10px;text-align:left">班别</th>';
  html += '</tr>';

  for (var i = 0; i < undaList.length; i++) {
    var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    html += '<tr style="background:' + bg + '">';
    html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _sc_escapeHtml(undaList[i].name) + '</td>';
    html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _sc_escapeHtml(undaList[i].sapID) + '</td>';
    html += '<td style="padding:10px;border-bottom:1px solid #ecf0f1">' + _sc_escapeHtml(undaList[i].team) + '组</td>';
    html += '</tr>';
  }

  html += '</table>';

  // 提示 + E&E 超链接
  html += '<p style="color:#e74c3c;font-size:13px;margin-top:12px">';
  html += '共 <b>' + undaList.length + '</b> 人未打，请尽快在 ';
  html += '<a href="' + _sc_EE_URL + '" style="color:#E60012;text-decoration:underline">E&E 电子考勤记录</a>';
  html += ' 中补录以上员工的考勤数据。';
  html += '</p>';

  html += '</div>';

  // ===== 尾部（§12.1 模式A） =====
  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px;padding:0 24px 16px">';
  html += '此邮件由 ' + _sc_SENDER_NAME + ' 自动发送，请勿回复。';
  html += '</p>';

  html += '</div></body></html>';
  return html;
}

/** HTML 转义 */
function _sc_escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ========== 测试函数 ==========

/** 手动测试：检查当周考勤并发送提醒 */
function testCheckAttendanceSchedule() {
  console.log("=== 注塑排班提醒测试 ===");
  var result = checkAttendanceSchedule();
  console.log("结果: " + JSON.stringify(result));
}

/** 诊断：查看当前周的 E&E 原始数据 */
function testDiagnoseCurrentWeek() {
  var weekInfo = _sc_getCurrentWeek();
  console.log("=== 当前周诊断: " + weekInfo.monday + " ~ " + weekInfo.sunday + " ===");

  var sheetCache = {};
  weekInfo.dates.forEach(function (dateStr) {
    var monthKey = dateStr.substring(0, 7);
    if (!(monthKey in sheetCache)) {
      sheetCache[monthKey] = _sc_findMonthSheet(dateStr) || null;
    }
    var eeSheet = sheetCache[monthKey];
    if (!eeSheet) {
      console.warn(dateStr + ": 未找到 sheet");
      return;
    }

    var colIndex = _sc_findDayColumn(eeSheet, dateStr);
    if (colIndex < 0) {
      console.warn(dateStr + ": 未找到列");
      return;
    }

    var data = _sc_readDayData(eeSheet, colIndex);
    var injData = data.filter(function (p) { return p.process === "INJ"; });
    var emptyCount = injData.filter(function (p) { return !p.cellValue; }).length;
    console.log(dateStr + " | sheet=" + eeSheet.getSheetName() + " | 总" + data.length + "人 INJ" + injData.length + "人 空" + emptyCount + "人");
  });

  // userID 映射检查
  var userMap = _sc_buildUserMap();
  var clerkNames = ["曹悦"]; // 示例
  clerkNames.forEach(function (name) {
    var info = userMap[name];
    console.log("考勤员 " + name + ": " + (info ? "email=" + info.email + " supervisor=" + info.supervisor : "未找到"));
  });
}
