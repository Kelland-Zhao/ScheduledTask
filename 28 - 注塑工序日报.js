// V20260613.02 — 注塑工序日报
// 入口：sendDailyProcessReport（每日 08:30 定时 or 手动）
// 数据源：_pd_TARGET_SHEET_ID → MasterData（人员工时） + TB1/TB2（开机数直接读取，独立于人员排班）
// 逻辑：获取第二天排班数据，按车间×班次汇总，HTML 邮件发送给 INJ S&C 管理层

// ========== 配置 ==========
var _dr_TEST_MODE = false;              // 已投产
var _dr_TEST_EMAIL = "kelland_zhao@colpal.com";
var _dr_SENDER_NAME = "工序日报系统";

// userID 表收件人筛选
var _dr_PERM_PROCESS_COL = 14;           // O列(0-indexed): 工序
var _dr_PERM_ROLE_COL = 15;              // P列: 职位
var _dr_PERM_EMAIL_COL = 9;              // J列: GMail
var _dr_PERM_PROCESS_VAL = "INJ";
var _dr_PERM_ROLE_VAL = "S&C";

// 班次定义（展示顺序：夜→早→中）
var _dr_SHIFT_ORDER = ["1夜", "2早", "3中"];
var _dr_SHIFT_DISPLAY = { "1夜": "夜班", "2早": "早班", "3中": "中班" };

// ========== 主入口 ==========
function sendDailyProcessReport(e) {
  var trigger = e ? "定时" : "手动";
  try {
    // 1. 计算目标日期（第二天）
    var tomorrow = tomorrowDate();
    var tomorrowDateStr = Utilities.formatDate(tomorrow, currentTimeZone, "yyyy.MM.dd");
    var tomorrowDisplay = _dr_formatDateChinese(tomorrow);

    console.log("目标日期: " + tomorrowDateStr);

    // 2. 读取 MasterData
    var ss = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    var masterSheet = ss.getSheetByName("MasterData");
    if (!masterSheet) throw new Error("未找到 MasterData 工作表");

    var lastRow = masterSheet.getLastRow();
    if (lastRow < 2) {
      writeLog("sendDailyProcessReport", "跳过", "MasterData 无数据", trigger, "");
      return;
    }

    var data = masterSheet.getRange(2, 1, lastRow - 1, 8).getValues();

    // 3. 筛选第二天数据并聚合（人员工时来自 MasterData）
    var summary = _dr_aggregateByWorkshopShift(data, tomorrowDateStr);

    // 3.5 开机数从 TB1/TB2 直接读取（独立于人员排班数据）
    var machineCounts = _dr_getMachineCounts(tomorrowDateStr);
    ["TB1", "TB2"].forEach(function(ws) {
      _dr_SHIFT_ORDER.forEach(function(sh) {
        if (machineCounts[ws] && machineCounts[ws][sh] !== undefined) {
          summary[ws][sh].machines = machineCounts[ws][sh];
        }
      });
    });

    // 4. 构建邮件 HTML
    var todayStr = formatVariableAsDate(new Date());
    var html = _dr_buildEmailHtml(summary, tomorrowDisplay, todayStr);

    // 5. 获取收件人
    var recipients = _dr_TEST_MODE ? [_dr_TEST_EMAIL] : _dr_getRecipients();

    if (recipients.length > 0) {
      var subject = "【注塑工序日报】 " + todayStr + " 注塑工序日报";

      var options = { htmlBody: html, name: _dr_SENDER_NAME };
      if (!_dr_TEST_MODE && GMAIL_CC) options.cc = GMAIL_CC;

      GmailApp.sendEmail(recipients.join(","), subject, "", options);

      // 构建 Log 摘要
      var logParts = [];
      ["TB1", "TB2"].forEach(function(ws) {
        var parts = _dr_SHIFT_ORDER.map(function(sh) {
          return _dr_SHIFT_DISPLAY[sh] + ":" + (summary[ws][sh].people || 0) + "人/" + (summary[ws][sh].manhours || 0) + "h";
        });
        logParts.push(ws + "(" + parts.join(" ") + ")");
      });

      writeLog("sendDailyProcessReport", "成功", logParts.join(" "), trigger, "TO: " + recipients.join(","));
      console.log("邮件发送成功");
    } else {
      writeLog("sendDailyProcessReport", "跳过", "无匹配收件人(O=INJ,P=S&C)", trigger, "");
      console.warn("未找到匹配收件人");
    }

  } catch (err) {
    console.error(err.stack || err.message);
    try { writeLog("sendDailyProcessReport", "失败", err.message, trigger, err.stack || ""); } catch (e2) {}
    throw err;
  }
}

// ========== 数据聚合 ==========
/**
 * 从 MasterData 筛选目标日期的行，按车间×班次聚合
 * @returns { TB1: { "1夜": {machines,people,manhours}, ... }, TB2: {...} }
 */
function _dr_aggregateByWorkshopShift(data, targetDateStr) {
  // 初始化所有车间×班次组合为 0
  var summary = {};
  ["TB1", "TB2"].forEach(function(ws) {
    summary[ws] = {};
    _dr_SHIFT_ORDER.forEach(function(sh) {
      summary[ws][sh] = { machines: 0, people: 0, manhours: 0 };
    });
  });

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateShift = String(row[0] || "").trim();
    var workshop = String(row[1] || "").trim();
    var operatingQty = Number(row[2]) || 0;
    var personName = String(row[3] || "").trim();
    var manhour = Number(row[4]) || 0;

    // 过滤：日期不匹配 or 车间不是 TB1/TB2 or 无姓名
    if (!dateShift || dateShift.indexOf(targetDateStr) !== 0) continue;
    if (["TB1", "TB2"].indexOf(workshop) === -1) continue;
    if (!personName) continue;

    // 提取班次（如 "2026.06.12_2早" → "2早"）
    var shiftPart = _dr_splitShift(dateShift);
    if (!shiftPart || _dr_SHIFT_ORDER.indexOf(shiftPart) === -1) continue;

    // 累加工时
    summary[workshop][shiftPart].manhours += manhour;
    // 记录人数（同一人同名在同一班次只算一次）
    if (!summary[workshop][shiftPart]._nameSet) {
      summary[workshop][shiftPart]._nameSet = {};
    }
    if (!summary[workshop][shiftPart]._nameSet[personName]) {
      summary[workshop][shiftPart]._nameSet[personName] = true;
      summary[workshop][shiftPart].people++;
    }
    // 开机数取最大值（同一日期班次车间下所有行相同）
    if (operatingQty > summary[workshop][shiftPart].machines) {
      summary[workshop][shiftPart].machines = operatingQty;
    }
  }

  // 清理 _nameSet
  ["TB1", "TB2"].forEach(function(ws) {
    _dr_SHIFT_ORDER.forEach(function(sh) {
      delete summary[ws][sh]._nameSet;
    });
  });

  return summary;
}

/** 从 TB1/TB2 直接读取指定日期的开机数（独立于人员排班） */
function _dr_getMachineCounts(targetDateStr) {
  var result = { TB1: {}, TB2: {} };
  try {
    var ss = SpreadsheetApp.openById(_pd_TARGET_SHEET_ID);
    ["TB1", "TB2"].forEach(function(ws) {
      var sheet = ss.getSheetByName(ws);
      if (!sheet) return;
      var lastCol = sheet.getLastColumn();
      if (lastCol < 2) return;
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var machineRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      for (var c = 1; c < headers.length; c++) {
        var dateShift = String(headers[c] || "").trim();
        if (!dateShift || dateShift.indexOf(targetDateStr) !== 0) continue;
        var shiftPart = _dr_splitShift(dateShift);
        if (!shiftPart || _dr_SHIFT_ORDER.indexOf(shiftPart) === -1) continue;
        result[ws][shiftPart] = Number(machineRow[c]) || 0;
      }
    });
  } catch (err) {
    console.error("读取开机数失败: " + err.message);
  }
  return result;
}

/** 从日期班次字符串提取班次部分 "2026.06.12_2早" → "2早" */
function _dr_splitShift(dateShift) {
  var idx = dateShift.lastIndexOf("_");
  return idx >= 0 ? dateShift.substring(idx + 1) : "";
}

// ========== 收件人 ==========
/** 从 userID 表读取 O列=INJ 且 P列=S&C 的邮箱 */
function _dr_getRecipients() {
  try {
    var sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    var maxCol = Math.max(_dr_PERM_ROLE_COL + 1, _dr_PERM_EMAIL_COL + 1);
    var data = sheet.getRange(1, 1, lastRow, maxCol).getValues();
    var recipients = [];

    for (var i = 2; i < data.length; i++) {
      var processVal = String(data[i][_dr_PERM_PROCESS_COL] || "").trim();
      var roleVal = String(data[i][_dr_PERM_ROLE_COL] || "").trim();
      var email = String(data[i][_dr_PERM_EMAIL_COL] || "").trim();

      if (processVal === _dr_PERM_PROCESS_VAL && roleVal === _dr_PERM_ROLE_VAL && email) {
        recipients.push(email.toLowerCase());
      }
    }

    console.log("匹配收件人: " + recipients.length + " 人");
    return recipients;
  } catch (err) {
    console.error("获取收件人失败: " + err.message);
    return [];
  }
}

// ========== 邮件 HTML ==========
function _dr_buildEmailHtml(summary, tomorrowDisplay, todayStr) {
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:960px;margin:0 auto">';

  // 红色标题栏（标题居左 + Logo 居右）
  html += '<div style="background:#E60012;color:white;padding:14px 28px">';
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%"><tr>';
  html += '<td style="vertical-align:middle">';
  html += '<h2 style="margin:0;font-size:20px">注塑工序日报</h2>';
  html += '<p style="margin:4px 0 0;opacity:0.85;font-size:12px">发送时间：' + todayStr + '</p>';
  html += '</td>';
  html += '<td style="width:50px;vertical-align:middle;text-align:right;padding-left:16px">';
  html += '<img src="' + _dr_COLGATE_LOGO + '" style="height:36px;display:block" alt="Colgate">';
  html += '</td></tr></table>';
  html += '</div>';

  html += '<div style="padding:20px 28px">';

  // ===== 第一部分：人员安排信息 =====
  html += '<h3 style="color:#333;border-bottom:2px solid #E60012;padding-bottom:8px;margin-top:0;font-size:17px">一、人员安排信息</h3>';
  html += '<p style="color:#888;font-size:13px;margin:8px 0 0">数据日期：' + tomorrowDisplay + '</p>';

  // TB1 / TB2 并排布局（邮件兼容 table layout）
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin-top:12px"><tr>';
  html += '<td style="width:50%;vertical-align:top;padding-right:10px">';
  html += _dr_buildWorkshopTable("TB1", summary["TB1"]);
  html += '</td>';
  html += '<td style="width:50%;vertical-align:top;padding-left:10px">';
  html += _dr_buildWorkshopTable("TB2", summary["TB2"]);
  html += '</td>';
  html += '</tr></table>';

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:28px">此邮件由 工序日报系统 自动发送</p>';
  html += '</div></div></body></html>';
  return html;
}

/** 构建单个车间的排班表格 */
function _dr_buildWorkshopTable(workshopName, shiftData) {
  var headers = ["班次", "开机数", "上班人数", "合计工时"];
  var rows = [];
  var totalMachines = 0;
  var totalPeople = 0;
  var totalManhours = 0;

  _dr_SHIFT_ORDER.forEach(function(sh) {
    var d = shiftData[sh];
    var machines = d.machines || 0;
    var people = d.people || 0;
    var manhours = Math.round((d.manhours || 0) * 10) / 10;

    rows.push([
      _dr_SHIFT_DISPLAY[sh],
      machines,
      people,
      manhours
    ]);

    totalMachines += machines;
    totalPeople += people;
    totalManhours += manhours;
  });

  // 合计行
  rows.push([
    '<b>合计</b>',
    '<b>' + totalMachines + '</b>',
    '<b>' + totalPeople + '</b>',
    '<b>' + Math.round(totalManhours * 10) / 10 + '</b>'
  ]);

  var html = '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;overflow:hidden">';

  // 车间名表头
  html += '<div style="background:#E60012;color:white;text-align:center;padding:8px;font-size:15px;font-weight:bold">' + workshopName + ' 车间</div>';

  // 数据表格
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;font-size:13px;border-collapse:collapse">';
  html += '<tr style="background:#fde0e0;color:#333;font-weight:bold">';
  headers.forEach(function(h) {
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0">' + escapeHtml(h) + '</td>';
  });
  html += '</tr>';

  for (var i = 0; i < rows.length; i++) {
    var isTotal = (i === rows.length - 1);
    var bg = isTotal ? "#fff3f3" : "#ffffff";
    var fontWeight = isTotal ? "font-weight:bold;" : "";
    var borderTop = isTotal ? "border-top:2px solid #E60012;" : "";
    html += '<tr style="background:' + bg + ';' + fontWeight + borderTop + '">';
    for (var j = 0; j < rows[i].length; j++) {
      var cellColor = isTotal ? "color:#E60012;" : "";
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5;' + cellColor + '">' + (typeof rows[i][j] === 'string' ? rows[i][j] : escapeHtml(String(rows[i][j]))) + '</td>';
    }
    html += '</tr>';
  }

  html += '</table></div>';
  return html;
}

/** 将 Date 转为中文日期 "2026年6月12日（周四）" */
function _dr_formatDateChinese(date) {
  var weekDays = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"];
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  var w = weekDays[date.getDay()];
  return y + "年" + m + "月" + d + "日（" + w + "）";
}
