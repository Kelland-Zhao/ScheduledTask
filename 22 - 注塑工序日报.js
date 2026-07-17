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

    // 4. 备件信息
    var sparePartsInfo = _dr_getSparePartsInfo();

    // 4.5 注塑测试信息
    var testInfo = _dr_getTestInfo(tomorrow);

    // 5. 构建邮件 HTML
    var todayStr = formatVariableAsDate(new Date());
    var html = _dr_buildEmailHtml(summary, sparePartsInfo, testInfo, tomorrowDisplay, todayStr);

    // 6. 获取收件人
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

// ========== 备件信息 ==========
var _dr_SPAREPARTS_SHEET_ID = "1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs";

/** 从备件管理表读取备件信息数据 */
function _dr_getSparePartsInfo() {
  var result = {
    totalValue: 0,
    dataDate: "",
    prevDate: "",
    newArrivals: [],
    offlineNoStock: []
  };

  try {
    var ss = SpreadsheetApp.openById(_dr_SPAREPARTS_SHEET_ID);

    // 1. 读取备件基础信息
    var biSheet = ss.getSheetByName("备件基础信息");
    if (!biSheet || biSheet.getLastRow() <= 1) return result;
    var biData = biSheet.getRange(2, 1, biSheet.getLastRow() - 1, 8).getValues();

    var totalValue = 0;
    var matValueMap = {}; // 物料→库存金额
    biData.forEach(function(row) {
      var mat = String(row[0] || "").trim();
      if (!mat) return;
      var stock = Number(row[4]) || 0;
      var value = Number(row[7]) || 0;
      totalValue += value;
      matValueMap[mat] = value;

      // Offline控制=Y 且 库存=0
      var offlineControl = String(row[5] || "").trim().toUpperCase();
      if (offlineControl === "Y" && stock === 0) {
        result.offlineNoStock.push({
          material: mat,
          description: String(row[1] || "").trim(),
          machineModel: String(row[2] || "").trim(),
          safetyStock: Number(row[3]) || 0
        });
      }
    });
    result.totalValue = Math.round(totalValue * 100) / 100;

    // 2. 读取 MasterData 获取最新日期和前一天对比
    var mdSheet = ss.getSheetByName("MasterData");
    if (!mdSheet || mdSheet.getLastRow() <= 1) return result;
    var mdData = mdSheet.getRange(2, 1, mdSheet.getLastRow() - 1, 5).getValues();

    // 收集各日期各物料的库存（取当日最新记录）
    var dateMap = {};
    mdData.forEach(function(row) {
      var mat = String(row[0] || "").trim();
      if (!mat || mat.charAt(0).toUpperCase() === "Z") return;
      var desc = String(row[1] || "").trim();
      var stock = Number(row[2]) || 0;
      var dateVal = row[4];
      if (!dateVal) return;
      var dateStr = dateVal instanceof Date
        ? Utilities.formatDate(dateVal, currentTimeZone, "yyyy-MM-dd")
        : String(dateVal).trim();
      if (!dateMap[dateStr]) dateMap[dateStr] = {};
      if (!dateMap[dateStr][mat] || dateVal > dateMap[dateStr][mat]._rawDate) {
        dateMap[dateStr][mat] = { desc: desc, stock: stock, _rawDate: dateVal };
      }
    });

    // 排序日期取最近2天
    var dates = Object.keys(dateMap).sort().reverse();
    result.dataDate = dates[0] || "";

    if (dates.length >= 2) {
      result.prevDate = dates[1];
      _dr_compareDates(dateMap, dates[0], dates[1], matValueMap, result.newArrivals);
    }

    result.newArrivals.sort(function(a, b) { return b.value - a.value; });
    result.offlineNoStock.sort(function(a, b) { return a.material.localeCompare(b.material); });

  } catch (err) {
    console.error("读取备件信息失败: " + err.message);
  }

  return result;
}

/** 对比两天 MasterData，找出库存增加的物料，写入 targetArr */
function _dr_compareDates(dateMap, newDate, oldDate, matValueMap, targetArr) {
  var newMap = dateMap[newDate];
  var oldMap = dateMap[oldDate];
  Object.keys(newMap).forEach(function(mat) {
    var newStock = newMap[mat].stock;
    var oldStock = (oldMap[mat] && oldMap[mat].stock) || 0;
    if (newStock > oldStock) {
      targetArr.push({
        material: mat,
        description: newMap[mat].desc,
        latestStock: newStock,
        prevStock: oldStock,
        increase: newStock - oldStock,
        value: matValueMap[mat] || 0
      });
    }
  });
}

/** 构建单组新到货表格 */
function _dr_buildNewArrivalTable(title, newDate, oldDate, arrivals) {
  var html = '';
  html += '<h4 style="color:#555;margin-top:16px;margin-bottom:8px;font-size:15px">' + escapeHtml(title) + '（' + escapeHtml(newDate) + ' vs ' + escapeHtml(oldDate) + '）</h4>';

  if (arrivals.length > 0) {
    html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;font-size:13px;border-collapse:collapse">';
    html += '<tr style="background:#fde0e0;color:#333;font-weight:bold">';
    ["物料编码", "物料描述", "前日库存", "当前库存", "增加", "库存金额"].forEach(function(h) {
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0">' + escapeHtml(h) + '</td>';
    });
    html += '</tr>';

    var stPrevStock = 0, stLatestStock = 0, stIncrease = 0, stValue = 0;
    arrivals.forEach(function(item, idx) {
      stPrevStock += item.prevStock;
      stLatestStock += item.latestStock;
      stIncrease += item.increase;
      stValue += item.value;

      var bg = idx % 2 === 0 ? "#ffffff" : "#fff8f8";
      html += '<tr style="background:' + bg + '">';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + escapeHtml(item.material) + '</td>';
      html += '<td style="padding:8px;text-align:left;border-bottom:1px solid #f5f5f5">' + escapeHtml(item.description) + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + item.prevStock + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + item.latestStock + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5;color:#E60012;font-weight:bold">+' + item.increase + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">¥' + _dr_formatMoney(item.value) + '</td>';
      html += '</tr>';
    });

    // 小计行
    html += '<tr style="background:#fff3f3;border-top:2px solid #E60012;font-weight:bold">';
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0" colspan="2">小计</td>';
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0;color:#E60012">' + stPrevStock + '</td>';
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0;color:#E60012">' + stLatestStock + '</td>';
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0;color:#E60012">+' + stIncrease + '</td>';
    html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0;color:#E60012">¥' + _dr_formatMoney(stValue) + '</td>';
    html += '</tr>';

    html += '</table>';
  } else {
    html += '<p style="color:#999;font-size:13px;margin:8px 0">无新增到货</p>';
  }

  return html;
}

/** 构建备件信息 HTML 块 */
function _dr_buildSparePartsSection(sp) {
  var html = '';

  // ===== 第二部分：备件信息 =====
  html += '<h3 style="color:#333;border-bottom:2px solid #E60012;padding-bottom:8px;margin-top:28px;font-size:17px">二、备件信息</h3>';

  // 2.1 库存总览
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin-top:12px"><tr>';

  // 左：库存总金额卡片
  html += '<td style="width:50%;vertical-align:top;padding-right:10px">';
  html += '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:16px;text-align:center">';
  html += '<p style="margin:0;color:#888;font-size:12px">库存总金额</p>';
  html += '<p style="margin:8px 0 0;font-size:28px;font-weight:bold;color:#E60012">¥' + _dr_formatMoney(sp.totalValue) + '</p>';
  html += '</div></td>';

  // 右：数据日期卡片
  html += '<td style="width:50%;vertical-align:top;padding-left:10px">';
  html += '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:16px;text-align:center">';
  html += '<p style="margin:0;color:#888;font-size:12px">数据日期</p>';
  html += '<p style="margin:8px 0 0;font-size:28px;font-weight:bold;color:#333">' + escapeHtml(sp.dataDate || "-") + '</p>';
  html += '</div></td>';

  html += '</tr></table>';

  // 2.2 VS前一天新到备件
  html += _dr_buildNewArrivalTable("VS前一天新到备件", sp.dataDate, sp.prevDate, sp.newArrivals);

  // 2.3 建议Offline控制但无库存
  html += '<h4 style="color:#555;margin-top:20px;margin-bottom:8px;font-size:15px">建议 Offline 控制但无库存的备件</h4>';
  if (sp.offlineNoStock.length > 0) {
    html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;font-size:13px;border-collapse:collapse">';
    html += '<tr style="background:#fde0e0;color:#333;font-weight:bold">';
    ["物料编码", "物料描述", "机型", "安全库存"].forEach(function(h) {
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0">' + escapeHtml(h) + '</td>';
    });
    html += '</tr>';
    sp.offlineNoStock.forEach(function(item, idx) {
      var bg = "#FFC107"; // 全部高亮黄色
      html += '<tr style="background:' + bg + ';font-weight:bold">';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #e6ac00">' + escapeHtml(item.material) + '</td>';
      html += '<td style="padding:8px;text-align:left;border-bottom:1px solid #e6ac00">' + escapeHtml(item.description) + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #e6ac00">' + escapeHtml(item.machineModel) + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #e6ac00">' + item.safetyStock + '</td>';
      html += '</tr>';
    });
    html += '</table>';
  } else {
    html += '<p style="color:#999;font-size:13px;margin:8px 0">无</p>';
  }

  return html;
}

// ========== 注塑测试信息 ==========
var _dr_TEST_SPREADSHEET_ID = "17ys3UDFWjhfaPnk0TErqqeU0FnMP7nsRoRsTmlmm2fg";

/** 从注塑测试表读取明天的测试记录并返回摘要 */
function _dr_getTestInfo(targetDate) {
  var result = {
    total: 0,
    problemCount: 0,
    records: []
  };

  try {
    var ss = SpreadsheetApp.openById(_dr_TEST_SPREADSHEET_ID);
    var testSheet = ss.getSheetByName("注塑测试");
    if (!testSheet || testSheet.getLastRow() <= 1) return result;

    // 读取 21 列（A-U），与 31 号文件列索引一致
    var lastRow = testSheet.getLastRow();
    var data = testSheet.getRange(2, 1, lastRow - 1, 21).getValues();

    // 将目标日期转为 yyyy-MM-dd 用于比较
    var targetKey = Utilities.formatDate(targetDate, currentTimeZone, "yyyy-MM-dd");

    data.forEach(function(row, index) {
      // 日期列 B（索引 1）
      var dateKey = "";
      var dateVal = row[1];
      if (dateVal instanceof Date) {
        dateKey = Utilities.formatDate(dateVal, currentTimeZone, "yyyy-MM-dd");
      } else if (dateVal) {
        dateKey = String(dateVal).trim();
      }
      if (dateKey !== targetKey) return;

      // 跳过满产/模具不在线/取消的记录
      var status = String(row[15] || "").trim();
      if (status === "满产" || status === "模具不在线" || status === "取消") return;

      result.total++;

      // 检查缺失项
      var problems = [];
      if (!String(row[11] || "").trim()) problems.push("测试机台未维护");
      if (!String(row[14] || "").trim()) problems.push("测试负责人未维护");
      if (!String(row[16] || "").trim()) problems.push("测试样单链接未维护");

      if (problems.length > 0) result.problemCount++;

      result.records.push({
        rowNumber: index + 2,
        product: String(row[5] || "").trim(),
        description: String(row[6] || "").trim(),
        machine: String(row[11] || "").trim(),
        projectOwner: String(row[12] || "").trim(),
        testOwner: String(row[14] || "").trim(),
        status: status,
        problems: problems
      });
    });

  } catch (err) {
    console.error("读取注塑测试信息失败: " + err.message);
  }

  return result;
}

/** 构建注塑测试信息 HTML 部分 */
function _dr_buildTestInfoSection(testInfo) {
  var html = '';

  html += '<h3 style="color:#333;border-bottom:2px solid #E60012;padding-bottom:8px;margin-top:28px;font-size:17px">三、注塑测试信息</h3>';

  // 摘要卡片
  html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin-top:12px"><tr>';
  html += '<td style="width:33%;vertical-align:top;padding-right:10px">';
  html += '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:14px;text-align:center">';
  html += '<p style="margin:0;color:#888;font-size:12px">明日测试数</p>';
  html += '<p style="margin:8px 0 0;font-size:28px;font-weight:bold;color:#333">' + testInfo.total + '</p>';
  html += '</div></td>';

  html += '<td style="width:33%;vertical-align:top;padding:0 5px">';
  html += '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:14px;text-align:center">';
  html += '<p style="margin:0;color:#888;font-size:12px">正常</p>';
  html += '<p style="margin:8px 0 0;font-size:28px;font-weight:bold;color:#27ae60">' + (testInfo.total - testInfo.problemCount) + '</p>';
  html += '</div></td>';

  html += '<td style="width:33%;vertical-align:top;padding-left:10px">';
  html += '<div style="background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:14px;text-align:center">';
  html += '<p style="margin:0;color:#888;font-size:12px">缺失提醒</p>';
  var problemColor = testInfo.problemCount > 0 ? "#e74c3c" : "#27ae60";
  html += '<p style="margin:8px 0 0;font-size:28px;font-weight:bold;color:' + problemColor + '">' + testInfo.problemCount + '</p>';
  html += '</div></td>';
  html += '</tr></table>';

  // 明细表（仅显示有问题或全部记录）
  if (testInfo.records.length > 0) {
    // 表头
    html += '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;font-size:13px;border-collapse:collapse;margin-top:12px">';
    html += '<tr style="background:#fde0e0;color:#333;font-weight:bold">';
    ["#", "产品名称", "测试说明", "机台", "测试负责人", "状态", "缺失项"].forEach(function(h) {
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f0d0d0">' + escapeHtml(h) + '</td>';
    });
    html += '</tr>';

    testInfo.records.forEach(function(r, idx) {
      var bg = r.problems.length > 0 ? "#FFC107" : (idx % 2 === 0 ? "#ffffff" : "#fff8f8");
      html += '<tr style="background:' + bg + '">';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + (idx + 1) + '</td>';
      html += '<td style="padding:8px;text-align:left;border-bottom:1px solid #f5f5f5">' + escapeHtml(r.product || "-") + '</td>';
      html += '<td style="padding:8px;text-align:left;border-bottom:1px solid #f5f5f5">' + escapeHtml(r.description || "-") + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + escapeHtml(r.machine || "-") + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + escapeHtml(r.testOwner || "-") + '</td>';
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5">' + escapeHtml(r.status || "-") + '</td>';
      html += '<td style="padding:8px;text-align:left;border-bottom:1px solid #f5f5f5;color:#e74c3c">' + escapeHtml(r.problems.join("，") || "无") + '</td>';
      html += '</tr>';
    });

    html += '</table>';
  } else {
    html += '<p style="color:#999;font-size:13px;margin:12px 0">明日无测试计划</p>';
  }

  return html;
}

/** 金额格式化（#,###.## 无小数则省略） */
function _dr_formatMoney(val) {
  var parts = val.toFixed(2).split(".");
  var intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  var decPart = parts[1];
  return decPart === "00" ? intPart : intPart + "." + decPart;
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

/** 计算单个车间三班合计 */
function _dr_workshopTotals(shiftData) {
  var machines = 0, people = 0, manhours = 0;
  _dr_SHIFT_ORDER.forEach(function(sh) {
    machines += shiftData[sh].machines || 0;
    people += shiftData[sh].people || 0;
    manhours += shiftData[sh].manhours || 0;
  });
  manhours = Math.round(manhours * 10) / 10;
  var mpp = people > 0 ? Math.round(machines / people * 100) / 100 : 0;
  return { machines: machines, people: people, manhours: manhours, mpp: mpp };
}

/** 构建顶部摘要条（4卡片横排） */
function _dr_buildSummaryBar(tb1, tb2, spareParts, testInfo) {
  var cardStyle = 'background:#FFF9F9;border:1px solid #f0d0d0;border-radius:4px;padding:10px 8px;text-align:center;height:84px';
  var titleStyle = 'margin:0;color:#E60012;font-size:11px;font-weight:bold;line-height:1.2';
  var mainStyle = 'margin:6px 0 0;font-size:19px;font-weight:bold;color:#333;line-height:1.3';
  var subStyle = 'margin:4px 0 0;font-size:12px;color:#666;line-height:1.3';

  var html = '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin-bottom:4px"><tr>';

  // TB1 卡片
  html += '<td style="width:25%;vertical-align:top;padding-right:6px">';
  html += '<div style="' + cardStyle + '">';
  html += '<p style="' + titleStyle + '">TB1 车间</p>';
  html += '<p style="' + mainStyle + '">' + tb1.machines + '台  ' + tb1.people + '人</p>';
  html += '<p style="' + subStyle + '">' + tb1.manhours + 'h  人均' + tb1.mpp + '台</p>';
  html += '</div></td>';

  // TB2 卡片
  html += '<td style="width:25%;vertical-align:top;padding:0 3px">';
  html += '<div style="' + cardStyle + '">';
  html += '<p style="' + titleStyle + '">TB2 车间</p>';
  html += '<p style="' + mainStyle + '">' + tb2.machines + '台  ' + tb2.people + '人</p>';
  html += '<p style="' + subStyle + '">' + tb2.manhours + 'h  人均' + tb2.mpp + '台</p>';
  html += '</div></td>';

  // 备件库存卡片
  html += '<td style="width:25%;vertical-align:top;padding:0 3px">';
  html += '<div style="' + cardStyle + '">';
  html += '<p style="' + titleStyle + '">备件库存</p>';
  html += '<p style="' + mainStyle + '">¥' + _dr_formatMoney(spareParts.totalValue) + '</p>';
  html += '<p style="' + subStyle + '">数据日期：' + escapeHtml(spareParts.dataDate || "-") + '</p>';
  html += '</div></td>';

  // 明日测试卡片
  html += '<td style="width:25%;vertical-align:top;padding-left:6px">';
  html += '<div style="' + cardStyle + '">';
  html += '<p style="' + titleStyle + '">明日测试</p>';
  html += '<p style="' + mainStyle + '">' + testInfo.total + '项</p>';
  var testSubColor = testInfo.problemCount > 0 ? "#e74c3c" : "#27ae60";
  html += '<p style="margin:4px 0 0;font-size:12px;color:' + testSubColor + '">' + (testInfo.problemCount > 0 ? '⚠ ' + testInfo.problemCount + ' 缺失' : '✓ 全部就绪') + '</p>';
  html += '</div></td>';

  html += '</tr></table>';
  return html;
}

function _dr_buildEmailHtml(summary, spareParts, testInfo, tomorrowDisplay, todayStr) {
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

  // ===== 顶部摘要条 =====
  var tb1Totals = _dr_workshopTotals(summary["TB1"]);
  var tb2Totals = _dr_workshopTotals(summary["TB2"]);
  html += _dr_buildSummaryBar(tb1Totals, tb2Totals, spareParts, testInfo);

  // ===== 第一部分：人员安排信息 =====
  html += '<h3 style="color:#333;border-bottom:2px solid #E60012;padding-bottom:8px;margin-top:20px;font-size:17px">一、人员安排信息</h3>';
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

  // ===== 第二部分：备件信息 =====
  html += _dr_buildSparePartsSection(spareParts);

  // ===== 第三部分：注塑测试信息 =====
  html += _dr_buildTestInfoSection(testInfo);

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
    // 开机数>0 但无人排班 → 标记预警
    var warn = machines > 0 && people === 0;

    rows.push({
      cells: [ _dr_SHIFT_DISPLAY[sh], machines, people, manhours ],
      warn: warn
    });

    totalMachines += machines;
    totalPeople += people;
    totalManhours += manhours;
  });

  // 合计行
  rows.push({
    cells: [
      '<b>合计</b>',
      '<b>' + totalMachines + '</b>',
      '<b>' + totalPeople + '</b>',
      '<b>' + Math.round(totalManhours * 10) / 10 + '</b>'
    ],
    isTotal: true
  });

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
    var row = rows[i];
    var isTotal = row.isTotal;
    var bg = isTotal ? "#fff3f3" : "#ffffff";
    var fontWeight = isTotal ? "font-weight:bold;" : "";
    var borderTop = isTotal ? "border-top:2px solid #E60012;" : "";
    html += '<tr style="background:' + bg + ';' + fontWeight + borderTop + '">';
    for (var j = 0; j < row.cells.length; j++) {
      var cellColor = isTotal ? "color:#E60012;" : "";
      // 开机数>0 但无人排班 → 上班人数(j=2)和合计工时(j=3)标黄预警
      var cellBg = "";
      if (!isTotal && row.warn && (j === 2 || j === 3)) {
        cellBg = "background:#FFC107;";
      }
      html += '<td style="padding:8px;text-align:center;border-bottom:1px solid #f5f5f5;' + cellColor + cellBg + '">' + (typeof row.cells[j] === 'string' ? row.cells[j] : escapeHtml(String(row.cells[j]))) + '</td>';
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
