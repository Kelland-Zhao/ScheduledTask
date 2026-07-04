// V20260704.1 — 故障报告综合日报
// 功能：合并 07（待判断+处理中）、21（超期未上传）、23（跟进项）为每人每天一封综合日报
// 依赖：00-专项菜单.js（PERMISSION_SPREADSHEET_ID 等全局常量）
//       01-Common.js（writeLog, formatVariableAsDate）
//       22-故障报告周报.js（FR_CONFIG, getUserIDLookupMaps, extractNameFromPersonField 等共享函数）
// 规范：遵循 docs/邮件UI规范.md — 内联样式、Mode A 头部、双语尾部

// ========== 本地配置 ==========

var _UR_TEST_MODE = true;
var _UR_TEST_EMAIL = 'kelland_zhao@colpal.com';

var _UR_CONFIG = {
  // 数据源：交接班模块（复用 07 的 FAULT_CONFIG.SHEET_ID）
  SHIFT_SPREADSHEET_ID: '10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w',
  SHIFT_SHEETS: ['Shift_Records'],

  // 工序时间阈值（分钟）— 与 07 保持一致
  PROCESS_THRESHOLDS: { 'INJ': 240, 'TF': 120, 'PK': 60 },

  // 通用
  ADMIN_EMAIL: 'kelland_zhao@colpal.com',
  MAX_ITEMS_PER_SECTION: 20,
  SENDER_NAME: '故障报告提醒系统'
};

// ========== 入口 ==========

/** 主函数：故障报告综合日报（定时/手动） */
function sendUnifiedFaultReportDaily(e) {
  var fnName = 'sendUnifiedFaultReportDaily';
  var trigger = e ? '定时' : '手动';
  try {
    console.log('=== 开始执行故障报告综合日报 ===');

    // 1. 读取全部数据源
    var faultItems = _ur_getFilteredFaultItems();
    console.log('Section A 待判断故障: ' + faultItems.length + ' 条');

    var pendingReviews = _ur_getPendingReviews();
    console.log('Section B 处理中报告: ' + pendingReviews.length + ' 条');

    var overdueReports = getOverdueFailureReportData(7);
    console.log('Section C 超期未上传: ' + overdueReports.length + ' 条');

    var followUpData = _ur_getFollowUpData();
    console.log('Section D 跟进项: ' + followUpData.length + ' 条');

    // 2. 读取 userID 映射
    var userMaps = getUserIDLookupMaps();
    if (Object.keys(userMaps.nameToUser).length === 0) {
      console.warn('⚠️ userID 权限表无有效数据，跳过执行');
      writeLog(fnName, '警告', 'userID 权限表无有效数据，跳过执行', trigger, '');
      return;
    }

    // 3. 角色解析：按 email 归类
    var personMap = _ur_buildPersonPayloads(
      faultItems, pendingReviews, overdueReports, followUpData, userMaps
    );
    var personCount = Object.keys(personMap).length;
    console.log('角色解析完成: ' + personCount + ' 人需要发送');

    if (personCount === 0) {
      console.log('没有需要发送的收件人');
      writeLog(fnName, '成功', '没有需要发送的收件人，跳过执行', trigger, '');
      return;
    }

    // 4. 发送
    var sentCount = _ur_sendAll(personMap, userMaps, trigger);
    console.log('=== 故障报告综合日报执行完成，发送 ' + sentCount + ' 封 ===');
    writeLog(fnName, '成功', '发送 ' + sentCount + ' 封，覆盖 ' + personCount + ' 人', trigger, '');

  } catch (error) {
    console.error('故障报告综合日报执行出错:', error);
    writeLog(fnName, '失败', error.message, trigger, error.stack || '');
  }
}

// ========== 数据读取 ==========

/** 解析 "X小时Y分钟" 中文格式的维修时间（移植自 07 的 _parseRepairTime） */
function _ur_parseRepairTime(raw) {
  if (raw == null || raw === '') return 0;
  var str = String(raw);
  var num = Number(str);
  if (!isNaN(num)) return num;
  var total = 0;
  var hourMatch = str.match(/(\d+)\s*小时/);
  if (hourMatch) total += parseInt(hourMatch[1], 10) * 60;
  var minMatch = str.match(/(\d+)\s*分钟/);
  if (minMatch) total += parseInt(minMatch[1], 10);
  if (total === 0) {
    var digits = str.match(/(\d+)/g);
    if (digits) total = parseInt(digits[0], 10);
  }
  return total;
}

/** 将 Date/字符串转为 YYYY-MM-DD 用于字符串比较 */
function _ur_formatDateStr(raw) {
  if (!raw) return '';
  var d = raw instanceof Date ? raw : new Date(raw);
  if (isNaN(d.getTime())) return '';
  var yyyy = d.getFullYear();
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

/** 读取 Shift_Records 并筛选符合条件的故障条目（移植自 07） */
function _ur_getFilteredFaultItems() {
  var faultItems = [];
  try {
    var spreadsheet = SpreadsheetApp.openById(_UR_CONFIG.SHIFT_SPREADSHEET_ID);
    _UR_CONFIG.SHIFT_SHEETS.forEach(function(sheetName) {
      try {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) return;
        var data = sheet.getDataRange().getValues();
        if (data.length <= 1) return;
        var headers = data[0];

        // 检查必要列是否存在
        var requiredFields = ['编号', '工序', '维修时间', '是否需要填写故障报告', '状态'];
        var hasAll = requiredFields.every(function(f) { return headers.indexOf(f) !== -1; });
        if (!hasAll) return;

        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          var item = {
            id: row[headers.indexOf('编号')] || '',
            machineNo: row[headers.indexOf('机台号')] || '',
            problemDesc: row[headers.indexOf('问题描述')] || '',
            process: row[headers.indexOf('处理过程')] || '',
            repairTime: row[headers.indexOf('维修时间')] || 0,
            repairPerson: row[headers.indexOf('维修人')] || '',
            submitDate: row[headers.indexOf('提交日期')] || '',
            workshop: row[headers.indexOf('车间')] || '',
            status: row[headers.indexOf('状态')] || '',
            processType: row[headers.indexOf('工序')] || '',
            needFaultReport: row[headers.indexOf('是否需要填写故障报告')] || '',
            sourceSheet: sheetName
          };
          faultItems.push(item);
        }
      } catch (sheetErr) {
        console.error('处理 Shift_Records 表出错:', sheetErr);
      }
    });
  } catch (error) {
    console.error('读取 Shift_Records 出错:', error);
    return [];
  }
  // 应用筛选逻辑
  return _ur_filterFaultItems(faultItems);
}

/** 筛选符合条件的故障条目（移植自 07 的 filterQualifiedFaultItems） */
function _ur_filterFaultItems(faultItems) {
  return faultItems.filter(function(item) {
    try {
      // 工序必须在阈值表中
      if (!(item.processType in _UR_CONFIG.PROCESS_THRESHOLDS)) return false;

      // 维修时间超过阈值
      var threshold = _UR_CONFIG.PROCESS_THRESHOLDS[item.processType];
      var repairTime = _ur_parseRepairTime(item.repairTime);
      if (repairTime < threshold) return false;

      // 尚未标记为需要填写故障报告
      if (item.needFaultReport !== '' && item.needFaultReport != null) return false;

      // 只提醒已解决的条目
      var statusStr = String(item.status);
      if (statusStr.indexOf('已解决') === -1 && statusStr.indexOf('Solved') === -1) return false;

      // PK / TF 工序专属规则（对齐 EDS FailureReport_Manage）
      if (item.processType === 'PK' || item.processType === 'TF') {
        if (item.submitDate) {
          var dateStr = _ur_formatDateStr(item.submitDate);
          var cutoff = item.processType === 'TF' ? '2026-06-01' : '2026-05-15';
          if (dateStr < cutoff) return false;
        }
        var problemDesc = String(item.problemDesc || '');
        if (problemDesc.indexOf('转规格') !== -1) return false;
      }

      return true;
    } catch (error) {
      console.error('筛选故障条目 ' + item.id + ' 时出错:', error);
      return false;
    }
  });
}

/** 读取 Failure_Database 中处理中的故障报告（移植自 07 的 _getPendingReviewReports） */
function _ur_getPendingReviews() {
  try {
    var sheet = SpreadsheetApp.openById(FR_CONFIG.SPREADSHEET_ID)
      .getSheetByName(FR_CONFIG.FAILURE_SHEET_NAME);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var idxReviewStatus = headers.indexOf('审核状态 / Review Status');
    if (idxReviewStatus < 0) return [];
    var reports = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var reviewStatus = String(row[idxReviewStatus] || '').trim();
      if (!reviewStatus || reviewStatus === '已完成') continue;
      reports.push({
        id: row[headers.indexOf('编号')] || '',
        machineId: row[headers.indexOf('机台号')] || '',
        description: row[headers.indexOf('问题描述')] || '',
        process: mapFailureProcessToUserID(row[headers.indexOf('工序')] || ''),
        reportNumber: row[headers.indexOf('故障报告编号')] || '',
        owner: String(row[headers.indexOf('责任人')] || '').trim(),
        reviewer: String(row[headers.indexOf('审核人 / Reviewed By')] || '').trim(),
        reviewStatus: reviewStatus
      });
    }
    console.log('处理中故障报告（非已完成）: ' + reports.length + ' 条');
    return reports;
  } catch (e) {
    console.error('读取待审核报告失败: ' + e.message);
    return [];
  }
}

/** 读取 Failure_Report_followup 跟进项数据（移植自 23 的 getFollowUpData） */
function _ur_getFollowUpData() {
  try {
    var sheet = SpreadsheetApp.openById(FR_CONFIG.SPREADSHEET_ID)
      .getSheetByName(FR_CONFIG.FOLLOWUP_SHEET_NAME);
    if (!sheet) { console.error('跟进项目表未找到'); return []; }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var idx = getFieldIndexes(data[0]);
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[idx['序号 / followup_id']]) continue;
      result.push({
        id:          row[idx['序号 / followup_id']]               || '',
        reportNo:    row[idx['故障报告编号 / failure_report_no']]  || '',
        paType:      row[idx['行动类型 / pa_type']]                || '',
        paPlan:      row[idx['预防行动 / pa_plan']]                || '',
        owner:  String(row[idx['责任人 / pa_who']]                || ''),
        dueDate:     row[idx['完成时间 / pa_when']]                || '',
        verifier: String(row[idx['验证人 / pa_verifier']]          || ''),
        status:      row[idx['状态 / status']]                     || '',
        notes:       row[idx['跟进内容 / follow_up_notes']]        || ''
      });
    }
    console.log('跟进项目数据: ' + result.length + ' 条');
    return result;
  } catch (error) {
    console.error('读取跟进项目数据时出错:', error);
    return [];
  }
}

/** 距离截止日期的天数（负=逾期） */
function _ur_calcDaysUntilDue(dueDate) {
  try {
    var due = (dueDate instanceof Date) ? new Date(dueDate) : new Date(dueDate);
    if (isNaN(due.getTime())) return 999;
    due.setHours(0, 0, 0, 0);
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.round((due - today) / (1000 * 3600 * 24));
  } catch (e) {
    return 999;
  }
}

/** 格式化跟进项日期 */
function _ur_formatFollowUpDate(value) {
  if (!value) return '';
  var d = (value instanceof Date) ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : formatVariableAsDate(d);
}

// ========== 角色解析引擎 ==========

/** 创建空的 person payload */
function _ur_emptyPayload(email) {
  return {
    email: email,
    name: '',
    roles: { isProcessAdmin: false, isReportOwner: false, isFollowUpOwner: false, isFollowUpVerifier: false, processes: [] },
    sections: { A: [], B: [], C: [], D_owner: { dueSoon: [], overdue: [] }, D_verifier: [] }
  };
}

/** 判断 person payload 是否有任何待处理项 */
function _ur_hasAnyItems(payload) {
  return payload.sections.A.length > 0 ||
         payload.sections.B.length > 0 ||
         payload.sections.C.length > 0 ||
         payload.sections.D_owner.dueSoon.length > 0 ||
         payload.sections.D_owner.overdue.length > 0 ||
         payload.sections.D_verifier.length > 0;
}

/**
 * 核心函数：按 email 归类所有待处理项
 * 返回 { email: PersonPayload }
 */
function _ur_buildPersonPayloads(faultItems, pendingReviews, overdueReports, followUpData, userMaps) {
  var personMap = {};
  var nameToUser = userMaps.nameToUser || {};
  var processToAdmins = userMaps.processToAdmins || {};
  var nameToProcess = userMaps.nameToProcess || {};

  function ensure(email) {
    var key = String(email).toLowerCase().trim();
    if (!personMap[key]) personMap[key] = _ur_emptyPayload(email);
    return personMap[key];
  }

  function addProcessRole(email, process) {
    var p = ensure(email);
    p.roles.isProcessAdmin = true;
    if (p.roles.processes.indexOf(process) === -1) p.roles.processes.push(process);
    return p;
  }

  // Section A: 故障待判断 → 工序管理员
  faultItems.forEach(function(item) {
    var process = item.processType;
    var admins = processToAdmins[process] || [];
    admins.forEach(function(adminEmail) {
      addProcessRole(adminEmail, process).sections.A.push(item);
    });
  });

  // Section B: 处理中报告 → 工序管理员
  pendingReviews.forEach(function(review) {
    var process = review.process;
    if (!process) return;
    var admins = processToAdmins[process] || [];
    admins.forEach(function(adminEmail) {
      addProcessRole(adminEmail, process).sections.B.push(review);
    });
  });

  // 构建 email→name 反向映射
  var emailToName = {};
  for (var n in nameToUser) {
    if (nameToUser.hasOwnProperty(n)) {
      var u = nameToUser[n];
      if (u.email) emailToName[u.email.toLowerCase()] = n;
    }
  }

  // Section C: 超期未上传报告 → 报告责任人
  overdueReports.forEach(function(report) {
    var email = extractEmailFromPersonField(report.owner);
    var name = extractNameFromPersonField(report.owner);
    if (!email && name) {
      var user = nameToUser[name];
      if (user && user.email) email = user.email;
    }
    if (!email) return;
    var p = ensure(email);
    p.roles.isReportOwner = true;
    p.name = p.name || name || emailToName[email.toLowerCase()] || '';
    p.sections.C.push(report);
  });

  // Section D: 跟进项
  followUpData.forEach(function(item) {
    var status = item.status || '';
    if (status.indexOf('已通过') !== -1) return;

    // D1: 责任人（未通过/NA，且临期≤2天或逾期）
    if (status.indexOf('未通过') !== -1 || status.indexOf('NA') !== -1) {
      var ownerEmail = extractEmailFromPersonField(item.owner);
      if (ownerEmail) {
        var p = ensure(ownerEmail);
        p.roles.isFollowUpOwner = true;
        p.name = p.name || extractNameFromPersonField(item.owner) ||
                 emailToName[ownerEmail.toLowerCase()] || '';
        var days = _ur_calcDaysUntilDue(item.dueDate);
        if (days < 0) {
          p.sections.D_owner.overdue.push(item);
        } else if (days <= 2) {
          p.sections.D_owner.dueSoon.push(item);
        }
      }
    }

    // D2: 验证人（未验证）
    if (status.indexOf('未验证') !== -1) {
      var verifierEmail = extractEmailFromPersonField(item.verifier);
      if (verifierEmail) {
        var p = ensure(verifierEmail);
        p.roles.isFollowUpVerifier = true;
        p.name = p.name || extractNameFromPersonField(item.verifier) ||
                 emailToName[verifierEmail.toLowerCase()] || '';
        p.sections.D_verifier.push(item);
      }
    }
  });

  // 过滤无事项的人
  var result = {};
  for (var email in personMap) {
    if (personMap.hasOwnProperty(email) && _ur_hasAnyItems(personMap[email])) {
      result[email] = personMap[email];
    }
  }
  return result;
}

// ========== HTML 邮件生成 ==========

/** 动态主题行 */
function _ur_generateSubject(payload, isTest) {
  var date = formatVariableAsDate(new Date());
  var totalItems = payload.sections.A.length + payload.sections.B.length +
                   payload.sections.C.length +
                   payload.sections.D_owner.dueSoon.length +
                   payload.sections.D_owner.overdue.length +
                   payload.sections.D_verifier.length;
  var hasOverdue = payload.sections.C.length > 0 || payload.sections.D_owner.overdue.length > 0;
  var prefix = isTest ? '【测试】' : '';
  if (hasOverdue) {
    return prefix + '【故障报告综合日报】' + date + ' - 含逾期项 / Fault Report Daily Summary - Overdue Items';
  }
  return prefix + '【故障报告综合日报】' + date + ' - ' + totalItems + '项待处理 / Unified Fault Report Daily Summary';
}

/** 完整 HTML 邮件正文 */
function _ur_generateBody(payload) {
  var date = formatVariableAsDate(new Date());
  var name = payload.name || payload.email;

  // 主体颜色主题：根据是否有逾期决定
  var hasOverdue = payload.sections.C.length > 0 || payload.sections.D_owner.overdue.length > 0;

  // ===== 头部（Mode A） =====
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">';

  html += '<div style="background:#E60012;color:white;padding:16px 24px;">';
  html += '<h2 style="margin:0;font-size:20px;">【故障报告综合日报】Fault Report Daily Summary</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px;">收件人: ' + _ur_escapeHtml(name) + ' | ' + date + '</p>';
  if (hasOverdue) {
    html += '<p style="margin:4px 0 0;opacity:0.85;font-size:13px;color:#FFC107;">⚠ 包含逾期项 / Overdue Items Included</p>';
  }
  html += '</div>';

  html += '<div style="padding:24px;">';

  // ===== 摘要卡片 =====
  html += _ur_buildSummaryCards(payload);

  // ===== Section A: 故障待判断 =====
  if (payload.sections.A.length > 0) {
    html += _ur_buildSectionA(payload.sections.A);
  }

  // ===== Section B: 处理中故障报告 =====
  if (payload.sections.B.length > 0) {
    html += _ur_buildSectionB(payload.sections.B);
  }

  // ===== Section C: 超期未上传报告 =====
  if (payload.sections.C.length > 0) {
    html += _ur_buildSectionC(payload.sections.C);
  }

  // ===== Section D: 跟进项 =====
  var dOwnerItems = payload.sections.D_owner.dueSoon.concat(payload.sections.D_owner.overdue);
  var dVerifierItems = payload.sections.D_verifier;
  if (dOwnerItems.length > 0 || dVerifierItems.length > 0) {
    html += _ur_buildSectionD(payload.sections.D_owner, payload.sections.D_verifier);
  }

  html += '</div>';

  // ===== 尾部（Mode B） =====
  html += _ur_buildFooter();

  html += '</div></body></html>';
  return html;
}

/** HTML 转义 */
function _ur_escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 摘要卡片（4 列 table 布局，兼容各邮件客户端） */
function _ur_buildSummaryCards(payload) {
  var cards = [
    { label: '故障待判断<br><span style="font-size:10px;opacity:0.8;">Faults Pending</span>', value: payload.sections.A.length, color: '#E60012' },
    { label: '处理中报告<br><span style="font-size:10px;opacity:0.8;">In Review</span>', value: payload.sections.B.length, color: '#f39c12' },
    { label: '超期未上传<br><span style="font-size:10px;opacity:0.8;">Overdue</span>', value: payload.sections.C.length, color: payload.sections.C.length > 0 ? '#e74c3c' : '#2c3e50' },
    { label: '跟进项<br><span style="font-size:10px;opacity:0.8;">Follow-ups</span>',
      value: payload.sections.D_owner.dueSoon.length + payload.sections.D_owner.overdue.length + payload.sections.D_verifier.length,
      color: payload.sections.D_owner.overdue.length > 0 ? '#e74c3c' : '#2c3e50' }
  ];

  var html = '<table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin-bottom:24px;"><tr>';
  cards.forEach(function(card) {
    html += '<td style="text-align:center;padding:14px 8px;border:1px solid #ecf0f1;width:25%;">';
    html += '<div style="font-size:28px;font-weight:bold;color:' + card.color + ';">' + card.value + '</div>';
    html += '<div style="color:#7f8c8d;font-size:12px;margin-top:4px;line-height:1.4;">' + card.label + '</div>';
    html += '</td>';
  });
  html += '</tr></table>';
  return html;
}

/** 通用表格构建器 */
function _ur_buildTable(headers, rows, headerGradient, rowAltBg) {
  var hGrad = headerGradient || 'linear-gradient(135deg,#E60012,#C00010)';
  var altBg = rowAltBg || '#f8f9fa';
  var html = '<div style="overflow-x:auto;">';
  html += '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">';
  // 表头
  html += '<thead><tr style="background:' + hGrad + ';color:white;">';
  headers.forEach(function(h) {
    html += '<th style="padding:10px 8px;text-align:left;font-weight:600;font-size:13px;">' + h + '</th>';
  });
  html += '</tr></thead><tbody>';
  // 数据行
  rows.forEach(function(row, i) {
    var bg = i % 2 === 0 ? '#ffffff' : altBg;
    html += '<tr style="background:' + bg + ';">';
    row.forEach(function(cell) {
      html += '<td style="padding:10px 8px;border-bottom:1px solid #e9ecef;font-size:13px;color:#34495e;">' + cell + '</td>';
    });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

/** 超期/临期 Badge（遵循 UI 规范 §7.1） */
function _ur_buildBadge(days, isOverdue) {
  var grad, shadow;
  if (isOverdue) {
    grad = 'linear-gradient(135deg,#e74c3c,#c0392b)';
    shadow = '0 2px 6px rgba(231,76,60,0.3)';
    return '<div style="display:inline-block;text-align:center;">' +
      '<div style="background:' + grad + ';color:white;padding:6px 12px;border-radius:16px;font-weight:600;font-size:12px;box-shadow:' + shadow + ';display:inline-block;min-width:80px;">' +
      '<span style="display:block;">[逾期] ' + Math.abs(days) + '天</span>' +
      '<span style="display:block;font-size:10px;opacity:0.9;">Days Overdue</span></div></div>';
  }
  grad = 'linear-gradient(135deg,#f39c12,#e67e22)';
  shadow = '0 2px 6px rgba(243,156,18,0.3)';
  return '<div style="display:inline-block;text-align:center;">' +
    '<div style="background:' + grad + ';color:white;padding:6px 12px;border-radius:16px;font-weight:600;font-size:12px;box-shadow:' + shadow + ';display:inline-block;min-width:80px;">' +
    '<span style="display:block;">还剩 ' + days + '天</span>' +
    '<span style="display:block;font-size:10px;opacity:0.9;">Days Left</span></div></div>';
}

/** 审核状态 Badge */
function _ur_buildStatusBadge(status) {
  var colorMap = {
    '待提交': '#95a5a6',
    '主管审核中': '#f39c12',
    '已退回': '#e74c3c',
    '审核通过': '#27ae60'
  };
  var color = colorMap[status] || '#95a5a6';
  return '<span style="display:inline-block;background:' + color + ';color:white;padding:4px 8px;border-radius:12px;font-size:12px;">' + _ur_escapeHtml(status) + '</span>';
}

/** 分区标题 */
function _ur_buildSectionTitle(number, titleCn, titleEn, count, color) {
  var c = color || '#E60012';
  return '<h3 style="color:' + c + ';border-left:4px solid ' + c + ';padding-left:8px;' +
    'border-bottom:2px solid ' + c + ';padding-bottom:8px;margin-top:28px;margin-bottom:16px;font-size:17px;">' +
    number + '、' + titleCn + ' (' + count + '条)<br>' +
    '<span style="font-size:0.75em;font-weight:400;color:#7f8c8d;">' + titleEn + '</span></h3>';
}

/** 截断提示 */
function _ur_buildTruncationNote(total, shown) {
  if (total <= shown) return '';
  return '<p style="text-align:center;color:#e67e22;font-size:13px;margin-top:8px;">' +
    '（共 ' + total + ' 条，以上展示前 ' + shown + ' 条，完整清单请查看系统）</p>';
}

// ===== 各 Section HTML 构建 =====

function _ur_buildSectionA(items) {
  var maxShow = _UR_CONFIG.MAX_ITEMS_PER_SECTION;
  var display = items.slice(0, maxShow);
  var headers = [
    '编号<br><span style="font-size:0.8em;opacity:0.9;">ID</span>',
    '机台号<br><span style="font-size:0.8em;opacity:0.9;">Machine</span>',
    '问题描述<br><span style="font-size:0.8em;opacity:0.9;">Description</span>',
    '维修时间<br><span style="font-size:0.8em;opacity:0.9;">Repair Time</span>',
    '提交日期<br><span style="font-size:0.8em;opacity:0.9;">Submit Date</span>',
    '状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span>'
  ];
  var processLabels = { 'INJ': '注塑', 'TF': '植磨毛', 'PK': '包装' };

  var rows = display.map(function(item) {
    return [
      _ur_escapeHtml(item.id || '-'),
      _ur_escapeHtml(item.machineNo || '-'),
      '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.problemDesc || '-') + '</span>',
      '<span style="color:#E60012;font-weight:bold;">' + (item.repairTime || 0) + ' min</span>',
      _ur_formatDateStr(item.submitDate) || '-',
      '<span style="display:inline-block;background:#FF6B6B;color:white;padding:4px 8px;border-radius:12px;font-size:12px;">待判断</span>'
    ];
  });

  var html = _ur_buildSectionTitle('一', '故障待判断', 'Faults Pending Judgment', items.length, '#E60012');
  html += _ur_buildTable(headers, rows, 'linear-gradient(135deg,#E60012,#C00010)', '#f9f9f9');
  html += _ur_buildTruncationNote(items.length, maxShow);
  return html;
}

function _ur_buildSectionB(items) {
  var maxShow = _UR_CONFIG.MAX_ITEMS_PER_SECTION;
  var display = items.slice(0, maxShow);
  var headers = [
    '编号<br><span style="font-size:0.8em;opacity:0.9;">ID</span>',
    '机台号<br><span style="font-size:0.8em;opacity:0.9;">Machine</span>',
    '问题描述<br><span style="font-size:0.8em;opacity:0.9;">Description</span>',
    '责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span>',
    '故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span>',
    '审核状态<br><span style="font-size:0.8em;opacity:0.9;">Review Status</span>'
  ];

  var rows = display.map(function(item) {
    return [
      _ur_escapeHtml(item.id || '-'),
      _ur_escapeHtml(item.machineId || '-'),
      '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.description || '-') + '</span>',
      _ur_escapeHtml(extractNameFromPersonField(item.owner) || '-'),
      _ur_escapeHtml(item.reportNumber || '-'),
      _ur_buildStatusBadge(item.reviewStatus)
    ];
  });

  var html = _ur_buildSectionTitle('二', '处理中故障报告', 'Reports In Review', items.length, '#f39c12');
  html += _ur_buildTable(headers, rows, 'linear-gradient(135deg,#f39c12,#e67e22)', '#fffbf0');
  html += _ur_buildTruncationNote(items.length, maxShow);
  return html;
}

function _ur_buildSectionC(items) {
  var maxShow = _UR_CONFIG.MAX_ITEMS_PER_SECTION;
  var display = items.slice(0, maxShow);
  var headers = [
    '编号<br><span style="font-size:0.8em;opacity:0.9;">ID</span>',
    '机台号<br><span style="font-size:0.8em;opacity:0.9;">Machine</span>',
    '问题描述<br><span style="font-size:0.8em;opacity:0.9;">Description</span>',
    '责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span>',
    '分配日期<br><span style="font-size:0.8em;opacity:0.9;">Assign Date</span>',
    '超期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span>'
  ];

  var rows = display.map(function(item) {
    return [
      _ur_escapeHtml(item.id || '-'),
      _ur_escapeHtml(item.machineId || '-'),
      '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.description || '-') + '</span>',
      _ur_escapeHtml(extractNameFromPersonField(item.owner) || '-'),
      formatVariableAsDate(item.assignDate) || '-',
      _ur_buildBadge(item.overdueDays, true)
    ];
  });

  var html = _ur_buildSectionTitle('三', '超期未上传报告', 'Overdue Unuploaded Reports', items.length, '#d32f2f');
  html += _ur_buildTable(headers, rows, 'linear-gradient(135deg,#e74c3c,#c0392b)', '#fff5f5');
  html += _ur_buildTruncationNote(items.length, maxShow);
  return html;
}

function _ur_buildSectionD(dOwner, dVerifier) {
  var dueSoon = dOwner.dueSoon || [];
  var overdue = dOwner.overdue || [];
  var verifierItems = dVerifier || [];
  var totalOwner = dueSoon.length + overdue.length;
  var totalD = totalOwner + verifierItems.length;
  var maxShow = _UR_CONFIG.MAX_ITEMS_PER_SECTION;

  var html = _ur_buildSectionTitle('四', '跟进项', 'Follow-up Items', totalD, '#3f51b5');

  // D1: 逾期项
  if (overdue.length > 0) {
    var display = overdue.slice(0, maxShow);
    var ownerHeaders = [
      '故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span>',
      '预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span>',
      '期限<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span>',
      '状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span>',
      '逾期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span>'
    ];
    var rows = display.map(function(item) {
      return [
        _ur_escapeHtml(item.reportNo || '-'),
        '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.paPlan || '-') + '</span>',
        _ur_formatFollowUpDate(item.dueDate) || '-',
        _ur_escapeHtml(item.status || '-'),
        _ur_buildBadge(_ur_calcDaysUntilDue(item.dueDate), true)
      ];
    });
    html += '<h4 style="color:#d32f2f;margin:16px 0 8px;border-bottom:1px solid #f44336;padding-bottom:6px;">' +
      '【逾期】已逾期跟进项 Overdue Items (' + overdue.length + '条)</h4>';
    html += _ur_buildTable(ownerHeaders, rows, 'linear-gradient(135deg,#e74c3c,#c0392b)', '#fff5f5');
    html += _ur_buildTruncationNote(overdue.length, maxShow);
  }

  // D2: 临期项
  if (dueSoon.length > 0) {
    var display = dueSoon.slice(0, maxShow);
    var soonHeaders = [
      '故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span>',
      '预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span>',
      '期限<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span>',
      '状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span>',
      '剩余天数<br><span style="font-size:0.8em;opacity:0.9;">Days Left</span>'
    ];
    var rows = display.map(function(item) {
      return [
        _ur_escapeHtml(item.reportNo || '-'),
        '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.paPlan || '-') + '</span>',
        _ur_formatFollowUpDate(item.dueDate) || '-',
        _ur_escapeHtml(item.status || '-'),
        _ur_buildBadge(_ur_calcDaysUntilDue(item.dueDate), false)
      ];
    });
    html += '<h4 style="color:#e65100;margin:16px 0 8px;border-bottom:1px solid #f39c12;padding-bottom:6px;">' +
      '【临期】即将到期跟进项 Due Soon Items (' + dueSoon.length + '条)</h4>';
    html += _ur_buildTable(soonHeaders, rows, 'linear-gradient(135deg,#f39c12,#e67e22)', '#fffbf0');
    html += _ur_buildTruncationNote(dueSoon.length, maxShow);
  }

  // D3: 待验证项
  if (verifierItems.length > 0) {
    var display = verifierItems.slice(0, maxShow);
    var verifierHeaders = [
      '跟进编号<br><span style="font-size:0.8em;opacity:0.9;">Follow-up ID</span>',
      '故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span>',
      '预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span>',
      '责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span>',
      '期限<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span>',
      '状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span>'
    ];
    var rows = display.map(function(item) {
      return [
        _ur_escapeHtml(item.id || '-'),
        _ur_escapeHtml(item.reportNo || '-'),
        '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.paPlan || '-') + '</span>',
        _ur_escapeHtml(extractNameFromPersonField(item.owner) || '-'),
        _ur_formatFollowUpDate(item.dueDate) || '-',
        _ur_escapeHtml(item.status || '-')
      ];
    });
    html += '<h4 style="color:#283593;margin:16px 0 8px;border-bottom:1px solid #3f51b5;padding-bottom:6px;">' +
      '【待验证】待验证跟进项 Pending Verification (' + verifierItems.length + '条)</h4>';
    html += _ur_buildTable(verifierHeaders, rows, 'linear-gradient(135deg,#3f51b5,#283593)', '#f0f4ff');
    html += _ur_buildTruncationNote(verifierItems.length, maxShow);
  }

  return html;
}

/** 双语尾部（Mode B） */
function _ur_buildFooter() {
  return '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;margin-top:32px;">' +
    '<p style="margin-bottom:10px;">请及时处理以上事项。<br>' +
    '<span style="font-size:0.9em;opacity:0.8;">Please handle the above items promptly.</span></p>' +
    '<p style="margin:0;font-style:italic;">' +
    '此邮件由故障报告综合日报系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.7;">' +
    'This email is automatically sent by the system, please do not reply.</span></p></div>';
}

// ========== 发送 ==========

/** 计算 CC 列表 */
function _ur_computeCC(payload, nameToUser, processToAdmins) {
  var ccSet = {};
  ccSet[_UR_CONFIG.ADMIN_EMAIL] = true;

  // 工序管理员 CC（排除收件人本人）
  payload.roles.processes.forEach(function(process) {
    var admins = processToAdmins[process] || [];
    admins.forEach(function(email) {
      if (email.toLowerCase() !== payload.email.toLowerCase()) {
        ccSet[email] = true;
      }
    });
  });

  // Section C 报告责任人的 Line Manager
  payload.sections.C.forEach(function(report) {
    var name = extractNameFromPersonField(report.owner);
    if (name) {
      var user = nameToUser[name];
      if (user && user.lineManager) ccSet[user.lineManager] = true;
    }
  });

  // Section D1 跟进责任人的 Line Manager
  var allFollowUpItems = payload.sections.D_owner.dueSoon.concat(payload.sections.D_owner.overdue);
  allFollowUpItems.forEach(function(item) {
    var name = extractNameFromPersonField(item.owner);
    if (name) {
      var user = nameToUser[name];
      if (user && user.lineManager) ccSet[user.lineManager] = true;
    }
  });

  return Object.keys(ccSet);
}

/** 逐人发送邮件 */
function _ur_sendAll(personMap, userMaps, trigger) {
  var sentCount = 0;
  var nameToUser = userMaps.nameToUser || {};
  var processToAdmins = userMaps.processToAdmins || {};

  for (var email in personMap) {
    if (!personMap.hasOwnProperty(email)) continue;
    var payload = personMap[email];
    if (!_ur_hasAnyItems(payload)) continue;

    try {
      var isTest = _UR_TEST_MODE;
      var to = isTest ? _UR_TEST_EMAIL : payload.email;
      var subject = _ur_generateSubject(payload, isTest);
      var body = _ur_generateBody(payload);
      var ccList = _ur_computeCC(payload, nameToUser, processToAdmins);

      var options = {
        htmlBody: body,
        name: _UR_CONFIG.SENDER_NAME
      };

      if (ccList.length > 0) {
        options.cc = ccList.join(',');
      }

      GmailApp.sendEmail(to, subject, '请使用支持HTML的邮件客户端查看此邮件。', options);

      sentCount++;
      var sectionsDesc = [];
      if (payload.sections.A.length > 0) sectionsDesc.push('A:' + payload.sections.A.length);
      if (payload.sections.B.length > 0) sectionsDesc.push('B:' + payload.sections.B.length);
      if (payload.sections.C.length > 0) sectionsDesc.push('C:' + payload.sections.C.length);
      var dTotal = payload.sections.D_owner.dueSoon.length + payload.sections.D_owner.overdue.length + payload.sections.D_verifier.length;
      if (dTotal > 0) sectionsDesc.push('D:' + dTotal);

      console.log('✅ 综合日报已发送 → ' + to + ' (' + sectionsDesc.join(', ') + ') CC:' + ccList.join(', '));
      if (isTest) {
        console.log('   [测试模式] 原收件人应为: ' + email);
      }

    } catch (emailError) {
      console.error('发送综合日报给 ' + email + ' 时出错:', emailError);
    }
  }

  return sentCount;
}

// ========== 测试 ==========

/** 测试函数：发送至 kelland_zhao@colpal.com */
function testUnifiedFaultReportDaily() {
  _UR_TEST_MODE = true;
  sendUnifiedFaultReportDaily();
}
