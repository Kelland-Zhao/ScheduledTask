// V20260704.2 — 故障报告综合日报（Per-Process）
// 功能：合并 07（待判断+处理中）、21（超期未上传）、23（跟进项）为每工序一封综合日报
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

    // Failure_Database 只读一次，同时产出 Section B + C + reportNo→process 映射
    var dbData = _ur_readFailureDatabaseRaw();
    var allReports = dbData.reports;
    var pendingReviews = dbData.pendingReviews;
    console.log('Failure_Database 全量: ' + allReports.length + ' 条, 处理中: ' + pendingReviews.length + ' 条');

    // 构建 reportNumber → process 映射（用于 Section D 工序归属）
    var reportNoToProcess = {};
    allReports.forEach(function(r) {
      if (r.reportNumber) {
        reportNoToProcess[r.reportNumber] = mapFailureProcessToUserID(r.process);
      }
    });

    // Section C: 全部未上传（不限超期天数）
    var unuploadedReports = allReports.filter(function(r) {
      return !isFailureReportUploaded(r).isUploaded;
    });
    // 按超期天数降序排列
    unuploadedReports.sort(function(a, b) { return b.overdueDays - a.overdueDays; });
    console.log('Section C 未上传: ' + unuploadedReports.length + ' 条');

    var followUpData = _ur_getFollowUpData();
    console.log('Section D 跟进项: ' + followUpData.length + ' 条');

    // 2. 读取 userID 映射
    var userMaps = getUserIDLookupMaps();
    if (Object.keys(userMaps.nameToUser).length === 0) {
      console.warn('⚠️ userID 权限表无有效数据，跳过执行');
      writeLog(fnName, '警告', 'userID 权限表无有效数据，跳过执行', trigger, '');
      return;
    }

    // 3. 按工序归类
    var processMap = _ur_buildProcessPayloads(
      faultItems, pendingReviews, unuploadedReports, followUpData, reportNoToProcess
    );
    var processesWithData = Object.keys(processMap);
    console.log('工序归类完成: [' + processesWithData.join(', ') + ']');

    if (processesWithData.length === 0) {
      console.log('所有工序均无待处理项');
      writeLog(fnName, '成功', '所有工序均无待处理项，跳过执行', trigger, '');
      return;
    }

    // 4. 按工序发送
    var sentCount = _ur_sendProcessEmails(processMap, userMaps, trigger);
    console.log('=== 故障报告综合日报执行完成，发送 ' + sentCount + ' 封 ===');
    writeLog(fnName, '成功', '发送 ' + sentCount + ' 封，工序: [' + processesWithData.join(', ') + ']', trigger, '');

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

/** 读取 Failure_Database 一次，同时返回全量报告和处理中报告 */
function _ur_readFailureDatabaseRaw() {
  try {
    var sheet = SpreadsheetApp.openById(FR_CONFIG.SPREADSHEET_ID)
      .getSheetByName(FR_CONFIG.FAILURE_SHEET_NAME);
    if (!sheet) return { reports: [], pendingReviews: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { reports: [], pendingReviews: [] };
    var headers = data[0];
    var idx = getFieldIndexes(headers);
    var idxReviewStatus = headers.indexOf('审核状态 / Review Status');
    var idxReviewer = headers.indexOf('审核人 / Reviewed By');

    var reports = [];
    var pendingReviews = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var report = {
        id: row[idx['编号']] || '',
        machineId: row[idx['机台号']] || '',
        description: row[idx['问题描述']] || '',
        submitDate: row[idx['提交日期']] || '',
        workshop: row[idx['车间']] || '',
        process: row[idx['工序']] || '',
        reportNumber: row[idx['故障报告编号']] || '',
        assignDate: row[idx['分配日期']] || '',
        uploadDate: row[idx['上传日期']] || '',
        attachment: row[idx['附件']] || '',
        owner: String(row[idx['责任人']] || '').trim(),
        overdueDays: calculateOverdueDays(row[idx['分配日期']])
      };
      reports.push(report);

      // 同时判断是否处理中
      var reviewStatus = idxReviewStatus >= 0 ? String(row[idxReviewStatus] || '').trim() : '';
      if (reviewStatus && reviewStatus !== '已完成') {
        pendingReviews.push({
          id: report.id,
          machineId: report.machineId,
          description: report.description,
          process: mapFailureProcessToUserID(report.process),
          reportNumber: report.reportNumber,
          owner: report.owner,
          reviewer: idxReviewer >= 0 ? String(row[idxReviewer] || '').trim() : '',
          reviewStatus: reviewStatus
        });
      }
    }

    console.log('Failure_Database: ' + reports.length + ' 条全量, ' + pendingReviews.length + ' 条处理中');
    return { reports: reports, pendingReviews: pendingReviews };
  } catch (e) {
    console.error('读取 Failure_Database 失败: ' + e.message);
    return { reports: [], pendingReviews: [] };
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

// ========== 工序归类引擎 ==========

/** 判断 process payload 是否有任何待处理项 */
function _ur_hasAnyItems(payload) {
  return payload.sections.A.length > 0 ||
         payload.sections.B.length > 0 ||
         payload.sections.C.length > 0 ||
         payload.sections.D_owner.dueSoon.length > 0 ||
         payload.sections.D_owner.overdue.length > 0 ||
         payload.sections.D_verifier.length > 0;
}

/** 创建空的工序 payload */
function _ur_emptyProcessPayload(proc) {
  return {
    process: proc,
    displayName: PROCESS_DISPLAY_NAMES[proc] || proc,
    sections: { A: [], B: [], C: [], D_owner: { dueSoon: [], overdue: [] }, D_verifier: [] }
  };
}

/**
 * 核心函数：按工序归类所有待处理项
 * 返回 { 'INJ': ProcessPayload, 'TF': ProcessPayload, 'PK': ProcessPayload }
 */
function _ur_buildProcessPayloads(faultItems, pendingReviews, unuploadedReports, followUpData, reportNoToProcess) {
  var processes = ['INJ', 'TF', 'PK'];
  var processMap = {};
  processes.forEach(function(proc) {
    processMap[proc] = _ur_emptyProcessPayload(proc);
  });

  // Section A: 按 item.processType 分组
  faultItems.forEach(function(item) {
    var proc = item.processType;
    if (processMap[proc]) processMap[proc].sections.A.push(item);
  });

  // Section B: 按 review.process 分组（已是 INJ/TF/PK）
  pendingReviews.forEach(function(review) {
    var proc = review.process;
    if (processMap[proc]) processMap[proc].sections.B.push(review);
  });

  // Section C: 按 report.process 分组（需 IM→INJ 映射）
  unuploadedReports.forEach(function(report) {
    var proc = mapFailureProcessToUserID(report.process);
    if (processMap[proc]) processMap[proc].sections.C.push(report);
  });

  // Section D: 按 reportNo → process 解析工序
  followUpData.forEach(function(item) {
    var status = item.status || '';
    if (status.indexOf('已通过') !== -1) return;
    var proc = reportNoToProcess[item.reportNo];
    if (!proc || !processMap[proc]) return;
    var days = _ur_calcDaysUntilDue(item.dueDate);
    if (status.indexOf('未通过') !== -1 || status.indexOf('NA') !== -1) {
      if (days < 0) processMap[proc].sections.D_owner.overdue.push(item);
      else if (days <= 2) processMap[proc].sections.D_owner.dueSoon.push(item);
    }
    if (status.indexOf('未验证') !== -1) {
      processMap[proc].sections.D_verifier.push(item);
    }
  });

  // 过滤无事项的工序
  var result = {};
  for (var p in processMap) {
    if (processMap.hasOwnProperty(p) && _ur_hasAnyItems(processMap[p])) {
      result[p] = processMap[p];
    }
  }
  return result;
}

// ========== HTML 邮件生成 ==========

/** 动态主题行（Per-Process） */
function _ur_generateSubject(payload, isTest) {
  var date = formatVariableAsDate(new Date());
  var hasOverdue = payload.sections.C.some(function(r) { return r.overdueDays >= 7; }) ||
                   payload.sections.D_owner.overdue.length > 0;
  var prefix = isTest ? '【测试】' : '';
  var displayName = payload.displayName;
  var proc = payload.process;
  if (hasOverdue) {
    return prefix + '【故障报告综合日报】' + date + ' - ' + displayName + '(' + proc + ') 含逾期项 / Fault Report Daily Summary - ' + proc + ' (Overdue)';
  }
  return prefix + '【故障报告综合日报】' + date + ' - ' + displayName + '工序 / Fault Report Daily Summary - ' + proc + ' Process';
}

/** 完整 HTML 邮件正文 */
function _ur_generateBody(payload) {
  var date = formatVariableAsDate(new Date());
  var displayName = payload.displayName;
  var proc = payload.process;

  // 主体颜色主题：根据是否有逾期决定（Section C ≥7天 或 Section D 逾期项）
  var hasOverdue = payload.sections.C.some(function(r) { return r.overdueDays >= 7; }) ||
                   payload.sections.D_owner.overdue.length > 0;

  // ===== 头部（Mode A） =====
  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">';

  html += '<div style="background:#E60012;color:white;padding:16px 24px;">';
  html += '<h2 style="margin:0;font-size:20px;">【故障报告综合日报】Fault Report Daily Summary</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px;">' + displayName + '工序 / ' + proc + ' Process | ' + date + '</p>';
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

  // ===== Section C: 未上传报告 =====
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
    { label: '未上传<br><span style="font-size:10px;opacity:0.8;">Unuploaded</span>', value: payload.sections.C.length, color: payload.sections.C.length > 0 ? '#e74c3c' : '#2c3e50' },
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
    '未上传天数<br><span style="font-size:0.8em;opacity:0.9;">Days</span>'
  ];

  var rows = display.map(function(item) {
    var isOverdue = item.overdueDays >= 7;
    return [
      _ur_escapeHtml(item.id || '-'),
      _ur_escapeHtml(item.machineId || '-'),
      '<span style="word-wrap:break-word;">' + _ur_escapeHtml(item.description || '-') + '</span>',
      _ur_escapeHtml(extractNameFromPersonField(item.owner) || '-'),
      formatVariableAsDate(item.assignDate) || '-',
      _ur_buildBadge(item.overdueDays, isOverdue)
    ];
  });

  // 有超期项用红色主题，否则用中性色
  var hasOverdue = items.some(function(r) { return r.overdueDays >= 7; });
  var titleColor = hasOverdue ? '#d32f2f' : '#e67e22';
  var tableGradient = hasOverdue ? 'linear-gradient(135deg,#e74c3c,#c0392b)' : 'linear-gradient(135deg,#f39c12,#e67e22)';
  var rowBg = hasOverdue ? '#fff5f5' : '#fffbf0';

  var html = _ur_buildSectionTitle('三', '未上传报告', 'Unuploaded Reports', items.length, titleColor);
  html += _ur_buildTable(headers, rows, tableGradient, rowBg);
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

/** 计算工序邮件的 TO 和 CC 列表 */
function _ur_getProcessRecipients(payload, nameToUser, processToAdmins) {
  var proc = payload.process;
  var toSet = {};
  var ccSet = {};
  ccSet[_UR_CONFIG.ADMIN_EMAIL] = true;

  // TO: 工序管理员
  var admins = processToAdmins[proc] || [];
  admins.forEach(function(email) { toSet[email] = true; });

  // TO: Section C 报告责任人
  payload.sections.C.forEach(function(report) {
    var email = extractEmailFromPersonField(report.owner);
    var name = extractNameFromPersonField(report.owner);
    if (email) toSet[email] = true;
    if (name) {
      var user = nameToUser[name];
      if (user && user.lineManager) ccSet[user.lineManager] = true;
    }
  });

  // TO: Section D 责任人 + CC: Line Manager
  var allFollowUpItems = payload.sections.D_owner.dueSoon.concat(payload.sections.D_owner.overdue);
  allFollowUpItems.forEach(function(item) {
    var email = extractEmailFromPersonField(item.owner);
    var name = extractNameFromPersonField(item.owner);
    if (email) toSet[email] = true;
    if (name) {
      var user = nameToUser[name];
      if (user && user.lineManager) ccSet[user.lineManager] = true;
    }
  });

  // TO: Section D 验证人
  payload.sections.D_verifier.forEach(function(item) {
    var email = extractEmailFromPersonField(item.verifier);
    if (email) toSet[email] = true;
  });

  // CC: 工序管理员（已在 TO 中的跳过）
  admins.forEach(function(email) {
    if (!toSet[email]) ccSet[email] = true;
  });

  return { to: Object.keys(toSet), cc: Object.keys(ccSet) };
}

/** 按工序发送邮件 */
function _ur_sendProcessEmails(processMap, userMaps, trigger) {
  var sentCount = 0;
  var nameToUser = userMaps.nameToUser || {};
  var processToAdmins = userMaps.processToAdmins || {};

  for (var proc in processMap) {
    if (!processMap.hasOwnProperty(proc)) continue;
    var payload = processMap[proc];
    if (!_ur_hasAnyItems(payload)) continue;

    try {
      var recipients = _ur_getProcessRecipients(payload, nameToUser, processToAdmins);
      if (recipients.to.length === 0) {
        console.warn('⚠️ 工序 ' + proc + ': 无有效收件人，跳过');
        continue;
      }

      var isTest = _UR_TEST_MODE;
      var to = isTest ? _UR_TEST_EMAIL : recipients.to.join(',');
      var subject = _ur_generateSubject(payload, isTest);
      var body = _ur_generateBody(payload);
      var ccList = isTest ? [] : recipients.cc;

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

      console.log('✅ ' + proc + ' 综合日报已发送 → TO:' + recipients.to.length + '人 CC:' + (ccList.length) + '人 (' + sectionsDesc.join(', ') + ')');
      if (isTest) {
        console.log('   [测试模式] 原收件人: ' + recipients.to.join(', '));
      }

    } catch (emailError) {
      console.error('发送工序 ' + proc + ' 综合日报时出错:', emailError);
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
