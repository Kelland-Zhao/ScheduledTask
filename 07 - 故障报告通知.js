// V20260608.2 — 故障报告邮件通知模块
// 功能：自动检测交接班模块故障条目，对超时且未填写故障报告的条目发送邮件通知
// 变更：收件人从"通知清单"切换为 userID 表动态过滤

// ========== 故障报告配置 ==========
const FAULT_CONFIG = {
  // 数据源：交接班模块表格
  SHEET_ID: '10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w',

  // 工序车间数据表名称
  SHIFT_SHEETS: ['Shift_Records'],

  // 工序时间阈值（分钟）
  PROCESS_THRESHOLDS: {
    'INJ': 240,
    'TF': 120,
    'PK': 60
  },

  // userID 列索引（0-based）
  USERID_PROCESS_COL: 14,       // O列 — 工序（INJ/TF/PK 等）
  USERID_EMAIL_COL: 9,          // J列 — GMail（复用全局 PERMISSION_EMAIL_COL_IDX）
  USERID_FAULT_PERM_COL: 57,    // BF列 — 故障报告管理权限（Y=有权限）

  EMAIL_SUBJECT_PREFIX: '[故障报告提醒]',
  ADMIN_EMAIL: 'kelland_zhao@colpal.com',
  MAX_ITEMS_IN_EMAIL: 20
};

// 故障报告数据库（待审核数据源）
const FR_DATABASE_ID = '1YAPdZKVEOHgCGIJRQwWTQBmwaWIS4yd1SQKJJfRCtAU';
const FR_DATABASE_SHEET = 'Failure_Database';

// ========== 故障检测主函数 ==========

/** 主函数：故障检测和通知发送（定时 e 存在 / 手动 e 为 undefined） */
function mainFaultDetection(e) {
  const fnName = 'mainFaultDetection';
  const trigger = e ? '定时' : '手动';
  try {
    console.log('开始故障检测流程...');

    const faultItems = getAllFaultItems();
    console.log('获取到 ' + faultItems.length + ' 个故障条目');

    if (faultItems.length === 0) {
      writeLog(fnName, '成功', '无故障条目', trigger, '');
      return;
    }

    const qualifiedItems = filterQualifiedFaultItems(faultItems);
    console.log('筛选出 ' + qualifiedItems.length + ' 个符合通知条件的故障条目');

    if (qualifiedItems.length === 0) {
      writeLog(fnName, '成功', '无符合通知条件的条目（共扫描 ' + faultItems.length + ' 条）', trigger, '');
      return;
    }

    const notificationConfig = getNotificationConfig();
    console.log('获取到 ' + notificationConfig.length + ' 条通知配置');

    const sentCount = sendFaultNotifications(qualifiedItems, notificationConfig, trigger);
    console.log('成功发送 ' + sentCount + ' 封通知邮件');

    writeLog(fnName, '成功', '发送 ' + sentCount + ' 封通知，覆盖 ' + qualifiedItems.length + ' 个故障条目', trigger, '');

  } catch (error) {
    console.error('故障检测流程出错:', error);
    writeLog(fnName, '失败', error.message, trigger, error.stack || '');
    sendErrorNotification(error, trigger);
  }
}

// ========== 数据读取 ==========

function getAllFaultItems() {
  const faultItems = [];

  try {
    const spreadsheet = SpreadsheetApp.openById(FAULT_CONFIG.SHEET_ID);

    FAULT_CONFIG.SHIFT_SHEETS.forEach(function(sheetName) {
      try {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        for (let i = 1; i < data.length; i++) {
          const item = createFaultItem(data[i], headers, sheetName);
          if (item) faultItems.push(item);
        }
      } catch (sheetError) {
        console.error('处理表 ' + sheetName + ' 时出错:', sheetError);
      }
    });

  } catch (error) {
    console.error('获取故障条目时出错:', error);
    throw error;
  }

  return faultItems;
}

function createFaultItem(row, headers, sheetName) {
  try {
    const requiredFields = ['编号', '工序', '维修时间', '是否需要填写故障报告', '状态'];
    const hasRequiredFields = requiredFields.every(function(field) {
      return headers.indexOf(field) !== -1;
    });

    if (!hasRequiredFields) return null;

    function val(field) {
      const idx = headers.indexOf(field);
      return idx >= 0 ? row[idx] : '';
    }

    return {
      id: val('编号'),
      shift: val('班次'),
      machineNo: val('机台号'),
      problemDesc: val('问题描述'),
      process: val('处理过程'),
      status: val('状态'),
      repairPerson: val('维修人'),
      repairTime: val('维修时间') || 0,
      faultCode: val('故障代码'),
      workOrderNo: val('工单号'),
      handoverPerson: val('交接人'),
      submitDate: val('提交日期'),
      transferToMaintenance: val('是否转保养'),
      recorder: val('填写人'),
      workshop: val('车间'),
      processType: val('工序'),
      participants: val('参与人数'),
      machineModel: val('机型'),
      pmFollowUp: val('是否已安排PM跟进'),
      needFaultReport: val('是否需要填写故障报告'),
      followUp: val('是否跟随'),
      isFinal: val('判断是否最后'),
      sourceSheet: sheetName
    };

  } catch (error) {
    console.error('创建故障条目对象时出错:', error);
    return null;
  }
}

function extractWorkshop(sheetName) {
  if (sheetName.indexOf('TB1') !== -1) return 'TB1';
  if (sheetName.indexOf('TB2') !== -1) return 'TB2';
  return '';
}

// ========== 筛选逻辑 ==========

function filterQualifiedFaultItems(faultItems) {
  return faultItems.filter(function(item) {
    try {
      if (!(item.processType in FAULT_CONFIG.PROCESS_THRESHOLDS)) return false;

      const threshold = FAULT_CONFIG.PROCESS_THRESHOLDS[item.processType];
      const repairTime = Number(item.repairTime) || 0;
      if (repairTime < threshold) return false;

      if (item.needFaultReport !== '' && item.needFaultReport != null) return false;

      // 只提醒已解决的条目（修完了才需要判定是否写故障报告）
      var statusStr = String(item.status);
      if (statusStr.indexOf('已解决') === -1 && statusStr.indexOf('Solved') === -1) return false;

      // PK / TF 工序专属规则（对齐 EDS FailureReport_Manage）
      if (item.processType === 'PK' || item.processType === 'TF') {
        // 提交日期必须 ≥ 2026-05-15
        if (item.submitDate) {
          var submitDateObj = new Date(item.submitDate);
          if (submitDateObj < new Date('2026-05-15')) return false;
        }
        // 排除"转规格"问题
        var problemDesc = String(item.problemDesc || '');
        if (problemDesc.indexOf('转规格') !== -1) return false;
        // TF 额外排除处理过程中含"转规格"
        if (item.processType === 'TF') {
          var processDesc = String(item.process || '');
          if (processDesc.indexOf('转规格') !== -1) return false;
        }
      }

      return true;
    } catch (error) {
      console.error('筛选故障条目 ' + item.id + ' 时出错:', error);
      return false;
    }
  });
}

// ========== 通知配置读取 ==========

function getNotificationConfig() {
  try {
    const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);

    if (!sheet) {
      console.error('找不到 userID 表');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const configs = [];
    const targetProcesses = ['INJ', 'TF', 'PK'];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const process = String(row[FAULT_CONFIG.USERID_PROCESS_COL] || '').trim();
      const email = String(row[FAULT_CONFIG.USERID_EMAIL_COL] || '').trim();
      const hasPerm = String(row[FAULT_CONFIG.USERID_FAULT_PERM_COL] || '').trim();

      if (targetProcesses.indexOf(process) !== -1 && hasPerm === 'Y' && email) {
        configs.push({ process: process, mail: email });
      }
    }

    console.log('从 userID 过滤到 ' + configs.length + ' 条通知配置');
    return configs;

  } catch (error) {
    console.error('获取通知配置时出错:', error);
    throw error;
  }
}

// ========== 待审核故障报告 ==========

function _getPendingReviewReports() {
  try {
    var sheet = SpreadsheetApp.openById(FR_DATABASE_ID).getSheetByName(FR_DATABASE_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var idxReviewStatus = headers.indexOf('审核状态 / Review Status');
    if (idxReviewStatus < 0) return [];
    var reports = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[idxReviewStatus] || '').trim() !== '主管审核中') continue;
      reports.push({
        id: row[headers.indexOf('编号')] || '',
        machineId: row[headers.indexOf('机台号')] || '',
        description: row[headers.indexOf('问题描述')] || '',
        process: row[headers.indexOf('工序')] || '',
        reportNumber: row[headers.indexOf('故障报告编号')] || '',
        owner: String(row[headers.indexOf('责任人')] || '').trim(),
        reviewer: String(row[headers.indexOf('审核人 / Reviewed By')] || '').trim(),
        reviewStatus: '主管审核中'
      });
    }
    console.log('待审核故障报告: ' + reports.length + ' 条');
    return reports;
  } catch (e) {
    console.error('读取待审核报告失败: ' + e.message);
    return [];
  }
}

function _getEmailToSupervisorMap() {
  var map = {};
  try {
    var sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet) return map;
    var data = sheet.getDataRange().getValues();
    if (data.length < 3) return map;
    var headers = data[1];
    var idxEmail = headers.indexOf('GMail');
    var idxSupervisor = 60;
    if (idxEmail < 0) return map;
    for (var i = 2; i < data.length; i++) {
      var email = String(data[i][idxEmail] || '').trim().toLowerCase();
      var sv = String(data[i][idxSupervisor] || '').trim();
      if (email && sv) map[email] = sv;
    }
    console.log('Email→Supervisor映射: ' + Object.keys(map).length + ' 条');
  } catch (e) {
    console.error('读取Supervisor映射失败: ' + e.message);
  }
  return map;
}

function _extractEmail(field) {
  if (!field) return '';
  var m = String(field).match(/【(.+?)】/);
  return m ? m[1].trim().toLowerCase() : '';
}

function _getEmailToNameMap() {
  var map = {};
  try {
    var sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet) return map;
    var data = sheet.getDataRange().getValues();
    if (data.length < 3) return map;
    var headers = data[1];
    var idxEmail = headers.indexOf('GMail');
    var idxName = headers.indexOf('NAME');
    if (idxEmail < 0 || idxName < 0) return map;
    for (var i = 2; i < data.length; i++) {
      var email = String(data[i][idxEmail] || '').trim().toLowerCase();
      var name = String(data[i][idxName] || '').trim();
      if (email && name) map[email] = name;
    }
  } catch (e) {
    console.error('读取Name映射失败: ' + e.message);
  }
  return map;
}

// ========== 邮件发送 ==========

function sendFaultNotifications(faultItems, notificationConfig, trigger) {
  var sentCount = 0;

  try {
    // 获取待审核报告
    var pendingReviews = _getPendingReviewReports();
    var supervisorMap = _getEmailToSupervisorMap();
    var nameMap = _getEmailToNameMap();

    // 待审核报告按工序分组
    var pendingByProcess = {};
    pendingReviews.forEach(function(r) {
      var p = r.process;
      if (!p) return;
      if (!pendingByProcess[p]) pendingByProcess[p] = [];
      pendingByProcess[p].push(r);
    });

    var groupedItems = {};
    faultItems.forEach(function(item) {
      var pt = item.processType;
      if (!groupedItems[pt]) groupedItems[pt] = [];
      groupedItems[pt].push(item);
    });

    // 合并所有出现的工序
    var allProcesses = {};
    Object.keys(groupedItems).forEach(function(p) { allProcesses[p] = true; });
    Object.keys(pendingByProcess).forEach(function(p) { allProcesses[p] = true; });

    Object.keys(allProcesses).forEach(function(processType) {
      var items = groupedItems[processType] || [];
      var pendingForProcess = pendingByProcess[processType] || [];

      // 按提交日期降序排列，最近提交的优先展示
      items.sort(function(a, b) {
        var da = a.submitDate ? new Date(a.submitDate) : new Date(0);
        var db = b.submitDate ? new Date(b.submitDate) : new Date(0);
        return db - da;
      });

      // 收集待审核报告的 supervisor 邮箱
      var supervisorEmails = [];
      var svSeen = {};
      pendingForProcess.forEach(function(r) {
        var ownerEmail = _extractEmail(r.owner);
        var sv = supervisorMap[ownerEmail];
        if (sv) {
          if (!svSeen[sv]) { svSeen[sv] = true; supervisorEmails.push(sv); }
          r._supervisorName = nameMap[sv] || '';
        }
      });

      var recipients = getRecipients(processType, notificationConfig);

      if (recipients.length > 0 || pendingForProcess.length > 0) {
        var success = sendProcessEmail(processType, items, recipients, trigger, pendingForProcess, supervisorEmails);
        if (success) sentCount++;
      }
    });

  } catch (error) {
    console.error('发送故障通知时出错:', error);
    throw error;
  }

  return sentCount;
}

function getRecipients(processType, notificationConfig) {
  var recipients = [];

  try {
    notificationConfig.forEach(function(config) {
      if (config.process === processType) {
        var emails = config.mail.split(';').map(function(e) { return e.trim(); }).filter(function(e) { return e; });
        recipients = recipients.concat(emails);
      }
    });

  } catch (error) {
    console.error('查找工序 ' + processType + ' 通知收件人时出错:', error);
  }

  // 去重
  var unique = {};
  return recipients.filter(function(e) {
    if (unique[e]) return false;
    unique[e] = true;
    return true;
  });
}

function sendProcessEmail(processType, faultItems, recipients, trigger, pendingReviews, supervisorEmails) {
  try {
    var subject = generateEmailSubject(processType, faultItems, pendingReviews ? pendingReviews.length : 0);
    var body = generateEmailBody(processType, faultItems, pendingReviews);

    // CC: 管理员 + supervisor 邮箱
    var ccList = [FAULT_CONFIG.ADMIN_EMAIL];
    if (supervisorEmails && supervisorEmails.length > 0) {
      ccList = ccList.concat(supervisorEmails);
    }
    var ccSeen = {};
    var ccUnique = ccList.filter(function(e) { if (ccSeen[e]) return false; ccSeen[e] = true; return true; });

    recipients.forEach(function(recipient) {
      try {
        GmailApp.sendEmail(recipient, subject, '', { htmlBody: body, cc: ccUnique.join(',') });
        console.log('成功发送邮件给: ' + recipient);
        writeLog('sendProcessEmail', '成功', '已发送至 ' + recipient + '，故障=' + faultItems.length + ' 待审核=' + (pendingReviews ? pendingReviews.length : 0), trigger, processType);
      } catch (emailError) {
        console.error('发送邮件给 ' + recipient + ' 时出错:', emailError);
        writeLog('sendProcessEmail', '失败', '发送至 ' + recipient + ' 失败: ' + emailError.message, trigger, processType);
      }
    });

    return true;

  } catch (error) {
    console.error('发送工序 ' + processType + ' 通知时出错:', error);
    writeLog('sendProcessEmail', '失败', processType + ': ' + error.message, trigger, error.stack || '');
    return false;
  }
}

function generateEmailSubject(processType, faultItems, pendingCount) {
  var displayName = getProcessDisplayName(processType);
  var date = Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd');
  var parts = [];
  if (faultItems.length > 0) parts.push(faultItems.length + '个故障待判断');
  if (pendingCount > 0) parts.push(pendingCount + '个故障报告待审核');
  var detail = parts.length > 0 ? parts.join('和') : '无待处理项';
  return FAULT_CONFIG.EMAIL_SUBJECT_PREFIX + ' ' + date + ' - ' + displayName + '工序发现' + detail;
}

// ========== 工序名称 ==========

function getProcessDisplayName(processType) {
  var names = { 'INJ': '注塑', 'TF': '植磨毛', 'PK': '包装' };
  return names[processType] || processType;
}

function getProcessDisplayNameEn(processType) {
  var names = { 'INJ': 'Injection Molding', 'TF': 'Tufting', 'PK': 'Packaging' };
  return names[processType] || processType;
}

// ========== HTML 邮件正文 ==========

function generateEmailBody(processType, faultItems, pendingReviews) {
  var processName = getProcessDisplayName(processType);
  var totalCount = faultItems.length;
  var pendingCount = pendingReviews ? pendingReviews.length : 0;
  var maxShow = FAULT_CONFIG.MAX_ITEMS_IN_EMAIL;
  var displayItems = faultItems.slice(0, maxShow);
  var omittedCount = totalCount - maxShow;

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>故障报告提醒</title><style>';
  html += 'body{font-family:\'Microsoft YaHei\',Arial,sans-serif;margin:0;padding:20px;background-color:#f5f5f5}';
  html += '.container{max-width:1200px;margin:0 auto}';
  html += '.header{background:linear-gradient(135deg,#E60012 0%,#FF6B6B 100%);color:white;padding:30px;border-radius:15px;text-align:center;margin-bottom:30px;box-shadow:0 8px 25px rgba(230,0,18,0.3)}';
  html += '.header h1{margin:0;font-size:28px;font-weight:600}';
  html += '.header p{margin:10px 0 0 0;font-size:16px;opacity:0.9}';
  html += '.cards{display:flex;gap:20px;margin-bottom:25px;flex-wrap:wrap}';
  html += '.card{flex:1;min-width:200px;background:white;border-radius:12px;padding:20px;text-align:center;box-shadow:0 4px 15px rgba(0,0,0,0.1)}';
  html += '.card-value{font-size:36px;font-weight:bold}';
  html += '.card-label{font-size:13px;color:#7f8c8d;margin-top:4px}';
  html += '.table-container{background:white;border-radius:15px;padding:25px;margin-bottom:25px;box-shadow:0 4px 20px rgba(0,0,0,0.1);overflow-x:auto}';
  html += '.fault-table{width:100%;border-collapse:collapse;margin-top:20px}';
  html += '.fault-table th{background:#E60012;color:white;padding:12px 8px;text-align:center;font-weight:600;border:1px solid #d32f2f;line-height:1.4}';
  html += '.fault-table td{padding:10px 8px;border:1px solid #ddd;text-align:left}';
  html += '.fault-table tr:nth-child(even){background-color:#f9f9f9}';
  html += '.fault-table tr:hover{background-color:#f0f0f0}';
  html += '.review-table{width:100%;border-collapse:collapse;margin-top:20px}';
  html += '.review-table th{background:#f39c12;color:white;padding:12px 8px;text-align:center;font-weight:600;border:1px solid #e67e22;line-height:1.4}';
  html += '.review-table td{padding:10px 8px;border:1px solid #ddd;text-align:left}';
  html += '.review-table tr:nth-child(even){background-color:#fffbf0}';
  html += '.review-table tr:hover{background-color:#fff8e1}';
  html += '.repair-time{color:#E60012;font-weight:bold}';
  html += '.status-badge{background:#FF6B6B;color:white;padding:4px 8px;border-radius:12px;font-size:12px}';
  html += '.review-badge{background:#f39c12;color:white;padding:4px 8px;border-radius:12px;font-size:12px}';
  html += '.footer{text-align:center;margin-top:40px;color:#666;font-size:12px}';
  html += '@media(max-width:768px){.fault-table,.review-table{font-size:12px}.fault-table th,.fault-table td,.review-table th,.review-table td{padding:6px 4px}.cards{flex-direction:column}}';
  html += '</style></head><body><div class="container">';

  html += '<div class="header">';
  html += '<h1>【故障报告提醒】Fault Report Reminder</h1>';
  html += '<p>' + processName + '工序 / ' + getProcessDisplayNameEn(processType) + ' Process</p>';
  html += '</div>';

  // ===== 两张卡片 =====
  html += '<div class="cards">';
  html += '<div class="card"><div class="card-value" style="color:#E60012">' + totalCount + '</div><div class="card-label">待分配故障报告<br>Pending Fault Reports</div></div>';
  html += '<div class="card"><div class="card-value" style="color:#f39c12">' + pendingCount + '</div><div class="card-label">待审核故障报告<br>Pending Review Reports</div></div>';
  html += '</div>';

  if (omittedCount > 0) {
    html += '<div style="background:#FFF3CD;border:1px solid #FFC107;border-radius:8px;padding:15px;margin-bottom:20px;text-align:center">';
    html += '<strong>共 ' + totalCount + ' 条，以下展示最近 ' + maxShow + ' 条</strong>，还有 ' + omittedCount + ' 条请前往<a href="https://script.google.com/macros/s/AKfycbyaQjG5yFGYxU825DrODhSLl2bdfbYKpqAH4qOIzKoTJ4b-5qU/exec?v=FailureReport_Manage" target="_blank">故障报告管理模块</a>查看完整清单';
    html += '</div>';
  }

  html += '<div class="table-container">';
  html += '<h2 style="margin-top:0;color:#333">【故障详情列表】Fault Details List</h2>';
  html += '<table class="fault-table"><thead><tr>';
  html += '<th>序号<br>No.</th><th>故障编号<br>Fault ID</th><th>机台号<br>Machine No.</th><th>车间<br>Workshop</th><th>维修人<br>Repair Person</th><th>问题描述<br>Problem Description</th><th>处理过程<br>Process</th><th>维修时间<br>Repair Time</th><th>提交日期<br>Submit Date</th><th>状态<br>Status</th>';
  html += '</tr></thead><tbody>';

  displayItems.forEach(function(item, idx) {
    html += '<tr>';
    html += '<td><strong>#' + (idx + 1) + '</strong></td>';
    html += '<td>' + (item.id || '-') + '</td>';
    html += '<td>' + (item.machineNo || '-') + '</td>';
    html += '<td>' + (item.workshop || '-') + '</td>';
    html += '<td>' + (item.repairPerson || '-') + '</td>';
    html += '<td style="max-width:200px;word-wrap:break-word">' + (item.problemDesc || '-') + '</td>';
    html += '<td style="max-width:200px;word-wrap:break-word">' + (item.process || '-') + '</td>';
    html += '<td class="repair-time">' + (item.repairTime || 0) + ' min</td>';
    html += '<td>' + (item.submitDate ? formatVariableAsDate(item.submitDate) : '-') + '</td>';
    html += '<td><span class="status-badge">Pending</span></td>';
    html += '</tr>';
  });

  html += '</tbody></table></div>';

  if (omittedCount > 0) {
    html += '<p style="text-align:center;color:#E60012;font-weight:bold">（共 ' + totalCount + ' 条，以上为前 ' + maxShow + ' 条，剩余 ' + omittedCount + ' 条未展示）</p>';
  }

  // ===== 待审核故障报告 =====
  if (pendingCount > 0) {
    html += '<div class="table-container">';
    html += '<h2 style="margin-top:0;color:#e67e22">【待审核故障报告】Pending Review Reports (' + pendingCount + '条)</h2>';
    html += '<table class="review-table"><thead><tr>';
    html += '<th>编号<br>ID</th><th>机台号<br>Machine No.</th><th>问题描述<br>Problem Description</th><th>责任人<br>Owner</th><th>故障报告编号<br>Report No.</th><th>审核人<br>Reviewer</th><th>状态<br>Status</th>';
    html += '</tr></thead><tbody>';

    pendingReviews.forEach(function(r) {
      html += '<tr>';
      html += '<td>' + (r.id || '-') + '</td>';
      html += '<td>' + (r.machineId || '-') + '</td>';
      html += '<td style="max-width:200px;word-wrap:break-word">' + (r.description || '-') + '</td>';
      html += '<td>' + (r.owner || '-') + '</td>';
      html += '<td>' + (r.reportNumber || '-') + '</td>';
      html += '<td>' + (r._supervisorName || r.reviewer || '-') + '</td>';
      html += '<td><span class="review-badge">' + r.reviewStatus + '</span></td>';
      html += '</tr>';
    });

    html += '</tbody></table></div>';
  }

  html += '<div class="footer">';
  html += '<p>此邮件由EDS故障报告邮件通知系统自动发送<br>This email is automatically sent by EDS Fault Report Email Notification System</p>';
  html += '<p>请及时处理相关故障报告，确保设备正常运行<br>Please process the relevant fault reports promptly to ensure normal equipment operation</p>';
  html += '</div></div></body></html>';

  return html;
}

// ========== 错误通知 ==========

function sendErrorNotification(error, trigger) {
  try {
    var subject = '[系统错误] 故障报告邮件通知系统';
    var body = '故障报告邮件通知系统运行出错:\n\n错误信息: ' + error.message + '\n时间: ' + Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd HH:mm:ss') + '\n\n请检查系统日志。';

    GmailApp.sendEmail(FAULT_CONFIG.ADMIN_EMAIL, subject, body);

    writeLog('sendErrorNotification', '成功', '已发送错误通知至 ' + FAULT_CONFIG.ADMIN_EMAIL, trigger, error.message);

  } catch (emailError) {
    console.error('发送错误通知时出错:', emailError);
  }
}

// ========== 测试函数 ==========

function testFaultEmailSending() {
  var testItems = [{
    id: 'TEST001',
    processType: 'INJ',
    repairTime: 300,
    needFaultReport: '',
    status: '处理中',
    machineNo: 'V1FTA459',
    workshop: 'TB1',
    shift: 'A班',
    repairPerson: '测试用户',
    problemDesc: '测试故障描述',
    process: '测试处理过程',
    faultCode: 'TEST_CODE',
    workOrderNo: 'TEST_WO001',
    submitDate: '2024-12-01'
  }];

  var testConfig = [{
    process: 'INJ',
    mail: FAULT_CONFIG.ADMIN_EMAIL
  }];

  var sentCount = sendFaultNotifications(testItems, testConfig);
  console.log('测试邮件发送完成，成功发送 ' + sentCount + ' 封邮件');
  writeLog('testFaultEmailSending', '成功', '测试邮件已发送', '手动', '');
}
