// V20260527.1 — 故障报告邮件通知模块
// 功能：自动检测交接班模块故障条目，对超时且未填写故障报告的条目发送邮件通知

// ========== 故障报告配置 ==========
const FAULT_CONFIG = {
  // 数据源：交接班模块表格
  SHEET_ID: '10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w',

  // 工序车间数据表名称
  SHIFT_SHEETS: [
    'Shift_INJ_TB1', 'Shift_INJ_TB2',
    'Shift_TF_TB1', 'Shift_TF_TB2',
    'Shift_PK_TB1', 'Shift_PK_TB2'
  ],

  // 通知配置表名称（位于 SHEET_ID 表格中）
  NOTIFICATION_SHEET: '通知清单',

  // 工序时间阈值（分钟）
  PROCESS_THRESHOLDS: {
    'INJ': 240,
    'TF': 120,
    'PK': 60
  },

  // 工序字段映射（交接班模块工序 → 通知清单工序）
  PROCESS_MAPPING: {
    'INJ': 'IM',
    'TF': 'TF',
    'PK': 'PK'
  },

  EMAIL_SUBJECT_PREFIX: '[故障报告提醒]',
  ADMIN_EMAIL: 'kelland_zhao@colpal.com'
};

// ========== 邮件发件人（运行时获取，避免顶层 OAuth 报错）==========
function getFaultEmailSender() {
  return Session.getActiveUser().getEmail();
}

// ========== 故障检测主函数 ==========

/** 主函数：故障检测和通知发送（由 15 分钟定时触发器调用） */
function mainFaultDetection() {
  const fnName = 'mainFaultDetection';
  try {
    console.log('开始故障检测流程...');

    const faultItems = getAllFaultItems();
    console.log('获取到 ' + faultItems.length + ' 个故障条目');

    if (faultItems.length === 0) {
      writeLog(fnName, '成功', '无故障条目', '定时', '');
      return;
    }

    const qualifiedItems = filterQualifiedFaultItems(faultItems);
    console.log('筛选出 ' + qualifiedItems.length + ' 个符合通知条件的故障条目');

    if (qualifiedItems.length === 0) {
      writeLog(fnName, '成功', '无符合通知条件的条目（共扫描 ' + faultItems.length + ' 条）', '定时', '');
      return;
    }

    const notificationConfig = getNotificationConfig();
    console.log('获取到 ' + notificationConfig.length + ' 条通知配置');

    const sentCount = sendFaultNotifications(qualifiedItems, notificationConfig);
    console.log('成功发送 ' + sentCount + ' 封通知邮件');

    writeLog(fnName, '成功', '发送 ' + sentCount + ' 封通知，覆盖 ' + qualifiedItems.length + ' 个故障条目', '定时', '');

  } catch (error) {
    console.error('故障检测流程出错:', error);
    writeLog(fnName, '失败', error.message, '定时', error.stack || '');
    sendErrorNotification(error);
  }
}

// ========== 诊断函数 ==========

/** 诊断：输出每条故障的筛选细节到 Log，帮助排查收不到邮件的原因 */
function debugFaultDetection() {
  var items = getAllFaultItems();
  if (items.length === 0) {
    writeLog('debugFaultDetection', '诊断', 'getAllFaultItems 返回 0 条记录', '手动', '检查6个车间表是否有数据、表头是否匹配');
    return;
  }

  var report = [];
  items.forEach(function(item) {
    var reasons = [];
    var passed = true;

    if (!(item.processType in FAULT_CONFIG.PROCESS_THRESHOLDS)) {
      reasons.push('工序不支持: ' + item.processType);
      passed = false;
    }

    var threshold = FAULT_CONFIG.PROCESS_THRESHOLDS[item.processType] || 0;
    var repairTime = Number(item.repairTime) || 0;
    if (passed && repairTime < threshold) {
      reasons.push('维修时间不足: ' + repairTime + 'min(阈值' + threshold + 'min)');
      passed = false;
    }

    if (passed && item.needFaultReport !== '' && item.needFaultReport != null) {
      reasons.push('已填写故障报告: "' + item.needFaultReport + '"');
      passed = false;
    }

    var statusStr = String(item.status);
    if (passed && statusStr.indexOf('已解决') === -1 && statusStr.indexOf('Solved') === -1) {
      reasons.push('状态未解决: "' + item.status + '"');
      passed = false;
    }

    report.push((passed ? '✓' : '✗') + ' 编号:' + item.id +
      ' | 工序:' + item.processType +
      ' | 维修时间:' + repairTime + 'min' +
      ' | 故障报告:' + (item.needFaultReport || '(空)') +
      ' | 状态:' + item.status +
      ' | 车间:' + item.workshop +
      (reasons.length > 0 ? ' | 原因: ' + reasons.join('; ') : ''));
  });

  writeLog('debugFaultDetection', '诊断', '共' + items.length + '条\n' + report.join('\n'), '手动', '');

  // 同时输出通知配置
  var configs = getNotificationConfig();
  if (configs.length === 0) {
    writeLog('debugFaultDetection', '诊断', '通知配置为空，检查通知清单 Function=故障报告分配', '手动', '');
  } else {
    var cfgInfo = configs.map(function(c) { return c.process + ' → ' + c.mail; }).join(' | ');
    writeLog('debugFaultDetection', '诊断', '通知配置: ' + cfgInfo, '手动', '');
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
      workshop: extractWorkshop(sheetName),
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
    const spreadsheet = SpreadsheetApp.openById(FAULT_CONFIG.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(FAULT_CONFIG.NOTIFICATION_SHEET);

    if (!sheet) {
      console.error('找不到通知配置表: ' + FAULT_CONFIG.NOTIFICATION_SHEET);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const configs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[headers.indexOf('Function')] === '故障报告分配') {
        const config = {
          function: row[headers.indexOf('Function')],
          process: row[headers.indexOf('Process')],
          workshop: row[headers.indexOf('Workshop')],
          mail: row[headers.indexOf('Mail')]
        };

        if (config.process && config.workshop && config.mail) {
          configs.push(config);
        }
      }
    }

    return configs;

  } catch (error) {
    console.error('获取通知配置时出错:', error);
    throw error;
  }
}

// ========== 邮件发送 ==========

function sendFaultNotifications(faultItems, notificationConfig) {
  var sentCount = 0;

  try {
    var groupedItems = {};
    faultItems.forEach(function(item) {
      var pt = item.processType;
      if (!groupedItems[pt]) groupedItems[pt] = [];
      groupedItems[pt].push(item);
    });

    Object.keys(groupedItems).forEach(function(processType) {
      var items = groupedItems[processType];
      var recipients = getRecipients(processType, notificationConfig);

      if (recipients.length > 0) {
        var success = sendProcessEmail(processType, items, recipients);
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
    var mappedType = FAULT_CONFIG.PROCESS_MAPPING[processType] || processType;

    notificationConfig.forEach(function(config) {
      if (config.process === mappedType) {
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

function sendProcessEmail(processType, faultItems, recipients) {
  try {
    var subject = generateEmailSubject(processType, faultItems);
    var body = generateEmailBody(processType, faultItems);
    var sender = getFaultEmailSender();

    recipients.forEach(function(recipient) {
      try {
        GmailApp.sendEmail(recipient, subject, '', {
          htmlBody: body,
          from: sender
        });
        console.log('成功发送邮件给: ' + recipient);
      } catch (emailError) {
        console.error('发送邮件给 ' + recipient + ' 时出错:', emailError);
      }
    });

    return true;

  } catch (error) {
    console.error('发送工序 ' + processType + ' 通知时出错:', error);
    return false;
  }
}

function generateEmailSubject(processType, faultItems) {
  var displayName = getProcessDisplayName(processType);
  var count = faultItems.length;
  var date = Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd');

  return FAULT_CONFIG.EMAIL_SUBJECT_PREFIX + ' ' + date + ' - ' + displayName + '工序发现' + count + '个需要填写故障报告的问题';
}

// ========== 工序名称 ==========

function getProcessDisplayName(processType) {
  var names = { 'INJ': '注塑', 'TF': '涂装', 'PK': '包装' };
  return names[processType] || processType;
}

function getProcessDisplayNameEn(processType) {
  var names = { 'INJ': 'Injection Molding', 'TF': 'Painting', 'PK': 'Packaging' };
  return names[processType] || processType;
}

// ========== HTML 邮件正文 ==========

function generateEmailBody(processType, faultItems) {
  var processName = getProcessDisplayName(processType);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>故障报告提醒</title><style>';
  html += 'body{font-family:\'Microsoft YaHei\',Arial,sans-serif;margin:0;padding:20px;background-color:#f5f5f5}';
  html += '.container{max-width:1200px;margin:0 auto}';
  html += '.header{background:linear-gradient(135deg,#E60012 0%,#FF6B6B 100%);color:white;padding:30px;border-radius:15px;text-align:center;margin-bottom:30px;box-shadow:0 8px 25px rgba(230,0,18,0.3)}';
  html += '.header h1{margin:0;font-size:28px;font-weight:600}';
  html += '.header p{margin:10px 0 0 0;font-size:16px;opacity:0.9}';
  html += '.table-container{background:white;border-radius:15px;padding:25px;margin-bottom:25px;box-shadow:0 4px 20px rgba(0,0,0,0.1);overflow-x:auto}';
  html += '.fault-table{width:100%;border-collapse:collapse;margin-top:20px}';
  html += '.fault-table th{background:#E60012;color:white;padding:12px 8px;text-align:center;font-weight:600;border:1px solid #d32f2f;line-height:1.4}';
  html += '.fault-table td{padding:10px 8px;border:1px solid #ddd;text-align:left}';
  html += '.fault-table tr:nth-child(even){background-color:#f9f9f9}';
  html += '.fault-table tr:hover{background-color:#f0f0f0}';
  html += '.repair-time{color:#E60012;font-weight:bold}';
  html += '.status-badge{background:#FF6B6B;color:white;padding:4px 8px;border-radius:12px;font-size:12px}';
  html += '.footer{text-align:center;margin-top:40px;color:#666;font-size:12px}';
  html += '@media(max-width:768px){.fault-table{font-size:12px}.fault-table th,.fault-table td{padding:6px 4px}}';
  html += '</style></head><body><div class="container">';

  html += '<div class="header">';
  html += '<h1>【故障报告提醒】Fault Report Reminder</h1>';
  html += '<p>' + processName + '工序发现需要判定是否需要故障报告<br>' + getProcessDisplayNameEn(processType) + ' process found issues that need to determine whether fault reports are required</p>';
  html += '</div>';

  html += '<div class="table-container">';
  html += '<h2 style="margin-top:0;color:#333">【故障详情列表】Fault Details List</h2>';
  html += '<table class="fault-table"><thead><tr>';
  html += '<th>序号<br>No.</th><th>故障编号<br>Fault ID</th><th>机台号<br>Machine No.</th><th>车间<br>Workshop</th><th>维修人<br>Repair Person</th><th>问题描述<br>Problem Description</th><th>处理过程<br>Process</th><th>维修时间<br>Repair Time</th><th>提交日期<br>Submit Date</th><th>状态<br>Status</th>';
  html += '</tr></thead><tbody>';

  faultItems.forEach(function(item, idx) {
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
  html += '<div class="footer">';
  html += '<p>此邮件由EDS故障报告邮件通知系统自动发送<br>This email is automatically sent by EDS Fault Report Email Notification System</p>';
  html += '<p>请及时处理相关故障报告，确保设备正常运行<br>Please process the relevant fault reports promptly to ensure normal equipment operation</p>';
  html += '</div></div></body></html>';

  return html;
}

// ========== 错误通知 ==========

function sendErrorNotification(error) {
  try {
    var subject = '[系统错误] 故障报告邮件通知系统';
    var body = '故障报告邮件通知系统运行出错:\n\n错误信息: ' + error.message + '\n时间: ' + Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd HH:mm:ss') + '\n\n请检查系统日志。';

    GmailApp.sendEmail(FAULT_CONFIG.ADMIN_EMAIL, subject, body, {
      from: getFaultEmailSender()
    });

    writeLog('sendErrorNotification', '成功', '已发送错误通知至 ' + FAULT_CONFIG.ADMIN_EMAIL, '定时', error.message);

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
    function: '故障报告分配',
    process: 'IM',
    workshop: 'TB1',
    mail: getFaultEmailSender()
  }];

  var sentCount = sendFaultNotifications(testItems, testConfig);
  console.log('测试邮件发送完成，成功发送 ' + sentCount + ' 封邮件');
  writeLog('testFaultEmailSending', '成功', '测试邮件已发送', '手动', '');
}
