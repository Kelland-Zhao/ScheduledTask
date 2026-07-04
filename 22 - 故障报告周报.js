// V20260608.1 — 故障报告周报提醒 + 共享工具（22/24 依赖此文件）
// 功能：每周一提醒所有未上传的故障报告，按工序分组发送
// 数据源：Failure_Database（来自原 Failure_Report_Reminder 项目表格）

// ========== 故障报告配置 ==========
const FR_CONFIG = {
  SPREADSHEET_ID: '1YAPdZKVEOHgCGIJRQwWTQBmwaWIS4yd1SQKJJfRCtAU',
  FAILURE_SHEET_NAME: 'Failure_Database',
  FOLLOWUP_SHEET_NAME: 'Failure_Report_followup',
  OVERDUE_DAYS: 7
};

// ========== 工序显示名称 ==========
const PROCESS_DISPLAY_NAMES = {
  'INJ': '注塑',
  'TF': '植磨毛',
  'PK': '包装'
};

// ========== 主函数：每周故障报告提醒 ==========
function weeklyFailureReportReminder(e) {
  const trigger = e ? '定时' : '手动';
  try {
    console.log('=== 开始执行每周故障报告提醒 ===');

    const failureData = getFailureReportData();
    if (!failureData || failureData.length === 0) {
      console.log('没有找到需要提醒的故障报告数据');
      writeLog('weeklyFailureReportReminder', '成功', '没有故障报告数据，跳过执行', trigger, '');
      return;
    }

    const { nameToUser, processToAdmins } = getUserIDLookupMaps();
    if (Object.keys(nameToUser).length === 0) {
      console.warn('⚠️ userID 权限表无有效数据，无法发送周报');
      writeLog('weeklyFailureReportReminder', '警告', 'userID 权限表无有效数据，跳过执行', trigger, '');
      return;
    }

    sendWeeklyRemindersByProcess(failureData, nameToUser, processToAdmins, trigger);

    console.log('=== 每周故障报告提醒执行完成 ===');

  } catch (error) {
    console.error('每周故障报告提醒执行出错:', error);
    writeLog('weeklyFailureReportReminder', '失败', error.message, trigger, '');
  }
}

// ========== 数据读取 ==========

function getFailureReportData() {
  try {
    console.log('开始获取故障报告数据...');

    const spreadsheet = SpreadsheetApp.openById(FR_CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(FR_CONFIG.FAILURE_SHEET_NAME);

    if (!sheet) {
      console.error('故障报告数据表未找到');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('故障报告数据表为空');
      return [];
    }

    const headers = data[0];
    const fieldIndexes = getFieldIndexes(headers);

    const failureData = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const failureRecord = {
        id: row[fieldIndexes['编号']] || '',
        machineId: row[fieldIndexes['机台号']] || '',
        description: row[fieldIndexes['问题描述']] || '',
        submitDate: row[fieldIndexes['提交日期']] || '',
        workshop: row[fieldIndexes['车间']] || '',
        process: row[fieldIndexes['工序']] || '',
        reportNumber: row[fieldIndexes['故障报告编号']] || '',
        assignDate: row[fieldIndexes['分配日期']] || '',
        uploadDate: row[fieldIndexes['上传日期']] || '',
        attachment: row[fieldIndexes['附件']] || '',
        owner: row[fieldIndexes['责任人']] || '',
        overdueDays: calculateOverdueDays(row[fieldIndexes['分配日期']])
      };

      failureData.push(failureRecord);
    }

    console.log('✅ 成功获取 ' + failureData.length + ' 条故障报告数据');
    return failureData;

  } catch (error) {
    console.error('获取故障报告数据时出错:', error);
    return [];
  }
}

function getOverdueFailureReportData(overdueDays) {
  try {
    const threshold = overdueDays !== undefined ? overdueDays : FR_CONFIG.OVERDUE_DAYS;
    const allData = getFailureReportData();

    const overdueData = allData.filter(function(record) {
      const isOverdue = record.overdueDays >= threshold;
      const isUnuploaded = !isFailureReportUploaded(record).isUploaded;
      return isOverdue && isUnuploaded;
    });

    console.log('找到 ' + overdueData.length + ' 条超期且未上传的故障报告（超期天数 >= ' + threshold + '天）');

    overdueData.sort(function(a, b) { return b.overdueDays - a.overdueDays; });

    return overdueData;

  } catch (error) {
    console.error('获取超期故障报告数据时出错:', error);
    return [];
  }
}

// ========== userID 权限表读取 ==========

function getUserIDLookupMaps() {
  try {
    console.log('=== 开始读取 userID 权限表 ===');

    const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID)
      .getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet) { console.error('userID 表未找到'); return { nameToUser: {}, processToAdmins: {} }; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 3) {
      console.warn('userID 表数据不足（需要至少3行：cat/header/data）');
      return { nameToUser: {}, processToAdmins: {} };
    }

    const headers = data[1];
    const idxName = headers.findIndex(function(h) { return h === 'NAME'; });
    const idxEmail = headers.findIndex(function(h) { return h === 'GMail'; });
    const idxProcess = 14;
    const idxFaultAdmin = 57;
    const idxLineMgr = 60;

    if (idxName === -1 || idxEmail === -1) {
      console.error('userID 表缺少 NAME 或 GMail 列');
      return { nameToUser: {}, processToAdmins: {} };
    }

    const nameToUser = {};
    const processToAdmins = {};
    const nameToProcess = {};

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const name = String(row[idxName] || '').trim();
      const email = String(row[idxEmail] || '').trim();
      const process = String(row[idxProcess] || '').trim();
      const faultAdmin = String(row[idxFaultAdmin] || '').trim();
      const lineMgr = String(row[idxLineMgr] || '').trim();

      if (name && email) {
        nameToUser[name] = { email: email, lineManager: lineMgr };
      }

      if (name && process && process !== 'ALL') {
        nameToProcess[name] = process;
      }

      if (process && process !== 'ALL' && email && faultAdmin === 'Y') {
        if (!processToAdmins[process]) processToAdmins[process] = [];
        if (processToAdmins[process].indexOf(email) === -1) {
          processToAdmins[process].push(email);
        }
      }
    }

    console.log('✅ nameToUser: ' + Object.keys(nameToUser).length + ' 条记录');
    console.log('✅ processToAdmins 工序: [' + Object.keys(processToAdmins).join(', ') + ']');

    return { nameToUser: nameToUser, processToAdmins: processToAdmins, nameToProcess: nameToProcess };

  } catch (error) {
    console.error('读取 userID 权限表时出错:', error);
    return { nameToUser: {}, processToAdmins: {}, nameToProcess: {} };
  }
}

function mapFailureProcessToUserID(failureProcess) {
  const MAP = { 'IM': 'INJ' };
  return MAP[failureProcess] || failureProcess;
}

// ========== 人员字段解析（共享函数） ==========

function extractNameFromPersonField(field) {
  if (!field) return '';
  return String(field).replace(/【.+?】/, '').trim();
}

function extractEmailFromPersonField(field) {
  if (!field) return '';
  const match = String(field).match(/【(.+?)】/);
  return match ? match[1].trim() : '';
}

// ========== 周报邮件发送（按工序） ==========

function sendWeeklyRemindersByProcess(failureData, nameToUser, processToAdmins, trigger) {
  try {
    console.log('=== 开始按工序发送周报邮件 ===');

    var unuploaded = failureData.filter(function(r) { return !isFailureReportUploaded(r).isUploaded; });
    console.log('📊 未上传故障报告: ' + unuploaded.length + ' 条');

    var processGroups = {};
    unuploaded.forEach(function(record) {
      var mappedProcess = mapFailureProcessToUserID(record.process);
      if (!mappedProcess) return;
      if (!processGroups[mappedProcess]) processGroups[mappedProcess] = [];
      processGroups[mappedProcess].push(record);
    });

    var processes = Object.keys(processGroups);
    console.log('📊 涉及工序: [' + processes.join(', ') + ']');

    if (processes.length === 0) {
      console.log('⚠️ 没有需要发送的工序');
      writeLog('weeklyFailureReportReminder', '成功', '没有需要发送的工序', trigger, '');
      return;
    }

    var sentCount = 0;

    for (var p = 0; p < processes.length; p++) {
      var process = processes[p];
      var records = processGroups[process];

      // 提取唯一责任人姓名
      var ownerRawNames = [];
      var seen = {};
      records.forEach(function(r) {
        var raw = String(r.owner || '').trim();
        if (raw && !seen[raw]) { seen[raw] = true; ownerRawNames.push(raw); }
      });
      var ownerLookupNames = ownerRawNames.map(function(raw) { return extractNameFromPersonField(raw); });
      console.log('\n🔍 工序 ' + process + ': ' + records.length + '条报告, 责任人: [' + ownerLookupNames.join(', ') + ']');

      var toEmails = [];
      var ccLineMgrs = [];
      var unmatchedNames = [];

      ownerLookupNames.forEach(function(name, idx) {
        var rawName = ownerRawNames[idx];
        var user = nameToUser[name];
        if (user && user.email) {
          toEmails.push(user.email);
          if (user.lineManager) ccLineMgrs.push(user.lineManager);
          console.log('  ✅ ' + name + ' → ' + user.email + (user.lineManager ? ' (LM: ' + user.lineManager + ')' : ''));
        } else {
          unmatchedNames.push(rawName);
          console.warn('  ⚠️ ' + name + ' (原始: ' + rawName + ') → 未在 userID 中找到');
        }
      });

      if (unmatchedNames.length > 0) {
        console.warn('⚠️ 工序 ' + process + ': ' + unmatchedNames.length + ' 个责任人未匹配: [' + unmatchedNames.join(', ') + ']');
      }

      if (toEmails.length === 0) {
        console.warn('⚠️ 工序 ' + process + ': 没有有效的责任人邮箱，跳过发送');
        writeLog('weeklyFailureReportReminder', '警告', '工序 ' + process + ': 无有效责任人邮箱，跳过 (' + records.length + '条报告)', trigger, '');
        continue;
      }

      // CC 列表
      var processAdmins = processToAdmins[process] || [];
      var ccAll = ccLineMgrs.concat(processAdmins);
      ccAll.push('kelland_zhao@colpal.com');
      var ccSeen = {};
      var ccList = ccAll.filter(function(e) { if (ccSeen[e]) return false; ccSeen[e] = true; return true; }).filter(Boolean);

      // 生成邮件
      var subject = '【提醒】 故障报告定期提醒 - ' + process + ' / Failure Report Weekly Reminder - ' + process;
      var htmlBody = generateWeeklyProcessEmailContent(process, records);

      GmailApp.sendEmail(toEmails.join(','), subject, '请使用支持HTML的邮件客户端查看此邮件。', {
        htmlBody: htmlBody,
        name: '故障报告提醒系统',
        cc: ccList.join(',')
      });

      sentCount++;
      console.log('✅ 周报已发送 → 工序: ' + process);
      console.log('   To: ' + toEmails.join(', '));
      console.log('   CC: ' + ccList.join(', '));
      console.log('   报告: ' + records.length + '条');

      writeLog('weeklyFailureReportReminder', '成功', '工序 ' + process + ': 发送成功 → To:' + toEmails.length + '人 CC:' + ccList.length + '人 报告:' + records.length + '条', trigger, '');
    }

    console.log('\n=== 周报发送完成，共 ' + sentCount + ' 封邮件 ===');

  } catch (error) {
    console.error('按工序发送周报邮件时出错:', error);
    writeLog('weeklyFailureReportReminder', '失败', error.message, trigger, '');
  }
}

// ========== 周报邮件 HTML 模板 ==========

function generateWeeklyProcessEmailContent(processName, records) {
  try {
    var formattedDate = formatVariableAsDate(new Date());

    var emailBody = '<div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; background-color: #f8f9fa; padding: 20px;">' +
      '<div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">' +
      '<h2 style="color: #2c3e50; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #3498db; padding-bottom: 10px;">' +
      '【提醒】 故障报告定期提醒 - ' + processName + ' 工序<br>' +
      '<span style="font-size: 0.8em; color: #7f8c8d;">Failure Report Weekly Reminder - ' + processName + ' Process</span>' +
      '</h2>' +
      '<p style="color: #34495e; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">' +
      '您好！以下是本周（' + formattedDate + '）<strong>未上传</strong>的 ' + processName + ' 工序故障报告汇总：<br>' +
      '<span style="font-size: 0.9em; color: #7f8c8d;">Hello! Below is this week\'s (' + formattedDate + ') <strong>unuploaded</strong> failure report summary for ' + processName + ' process:</span>' +
      '</p></div>';

    if (records.length > 0) {
      emailBody += '<div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">' +
        '<h3 style="color: #2c3e50; margin-bottom: 20px; border-bottom: 2px solid #3498db; padding-bottom: 10px;">' +
        '【详情】 未上传故障报告 Unuploaded Failure Reports (' + records.length + '条)</h3>' +
        '<div style="overflow-x: auto;">' +
        '<table style="width: 100%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">' +
        '<thead>' +
        '<tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">编号<br><span style="font-size:0.8em;opacity:0.9;">ID</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">机台号<br><span style="font-size:0.8em;opacity:0.9;">Machine</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">问题描述<br><span style="font-size:0.8em;opacity:0.9;">Description</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">车间<br><span style="font-size:0.8em;opacity:0.9;">Workshop</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">分配日期<br><span style="font-size:0.8em;opacity:0.9;">Assign Date</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">超期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span></th>' +
        '</tr></thead><tbody>';

      records.forEach(function(record, i) {
        var rowStyle = i % 2 === 0 ? 'background-color: #f8f9fa;' : 'background-color: #ffffff;';

        var overdueDisplay = '';
        if (record.overdueDays >= FR_CONFIG.OVERDUE_DAYS) {
          overdueDisplay = '<div style="display:inline-block;text-align:center;"><div style="background:linear-gradient(135deg,#e74c3c,#c0392b);color:white;padding:6px 12px;border-radius:16px;font-weight:600;font-size:12px;box-shadow:0 2px 6px rgba(231,76,60,0.3);display:inline-block;min-width:80px;"><span style="display:block;">【超期】 ' + record.overdueDays + '天</span><span style="display:block;font-size:10px;opacity:0.9;">Days</span></div></div>';
        } else {
          overdueDisplay = '<div style="display:inline-block;text-align:center;"><div style="background:linear-gradient(135deg,#f39c12,#e67e22);color:white;padding:6px 12px;border-radius:16px;font-weight:600;font-size:12px;box-shadow:0 2px 6px rgba(243,156,18,0.3);display:inline-block;min-width:80px;"><span style="display:block;">' + record.overdueDays + '天</span><span style="display:block;font-size:10px;opacity:0.9;">Days</span></div></div>';
        }

        emailBody += '<tr style="' + rowStyle + '">' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">' + (record.id || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (record.machineId || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:200px;word-wrap:break-word;">' + (record.description || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (record.workshop || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (record.owner || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (formatVariableAsDate(record.assignDate) || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + overdueDisplay + '</td>' +
          '</tr>';
      });

      emailBody += '</tbody></table></div>' +
        '<div style="margin-top:15px;text-align:right;">' +
        '<span style="color:#7f8c8d;font-size:14px;">共 ' + records.length + ' 条 | Total: ' + records.length + ' reports</span></div></div>';
    }

    emailBody += '<div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px;">' +
      '<div style="text-align: center; color: #7f8c8d; font-size: 14px; line-height: 1.6;">' +
      '<p style="margin-bottom: 10px;">请及时查看并处理相关故障报告。<br>' +
      '<span style="font-size: 0.9em; opacity: 0.8;">Please review and handle related failure reports promptly.</span></p>' +
      '<p style="margin: 0; font-style: italic;">此邮件由系统自动发送，请勿回复。<br>' +
      '<span style="font-size: 0.8em; opacity: 0.8;">This email is automatically sent by the system, please do not reply.</span></p>' +
      '</div></div></div>';

    return emailBody;

  } catch (error) {
    console.error('生成工序周报邮件内容时出错:', error);
    return '<p>邮件内容生成失败，请检查系统日志。</p>';
  }
}

// ========== 通用工具 ==========

function getFieldIndexes(headers) {
  const indexes = {};
  headers.forEach(function(header, index) {
    indexes[header] = index;
  });
  return indexes;
}

function calculateOverdueDays(assignDate) {
  try {
    if (!assignDate || !(assignDate instanceof Date)) {
      return 0;
    }
    const today = new Date();
    const timeDiff = today.getTime() - assignDate.getTime();
    const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));
    return Math.max(0, daysDiff);
  } catch (error) {
    console.error('计算超期天数时出错:', error);
    return 0;
  }
}

function isFailureReportUploaded(record) {
  try {
    if (!record) {
      return { isUploaded: false, reason: '记录为空', uploadDate: null, attachment: null };
    }

    const hasUploadDate = record.uploadDate && record.uploadDate.toString().trim() !== '';
    const hasAttachment = record.attachment && record.attachment.toString().trim() !== '';

    let isUploaded = false;
    let reason = '';

    if (hasUploadDate && hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有上传日期和附件';
    } else if (hasUploadDate && !hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有上传日期但无附件';
    } else if (!hasUploadDate && hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有附件但无上传日期';
    } else {
      isUploaded = false;
      reason = '未上传 - 无上传日期和附件';
    }

    return {
      isUploaded: isUploaded,
      reason: reason,
      uploadDate: record.uploadDate || null,
      attachment: record.attachment || null
    };

  } catch (error) {
    console.error('判断故障报告上传状态时出错:', error);
    return { isUploaded: false, reason: '判断出错', uploadDate: null, attachment: null };
  }
}

function getUnuploadedFailureReports(failureData) {
  try {
    if (!failureData || !Array.isArray(failureData)) return [];
    const unuploadedReports = failureData.filter(function(record) {
      return !isFailureReportUploaded(record).isUploaded;
    });
    console.log('📊 找到 ' + unuploadedReports.length + ' 条未上传的故障报告');
    return unuploadedReports;
  } catch (error) {
    console.error('获取未上传故障报告时出错:', error);
    return [];
  }
}

function getUploadedFailureReports(failureData) {
  try {
    if (!failureData || !Array.isArray(failureData)) return [];
    const uploadedReports = failureData.filter(function(record) {
      return isFailureReportUploaded(record).isUploaded;
    });
    console.log('📊 找到 ' + uploadedReports.length + ' 条已上传的故障报告');
    return uploadedReports;
  } catch (error) {
    console.error('获取已上传故障报告时出错:', error);
    return [];
  }
}

// ========== 测试函数 ==========

function testWeeklyReminder() {
  weeklyFailureReportReminder();
}
