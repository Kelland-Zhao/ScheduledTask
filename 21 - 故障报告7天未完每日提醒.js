// V20260608.1 — 故障报告7天未完每日提醒
// 功能：每日检测超期>=7天且未上传的故障报告，按工序分组发送紧急提醒
// 依赖：22 - 故障报告周报.js（共享函数：getOverdueFailureReportData, getUserIDLookupMaps 等）

// ========== 主函数：每日超期故障报告提醒 ==========

function dailyOverdueFailureReportReminder(e) {
  const trigger = e ? '定时' : '手动';
  try {
    console.log('=== 开始执行每日超期故障报告提醒 ===');

    const overdueData = getOverdueFailureReportData();
    if (!overdueData || overdueData.length === 0) {
      console.log('没有找到超期故障报告数据');
      writeLog('dailyOverdueFailureReportReminder', '成功', '没有超期故障报告数据，跳过执行', trigger, '');
      return;
    }

    const { nameToUser, processToAdmins } = getUserIDLookupMaps();
    if (Object.keys(nameToUser).length === 0) {
      console.warn('⚠️ userID 权限表无有效数据，无法发送日报');
      writeLog('dailyOverdueFailureReportReminder', '警告', 'userID 权限表无有效数据，跳过执行', trigger, '');
      return;
    }

    sendDailyOverdueRemindersByProcess(overdueData, nameToUser, processToAdmins, trigger);

    console.log('=== 每日超期故障报告提醒执行完成 ===');

  } catch (error) {
    console.error('每日超期故障报告提醒执行出错:', error);
    writeLog('dailyOverdueFailureReportReminder', '失败', error.message, trigger, '');
  }
}

// ========== 日报邮件发送（按工序） ==========

function sendDailyOverdueRemindersByProcess(overdueData, nameToUser, processToAdmins, trigger) {
  try {
    console.log('=== 开始按工序发送日报（超期）邮件 ===');
    console.log('📊 超期未上传故障报告: ' + overdueData.length + ' 条');

    // 按工序分组（IM -> INJ 映射）
    var processGroups = {};
    overdueData.forEach(function(record) {
      var mappedProcess = mapFailureProcessToUserID(record.process);
      if (!mappedProcess) return;
      if (!processGroups[mappedProcess]) processGroups[mappedProcess] = [];
      processGroups[mappedProcess].push(record);
    });

    var processes = Object.keys(processGroups);
    console.log('📊 涉及工序: [' + processes.join(', ') + ']');

    if (processes.length === 0) {
      console.log('⚠️ 没有需要发送的工序');
      writeLog('dailyOverdueFailureReportReminder', '成功', '没有需要发送的工序', trigger, '');
      return;
    }

    var sentCount = 0;

    for (var p = 0; p < processes.length; p++) {
      var process = processes[p];
      var records = processGroups[process];
      var displayName = PROCESS_DISPLAY_NAMES[process] || process;

      // 提取唯一责任人姓名
      var ownerRawNames = [];
      var seen = {};
      records.forEach(function(r) {
        var raw = String(r.owner || '').trim();
        if (raw && !seen[raw]) { seen[raw] = true; ownerRawNames.push(raw); }
      });
      var ownerLookupNames = ownerRawNames.map(function(raw) { return extractNameFromPersonField(raw); });
      console.log('\n🔍 工序 ' + process + '(' + displayName + '): ' + records.length + '条超期报告, 责任人: [' + ownerLookupNames.join(', ') + ']');

      // 查找责任人邮箱和 line manager
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
        writeLog('dailyOverdueFailureReportReminder', '警告', '工序 ' + process + '(' + displayName + '): 无有效责任人邮箱，跳过 (' + records.length + '条报告)', trigger, '');
        continue;
      }

      // CC 列表
      var processAdmins = processToAdmins[process] || [];
      var ccAll = ccLineMgrs.concat(processAdmins);
      ccAll.push('kelland_zhao@colpal.com');
      var ccSeen = {};
      var ccList = ccAll.filter(function(e) { if (ccSeen[e]) return false; ccSeen[e] = true; return true; }).filter(Boolean);

      // 生成邮件
      var subject = '【紧急】超期故障报告提醒 - ' + displayName + '(' + process + ') / Overdue Failure Report - ' + displayName + '(' + process + ')';
      var htmlBody = generateDailyOverdueProcessEmailContent(process, displayName, records);

      GmailApp.sendEmail(toEmails.join(','), subject, '请使用支持HTML的邮件客户端查看此邮件。', {
        htmlBody: htmlBody,
        name: '故障报告提醒系统',
        cc: ccList.join(',')
      });

      sentCount++;
      console.log('✅ 日报已发送 → 工序: ' + process + '(' + displayName + ')');
      console.log('   To: ' + toEmails.join(', '));
      console.log('   CC: ' + ccList.join(', '));
      console.log('   超期报告: ' + records.length + '条');

      writeLog('dailyOverdueFailureReportReminder', '成功', '工序 ' + process + '(' + displayName + '): 发送成功 → To:' + toEmails.length + '人 CC:' + ccList.length + '人 报告:' + records.length + '条', trigger, '');
    }

    console.log('\n=== 日报发送完成，共 ' + sentCount + ' 封邮件 ===');

  } catch (error) {
    console.error('按工序发送日报邮件时出错:', error);
    writeLog('dailyOverdueFailureReportReminder', '失败', error.message, trigger, '');
  }
}

// ========== 日报邮件 HTML 模板（超期红色主题） ==========

function generateDailyOverdueProcessEmailContent(processCode, displayName, records) {
  try {
    var formattedDate = formatVariableAsDate(new Date());
    var overdueDays = FR_CONFIG.OVERDUE_DAYS;

    var emailBody = '<div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; background-color: #f8f9fa; padding: 20px;">' +
      '<div style="background-color: #ffebee; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px; border-left: 5px solid #f44336;">' +
      '<h2 style="color: #d32f2f; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #f44336; padding-bottom: 10px;">' +
      '【紧急】超期未上传故障报告提醒 - ' + displayName + '(' + processCode + ') 工序<br>' +
      '<span style="font-size: 0.8em; color: #d32f2f;">Overdue Unuploaded Failure Report - ' + displayName + '(' + processCode + ') Process</span>' +
      '</h2>' +
      '<p style="color: #d32f2f; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">' +
      '紧急通知！以下 <strong>' + displayName + '</strong> 工序的故障报告已超期 <strong>≥' + overdueDays + '天</strong> 且未上传，请立即处理：<br>' +
      '<span style="font-size: 0.9em; color: #d32f2f; opacity: 0.8;">Urgent! The following <strong>' + displayName + '</strong> process failure reports are overdue ≥' + overdueDays + ' days and unuploaded:</span>' +
      '</p></div>';

    if (records.length > 0) {
      emailBody += '<div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">' +
        '<h3 style="color: #d32f2f; margin-bottom: 20px; border-bottom: 2px solid #f44336; padding-bottom: 10px;">' +
        '[详情] 超期未上传故障报告 Overdue Unuploaded Reports (' + records.length + '条)</h3>' +
        '<div style="overflow-x: auto;">' +
        '<table style="width: 100%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">' +
        '<thead><tr style="background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%); color: white;">' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">编号<br><span style="font-size:0.8em;opacity:0.9;">ID</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">机台号<br><span style="font-size:0.8em;opacity:0.9;">Machine</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">问题描述<br><span style="font-size:0.8em;opacity:0.9;">Description</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">车间<br><span style="font-size:0.8em;opacity:0.9;">Workshop</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">分配日期<br><span style="font-size:0.8em;opacity:0.9;">Assign Date</span></th>' +
        '<th style="padding: 12px; text-align: left; font-weight: 600;">超期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span></th>' +
        '</tr></thead><tbody>';

      records.forEach(function(record, i) {
        var rowStyle = i % 2 === 0 ? 'background-color: #fff5f5;' : 'background-color: #ffffff;';
        emailBody += '<tr style="' + rowStyle + '">' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">' + (record.id || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (record.machineId || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:200px;word-wrap:break-word;">' + (record.description || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (record.workshop || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (record.owner || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (formatVariableAsDate(record.assignDate) || '') + '</td>' +
          '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' +
          '<div style="display:inline-block;text-align:center;"><div style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;padding:6px 12px;border-radius:16px;font-weight:600;font-size:12px;box-shadow:0 2px 6px rgba(244,67,54,0.3);display:inline-block;min-width:80px;">' +
          '<span style="display:block;">[超期] ' + record.overdueDays + '天</span><span style="display:block;font-size:10px;opacity:0.9;">Days</span></div></div>' +
          '</td></tr>';
      });

      emailBody += '</tbody></table></div>' +
        '<div style="margin-top:15px;text-align:right;"><span style="color:#7f8c8d;font-size:14px;">共 ' + records.length + ' 条 | Total: ' + records.length + ' reports</span></div></div>';
    }

    emailBody += '<div style="background-color: #ffebee; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px;">' +
      '<div style="text-align: center; color: #d32f2f; font-size: 14px; line-height: 1.6;">' +
      '<p style="margin-bottom: 10px; font-weight: 600;">【重要】请立即处理这些超期故障报告！<br>' +
      '<span style="font-size: 0.9em; opacity: 0.8;">Please handle these overdue failure reports immediately!</span></p>' +
      '<p style="margin: 0; font-style: italic;">此邮件由系统自动发送，请勿回复。<br>' +
      '<span style="font-size: 0.8em; opacity: 0.7;">This email is automatically sent by the system, please do not reply.</span></p>' +
      '</div></div></div>';

    return emailBody;

  } catch (error) {
    console.error('生成工序日报邮件内容时出错:', error);
    return '<p>邮件内容生成失败，请检查系统日志。</p>';
  }
}

// ========== 测试函数 ==========

function testDailyOverdueReminder() {
  dailyOverdueFailureReportReminder();
}
