// V20260608.1 — 故障报告跟进项提醒
// 功能：每日检查 Failure_Report_followup 表，提醒责任人（临期/逾期）和验证人（待验证）
// 依赖：22 - 故障报告周报.js（共享函数：getUserIDLookupMaps, extractNameFromPersonField, extractEmailFromPersonField）

// ========== 管理员邮箱（从 userID BH列"项目跟进权限管理"读取） ==========

function getAdminEmails() {
  try {
    const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID)
      .getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet) { console.error('权限表未找到'); return []; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return [];

    // Row 1 = category, Row 2 = headers, Row 3+ = data
    const headers = data[1];
    const bhIndex = headers.findIndex(function(h) { return h === '项目跟进权限管理'; });
    const gmailIndex = headers.findIndex(function(h) { return h === 'GMail'; });
    if (bhIndex === -1 || gmailIndex === -1) {
      console.error('权限表缺少必要列（项目跟进权限管理 或 GMail）');
      return [];
    }

    const admins = [];
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][bhIndex] || '').trim() === '管理员') {
        const email = String(data[i][gmailIndex] || '').trim();
        if (email) admins.push(email);
      }
    }
    console.log('管理员邮箱: ' + (admins.join(', ') || '无'));
    return admins;
  } catch (error) {
    console.error('获取管理员邮箱时出错:', error);
    return [];
  }
}

// ========== 主函数：每日跟进项目提醒 ==========

function dailyFollowUpReminder(e) {
  const trigger = e ? '定时' : '手动';
  try {
    console.log('=== 开始执行故障报告跟进项目提醒 ===');

    const followUpData = getFollowUpData();
    if (!followUpData || followUpData.length === 0) {
      console.log('没有找到跟进项目数据');
      writeLog('dailyFollowUpReminder', '成功', '没有找到跟进项目数据，跳过执行', trigger, '');
      return;
    }

    // ownerMap[email] = { dueSoon: [], overdue: [] }
    const ownerMap = {};
    // verifierMap[email] = []
    const verifierMap = {};

    followUpData.forEach(function(item) {
      const verify = item.status || '';

      if (verify.indexOf('已通过') !== -1) return;

      if (verify.indexOf('未通过') !== -1 || verify.indexOf('NA') !== -1) {
        const email = extractEmailFromPersonField(item.owner);
        if (!email) return;
        const days = calcDaysUntilDue(item.dueDate);
        if (days < 0) {
          if (!ownerMap[email]) ownerMap[email] = { dueSoon: [], overdue: [] };
          ownerMap[email].overdue.push(item);
        } else if (days <= 2) {
          if (!ownerMap[email]) ownerMap[email] = { dueSoon: [], overdue: [] };
          ownerMap[email].dueSoon.push(item);
        }
        // days > 2: 不提醒

      } else if (verify.indexOf('未验证') !== -1) {
        const email = extractEmailFromPersonField(item.verifier);
        if (!email) return;
        if (!verifierMap[email]) verifierMap[email] = [];
        verifierMap[email].push(item);
      }
    });

    // 获取 userID 映射表（用于 CC: LM + 工序管理员）
    const { nameToUser, processToAdmins, nameToProcess } = getUserIDLookupMaps();

    // 辅助函数：从人员列表提取 CC（LM + 工序管理员）
    function collectCCFromPersons(items, personField) {
      const lms = [];
      const procSet = {};
      items.forEach(function(item) {
        const rawName = String(item[personField] || '').trim();
        const name = extractNameFromPersonField(rawName);
        if (name) {
          const user = nameToUser[name];
          if (user && user.lineManager) lms.push(user.lineManager);
          const proc = nameToProcess[name];
          if (proc) procSet[proc] = true;
        }
      });
      const admins = [];
      Object.keys(procSet).forEach(function(p) {
        if (processToAdmins[p]) admins.push.apply(admins, processToAdmins[p]);
      });
      return { lms: lms, admins: admins };
    }

    // 合并发送责任人提醒（一封邮件发给所有责任人）
    let ownerSent = 0;
    const allOwnerEmails = [];
    const allOwnerDueSoon = [];
    const allOwnerOverdue = [];

    for (const email in ownerMap) {
      const { dueSoon, overdue } = ownerMap[email];
      if (dueSoon.length === 0 && overdue.length === 0) continue;
      allOwnerEmails.push(email);
      allOwnerDueSoon.push.apply(allOwnerDueSoon, dueSoon.map(function(item) { return extendItem(item, '_ownerEmail', email); }));
      allOwnerOverdue.push.apply(allOwnerOverdue, overdue.map(function(item) { return extendItem(item, '_ownerEmail', email); }));
    }

    if (allOwnerEmails.length > 0) {
      const allOwnerItems = allOwnerDueSoon.concat(allOwnerOverdue);
      const ccOwner = collectCCFromPersons(allOwnerItems, 'owner');
      const ccAll = ccOwner.lms.concat(ccOwner.admins);
      ccAll.push('kelland_zhao@colpal.com');
      const ccSeen = {};
      const ccList = ccAll.filter(function(e) { if (ccSeen[e]) return false; ccSeen[e] = true; return true; }).filter(Boolean);

      const subject = allOwnerOverdue.length > 0
        ? '【逾期提醒】故障报告跟进项目逾期 / Follow-up Items Overdue'
        : '【临期提醒】故障报告跟进项目即将到期 / Follow-up Items Due Soon';
      GmailApp.sendEmail(allOwnerEmails.join(','), subject,
        '请使用支持HTML的邮件客户端查看此邮件。', {
          htmlBody: generateMergedOwnerEmailContent(allOwnerDueSoon, allOwnerOverdue),
          name: '故障报告提醒系统',
          cc: ccList.join(',')
        });
      ownerSent = 1;
      console.log('✅ 责任人合并提醒 → To: ' + allOwnerEmails.join(', ') + ' (临期:' + allOwnerDueSoon.length + ' 逾期:' + allOwnerOverdue.length + ') CC: ' + ccList.join(', '));
    }

    // 合并发送验证人提醒（一封邮件发给所有验证人）
    let verifierSent = 0;
    const allVerifierEmails = [];
    const allVerifierItems = [];

    for (const email in verifierMap) {
      const items = verifierMap[email];
      if (items.length === 0) continue;
      allVerifierEmails.push(email);
      allVerifierItems.push.apply(allVerifierItems, items.map(function(item) { return extendItem(item, '_verifierEmail', email); }));
    }

    if (allVerifierEmails.length > 0) {
      const ccVerifier = collectCCFromPersons(allVerifierItems, 'verifier');
      const ccAll = ccVerifier.lms.concat(ccVerifier.admins);
      ccAll.push('kelland_zhao@colpal.com');
      const ccSeen = {};
      const ccList = ccAll.filter(function(e) { if (ccSeen[e]) return false; ccSeen[e] = true; return true; }).filter(Boolean);

      GmailApp.sendEmail(allVerifierEmails.join(','),
        '【验证提醒】故障报告跟进项目待验证 / Follow-up Items Pending Verification',
        '请使用支持HTML的邮件客户端查看此邮件。', {
          htmlBody: generateMergedVerifierEmailContent(allVerifierItems),
          name: '故障报告提醒系统',
          cc: ccList.join(',')
        });
      verifierSent = 1;
      console.log('✅ 验证人合并提醒 → To: ' + allVerifierEmails.join(', ') + ' (' + allVerifierItems.length + '条待验证) CC: ' + ccList.join(', '));
    }

    writeLog('dailyFollowUpReminder', '成功', '责任人提醒 ' + ownerSent + ' 封，验证人提醒 ' + verifierSent + ' 封，共处理 ' + followUpData.length + ' 条跟进项目', trigger, '');
    console.log('=== 故障报告跟进项目提醒执行完成 ===');

  } catch (error) {
    console.error('故障报告跟进项目提醒执行出错:', error);
    writeLog('dailyFollowUpReminder', '失败', error.message, trigger, '');
  }
}

// ========== 数据读取 ==========

function getFollowUpData() {
  try {
    const sheet = SpreadsheetApp.openById(FR_CONFIG.SPREADSHEET_ID)
      .getSheetByName(FR_CONFIG.FOLLOWUP_SHEET_NAME);
    if (!sheet) { console.error('跟进项目表未找到'); return []; }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const idx = getFieldIndexes(data[0]);
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idx['序号 / followup_id']]) continue;
      result.push({
        id:          row[idx['序号 / followup_id']]               || '',
        reportNo:    row[idx['故障报告编号 / failure_report_no']]  || '',
        paType:      row[idx['行动类型 / pa_type']]                || '',
        paPlan:      row[idx['预防行动 / pa_plan']]                || '',
        owner:   String(row[idx['责任人 / pa_who']]                || ''),
        dueDate:     row[idx['完成时间 / pa_when']]                || '',
        verifier:String(row[idx['验证人 / pa_verifier']]           || ''),
        status:      row[idx['状态 / status']]                     || '',
        notes:       row[idx['跟进内容 / follow_up_notes']]        || ''
      });
    }

    console.log('✅ 读取到 ' + result.length + ' 条跟进项目数据');
    return result;
  } catch (error) {
    console.error('读取跟进项目数据时出错:', error);
    return [];
  }
}

// ========== 工具函数 ==========

function extendItem(item, key, value) {
  var newItem = {};
  for (var k in item) {
    if (item.hasOwnProperty(k)) newItem[k] = item[k];
  }
  newItem[key] = value;
  return newItem;
}

function calcDaysUntilDue(dueDate) {
  try {
    const due = (dueDate instanceof Date) ? new Date(dueDate) : new Date(dueDate);
    if (isNaN(due.getTime())) return 999;
    due.setHours(0, 0, 0, 0);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.round((due - today) / (1000 * 3600 * 24));
  } catch (e) {
    return 999;
  }
}

function formatFollowUpDate(value) {
  if (!value) return '';
  const d = (value instanceof Date) ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : formatVariableAsDate(d);
}

// ========== 责任人跟进提醒邮件（合并版） ==========

function generateMergedOwnerEmailContent(allDueSoon, allOverdue) {
  const hasOverdue = allOverdue.length > 0;
  const today = formatVariableAsDate(new Date());
  const accentColor = hasOverdue ? '#f44336' : '#f39c12';
  const darkColor = hasOverdue ? '#d32f2f' : '#e65100';
  const bgColor = hasOverdue ? '#ffebee' : '#fff8e1';

  function buildTable(items, isOverdue) {
    if (items.length === 0) return '';

    const headerGrad = isOverdue
      ? 'linear-gradient(135deg,#f44336,#d32f2f)'
      : 'linear-gradient(135deg,#f39c12,#e67e22)';
    const rowBgAlt = isOverdue ? '#fff5f5' : '#fffbf0';

    let rows = '';
    items.forEach(function(item, i) {
      const days = calcDaysUntilDue(item.dueDate);
      const ownerName = extractNameFromPersonField(item.owner);
      const badge = isOverdue
        ? '<div style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;"><span style="display:block;">[逾期] ' + Math.abs(days) + '天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Overdue</span></div>'
        : '<div style="background:linear-gradient(135deg,#f39c12,#e67e22);color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;"><span style="display:block;">还剩 ' + days + '天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Left</span></div>';
      rows += '<tr style="background-color:' + (i % 2 === 0 ? rowBgAlt : '#ffffff') + ';">' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + ownerName + '<br><span style="font-size:11px;color:#95a5a6;">' + (item._ownerEmail || '') + '</span></td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">' + (item.reportNo || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:220px;word-wrap:break-word;">' + (item.paPlan || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (formatFollowUpDate(item.dueDate) || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (item.status || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + badge + '</td>' +
        '</tr>';
    });

    return '<div style="overflow-x:auto;margin-bottom:20px;">' +
      '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">' +
      '<thead><tr style="background:' + headerGrad + ';color:white;">' +
      '<th style="padding:10px;text-align:left;font-weight:600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">期限<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span></th>' +
      '<th style="padding:10px;text-align:center;font-weight:600;">' + (isOverdue ? '逾期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span>' : '剩余天数<br><span style="font-size:0.8em;opacity:0.9;">Days Left</span>') + '</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div>';
  }

  let body = '<div style="font-family:Arial,sans-serif;max-width:900px;margin:0 auto;background-color:#f8f9fa;padding:20px;">' +
    '<div style="background:' + bgColor + ';border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid ' + accentColor + ';">' +
    '<h2 style="color:' + darkColor + ';text-align:center;margin-bottom:20px;border-bottom:3px solid ' + accentColor + ';padding-bottom:10px;">' +
    (hasOverdue ? '【逾期提醒】 故障报告跟进项目逾期' : '【临期提醒】 故障报告跟进项目即将到期') + '<br>' +
    '<span style="font-size:0.8em;">' + (hasOverdue ? 'Follow-up Items Overdue Reminder' : 'Follow-up Items Due Soon Reminder') + '</span>' +
    '</h2>' +
    '<p style="font-size:16px;line-height:1.6;color:' + darkColor + ';">' +
    '您好！（' + today + '）以下故障报告跟进项目需要处理：<br>' +
    '<span style="font-size:0.9em;opacity:0.85;">Hello! The following follow-up items require your attention (' + today + '):</span>' +
    '</p></div>';

  if (allOverdue.length > 0) {
    body += '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">' +
      '<h3 style="color:#d32f2f;border-bottom:2px solid #f44336;padding-bottom:10px;margin-bottom:20px;">' +
      '【逾期】 已逾期跟进项目 Overdue Items (' + allOverdue.length + '条)</h3>' +
      buildTable(allOverdue, true) + '</div>';
  }

  if (allDueSoon.length > 0) {
    body += '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">' +
      '<h3 style="color:#e65100;border-bottom:2px solid #f39c12;padding-bottom:10px;margin-bottom:20px;">' +
      '【临期】 即将到期跟进项目 Due Soon Items (' + allDueSoon.length + '条)</h3>' +
      buildTable(allDueSoon, false) + '</div>';
  }

  body += '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">' +
    '<div style="text-align:center;color:' + darkColor + ';font-size:14px;line-height:1.6;">' +
    '<p style="margin-bottom:10px;font-weight:600;">请及时处理以上跟进项目！<br>' +
    '<span style="font-size:0.9em;opacity:0.85;">Please handle the above follow-up items promptly!</span></p>' +
    '<p style="margin:0;font-style:italic;">此邮件由系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.7;">This email is automatically sent by the system, please do not reply.</span></p>' +
    '</div></div></div>';

  return body;
}

// ========== 验证人验证提醒邮件（合并版） ==========

function generateMergedVerifierEmailContent(allItems) {
  const today = formatVariableAsDate(new Date());

  function groupByEmail(items) {
    const map = {};
    items.forEach(function(item) {
      const email = item._verifierEmail;
      if (!map[email]) map[email] = [];
      map[email].push(item);
    });
    return map;
  }

  function buildPersonSection(email, items) {
    const name = extractNameFromPersonField((items[0] || {}).verifier || '');
    let rows = '';
    items.forEach(function(item, i) {
      rows += '<tr style="background-color:' + (i % 2 === 0 ? '#f0f4ff' : '#ffffff') + ';">' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">' + (item.id || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (item.reportNo || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (item.paType || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:220px;word-wrap:break-word;">' + (item.paPlan || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:200px;word-wrap:break-word;">' + (item.notes || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">' + (formatFollowUpDate(item.dueDate) || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (extractNameFromPersonField(item.owner) || '') + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + (item.status || '') + '</td>' +
        '</tr>';
    });

    return '<div style="margin-bottom:15px;">' +
      '<h4 style="color:#283593;margin-bottom:8px;padding:8px 12px;background:#f0f4ff;border-radius:6px;border-left:3px solid #3f51b5;">' +
      name + ' (' + email + ') - ' + items.length + '条</h4>' +
      '<div style="overflow-x:auto;">' +
      '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">' +
      '<thead><tr style="background:linear-gradient(135deg,#3f51b5,#283593);color:white;">' +
      '<th style="padding:10px;text-align:left;font-weight:600;">跟进编号<br><span style="font-size:0.8em;opacity:0.9;">Follow-up ID</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">行动类型<br><span style="font-size:0.8em;opacity:0.9;">Type</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">跟进内容<br><span style="font-size:0.8em;opacity:0.9;">Follow-up Notes</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">完成时间<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>' +
      '<th style="padding:10px;text-align:left;font-weight:600;">状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span></th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div></div>';
  }

  const groups = groupByEmail(allItems);

  let body = '<div style="font-family:Arial,sans-serif;max-width:900px;margin:0 auto;background-color:#f8f9fa;padding:20px;">' +
    '<div style="background:#e8eaf6;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid #3f51b5;">' +
    '<h2 style="color:#283593;text-align:center;margin-bottom:20px;border-bottom:3px solid #3f51b5;padding-bottom:10px;">' +
    '【验证提醒】 故障报告跟进项目待验证<br>' +
    '<span style="font-size:0.8em;">Follow-up Items Pending Verification</span>' +
    '</h2>' +
    '<p style="font-size:16px;line-height:1.6;color:#283593;">' +
    '您好！（' + today + '）以下跟进项目等待验证：<br>' +
    '<span style="font-size:0.9em;opacity:0.85;">Hello! The following follow-up items are pending verification (' + today + '):</span>' +
    '</p></div>' +
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">' +
    '<h3 style="color:#283593;border-bottom:2px solid #3f51b5;padding-bottom:10px;margin-bottom:20px;">' +
    '【详情】 待验证跟进项目 Pending Verification Items (' + allItems.length + '条)</h3>';

  for (const email in groups) {
    body += buildPersonSection(email, groups[email]);
  }

  body += '</div>' +
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">' +
    '<div style="text-align:center;color:#283593;font-size:14px;line-height:1.6;">' +
    '<p style="margin-bottom:10px;font-weight:600;">请及时对以上项目进行验证！<br>' +
    '<span style="font-size:0.9em;opacity:0.85;">Please verify the above items promptly!</span></p>' +
    '<p style="margin:0;font-style:italic;">此邮件由系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.7;">This email is automatically sent by the system, please do not reply.</span></p>' +
    '</div></div></div>';

  return body;
}

// ========== 测试函数 ==========

function testFollowUpReminder() {
  dailyFollowUpReminder();
}
