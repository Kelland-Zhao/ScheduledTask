// V20260603.01 — 辅导与反馈提醒
// 入口：sendCoachingReminder（每日 08:15 定时 or 手动）
// 逻辑：读取辅导与反馈表，按工序分组，向F列工程师发送未缴提醒邮件

// ========== 常量 ==========
const CF_SPREADSHEET_ID = '1XyYXB6RCkgZGaAuVv73b1gHOx81w1oOb50QtqTX_Xyk';
const CF_SHEET_NAME = '1-12 月辅导与反馈';
const CF_DATA_START_ROW = 6;          // 表头在第5行，数据从第6行开始
const CF_JUDY_EMAIL = 'judy_zhang@colpal.com';
const CF_USERID_EMAIL_COL = 9;        // J列（0-indexed）GMail
const CF_USERID_SUPERVISOR_COL = 60;  // BI列（0-indexed）直线上级邮箱

// ========== 主函数 ==========
function sendCoachingReminder(e) {
  const trigger = e ? '定时' : '手动';
  try {
    const userMap = _cf_buildUserMap();
    const pending = _cf_getPendingRows();

    if (pending.length === 0) {
      writeLog('sendCoachingReminder', '跳过', '所有记录均已提交，无需提醒', trigger, '');
      return;
    }

    // 按工序分组
    const byProcess = {};
    pending.forEach(function(r) {
      const proc = r.process || '未知';
      if (!byProcess[proc]) byProcess[proc] = [];
      byProcess[proc].push(r);
    });

    const sent = [];
    const warnings = [];

    Object.keys(byProcess).forEach(function(proc) {
      const rows = byProcess[proc];
      const result = _cf_sendProcessEmail(proc, rows, userMap, trigger);
      if (result.success) {
        sent.push(proc + '(' + rows.length + '人)');
      } else {
        warnings.push(proc + ':' + result.reason);
      }
    });

    const detail = '已发: ' + (sent.join('|') || '无') +
      (warnings.length ? ' | 警告: ' + warnings.join('|') : '');
    writeLog('sendCoachingReminder', '成功', detail, trigger, '共' + pending.length + '条未缴');

  } catch (err) {
    writeLog('sendCoachingReminder', '失败', err.message, trigger, '');
    throw err;
  }
}

// ========== 读取未缴行 ==========
function _cf_getPendingRows() {
  const sheet = SpreadsheetApp.openById(CF_SPREADSHEET_ID).getSheetByName(CF_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < CF_DATA_START_ROW) return [];

  const data = sheet.getRange(CF_DATA_START_ROW, 1, lastRow - CF_DATA_START_ROW + 1, 8).getValues();
  const rows = [];

  data.forEach(function(row) {
    const sapNo = String(row[0] || '').trim();
    const name  = String(row[1] || '').trim();
    const month = String(row[4] || '').trim();
    const coach = String(row[5] || '').trim();      // F: 辅导与反馈表
    const submitted = String(row[6] || '').trim();  // G: 已经提交
    const process   = String(row[7] || '').trim();  // H: 工序

    if (!sapNo && !name) return;   // 空行
    if (!coach) return;            // 无主管信息，无法提醒
    if (month < '202604') return;  // 202604 之前数据豁免
    if (submitted !== '') return;  // 非空 = 已提交或豁免（Y / 退休 / 考勤工伤...）

    rows.push({ sapNo: sapNo, name: name, month: month, coach: coach, process: process });
  });

  return rows;
}

// ========== 构建 userID 姓名→邮箱映射 ==========
function _cf_buildUserMap() {
  const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  // 读取 A~BI（61列），从第3行开始（跳过2行表头）
  const data = sheet.getRange(3, 1, lastRow - 2, 61).getValues();
  const map = {};

  data.forEach(function(row) {
    const rawName = String(row[1] || '').trim();
    if (!rawName) return;
    const email = String(row[CF_USERID_EMAIL_COL] || '').trim().toLowerCase();
    const supervisorEmail = String(row[CF_USERID_SUPERVISOR_COL] || '').trim().toLowerCase();
    const key = _cf_normalizeName(rawName);
    if (key && !map[key]) {
      map[key] = { email: email, supervisorEmail: supervisorEmail };
    }
  });

  return map;
}

// ========== 姓名预处理（用于 userID 匹配）==========
function _cf_normalizeName(name) {
  let s = String(name || '').trim();
  s = s.replace(/-[A-Za-z0-9]+$/, ''); // 去掉 "-Z" 类后缀
  s = s.replace(/_[A-Za-z0-9_]+$/, ''); // 去掉 "_QC" 类后缀
  return s.trim().toLowerCase();
}

// ========== 发送单工序邮件 ==========
function _cf_sendProcessEmail(proc, rows, userMap, trigger) {
  // 收集本工序不重复的主管
  const coachKeys = {};
  rows.forEach(function(r) {
    const key = _cf_normalizeName(r.coach);
    if (!coachKeys[key]) coachKeys[key] = r.coach; // key → 原始姓名
  });

  const toEmails = [];
  const ccSupEmails = [];
  const notFound = [];

  Object.keys(coachKeys).forEach(function(key) {
    const info = userMap[key];
    if (!info || !info.email) {
      notFound.push(coachKeys[key]);
      return;
    }
    if (toEmails.indexOf(info.email) < 0) toEmails.push(info.email);
    if (info.supervisorEmail && ccSupEmails.indexOf(info.supervisorEmail) < 0) {
      ccSupEmails.push(info.supervisorEmail);
    }
  });

  if (notFound.length > 0) {
    writeLog('sendCoachingReminder', '警告',
      proc + ' 工序部分工程师未找到邮箱: ' + notFound.join(','), trigger, '');
  }

  if (toEmails.length === 0) {
    return { success: false, reason: '所有工程师均未匹配到邮箱(' + notFound.join(',') + ')' };
  }

  // CC：直线领导 + Judy + Kelland
  const ccAll = ccSupEmails.slice();
  if (ccAll.indexOf(CF_JUDY_EMAIL) < 0) ccAll.push(CF_JUDY_EMAIL);
  if (ccAll.indexOf(GMAIL_CC) < 0) ccAll.push(GMAIL_CC);

  const subject = '[辅导与反馈提醒] ' + proc + ' 工序 - 待提交 ' + rows.length +
    ' 份 | Coaching Form Reminder - ' + proc + ' - ' + rows.length + ' Pending';
  const htmlBody = _cf_buildHtml(proc, rows, notFound);

  sendMail(toEmails.join(','), subject, htmlBody, ccAll.join(','));
  return { success: true };
}

// ========== 构建邮件 HTML ==========
function _cf_buildHtml(proc, rows, notFound) {
  const today = formatVariableAsDate(new Date());

  // 按月份→姓名排序
  rows.sort(function(a, b) {
    if (a.month !== b.month) return a.month < b.month ? -1 : 1;
    return a.name < b.name ? -1 : 1;
  });

  const tableRows = rows.map(function(r) {
    return [r.sapNo, r.name, r.month, r.coach];
  });
  const table = buildHtmlTable(
    ['SAP No', '姓名 / Name', '月份 / Month', '辅导工程师 / Coach'],
    tableRows, '#E60012'
  );

  let notFoundHtml = '';
  if (notFound.length > 0) {
    notFoundHtml = '<p style="color:#e60012;font-size:12px">⚠️ 以下工程师邮箱未在系统中找到，请手动通知 / Engineers not found in system: ' +
      notFound.join('、') + '</p>';
  }

  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333">' +
    '<h2 style="color:#E60012">辅导与反馈表未提交提醒<br>' +
    '<span style="font-size:14px;font-weight:normal">Coaching &amp; Feedback Form Pending Reminder</span></h2>' +
    '<p>发送日期 Date: <strong>' + today + '</strong>' +
    '&nbsp;&nbsp;|&nbsp;&nbsp;工序 Process: <strong>' + proc + '</strong>' +
    '&nbsp;&nbsp;|&nbsp;&nbsp;待提交数量 Pending: <strong>' + rows.length + '</strong></p>' +
    '<p>请尽快完成辅导与反馈表的提交。<br>' +
    'Please complete the coaching &amp; feedback form submission as soon as possible.</p>' +
    table + notFoundHtml +
    '<p style="color:#888;font-size:12px;margin-top:20px">' +
    '此邮件由系统自动发送 / This is an automated email</p>' +
    '</div>';
}
