// ============================================================
// 30 - 报警监控.js — 经理报警监控Report（邮件）
// 入口：sendManagerAlertReport
// 共用：getAlertSheetsConfig（29号 微信提醒也调用）
// 菜单：发送IoT经理报警监控Report
// ============================================================

// 各报警Sheet配置（statusCol=-1表示该Sheet无状态列，跳过未跟进统计）
function getAlertSheetsConfig() {
  let ss = SpreadsheetApp.openById(ALERT_SPREADSHEET_ID);
  return [
    {
      label:          'IoT/Opera周期报警',
      sheet:          ss.getSheetByName("IOT/Opera邮件反馈格式"),
      totalCols:      20,
      dateCol:        4,   // 日期（yyyy-MM-dd）
      timeCol:        5,   // 报警时间（HH:mm:ss）
      dateIsDatetime: false,
      statusCol:      16,  // 状态
      workshopCol:    2,   // 车间
      machineCol:     8,   // 机台号
      typeCol:        9,   // 参数类型
      diffCol:        12,  // 差异（负值=超下限）
    },
    {
      label:          '热流道温控箱报警',
      sheet:          ss.getSheetByName("IOT温控箱邮件反馈格式"),
      totalCols:      21,
      dateCol:        4,
      timeCol:        5,   // 报警时间
      dateIsDatetime: false,
      statusCol:      16,
      workshopCol:    2,
      machineCol:     7,   // 机台号
      typeCol:        8,   // 点位
    },
    {
      label:          'Full Opera机台炮筒温度报警',
      sheet:          ss.getSheetByName("Opera炮筒温度邮件反馈格式"),
      totalCols:      20,
      dateCol:        4,
      timeCol:        5,   // 报警时间
      dateIsDatetime: false,
      statusCol:      16,
      workshopCol:    2,
      machineCol:     8,
      typeCol:        9,
    },
    {
      label:          '漏料检查报警',
      sheet:          ss.getSheetByName("漏料邮件反馈格式"),
      totalCols:      12,
      dateCol:        4,   // 时间点（"yyyy-MM-dd HH:mm:ss"，已含时间）
      timeCol:        -1,  // datetime 合并在 dateCol 中
      dateIsDatetime: true,
      statusCol:      8,   // 状态
      workshopCol:    2,
      machineCol:     3,   // 机台
      typeCol:        5,   // 参数类型
    },
    {
      label:            '冷却水流量温度报警',
      sheet:            ss.getSheetByName("冷却水邮件反馈格式"),
      totalCols:        13,
      dateCol:          0,   // Date
      timeCol:          1,   // Hour（整数，如 "14"，无分钟）
      dateIsDatetime:   false,
      statusCol:        10,  // 状态
      workshopCol:      -1,
      defaultWorkshop:  'TB2',  // 冷却水仅覆盖TB2
      machineCol:       -1,
      typeCol:          -1,
    },
  ];
}

// 主函数：生成并发送经理报警监控邮件
function sendManagerAlertReport(e) {
  let configs = getAlertSheetsConfig();
  let now = new Date();
  let todayStr  = Utilities.formatDate(now, currentTimeZone, 'yyyy-MM-dd');
  let cutoffStr = Utilities.formatDate(new Date(now.getTime() - 30 * 24 * 3600 * 1000), currentTimeZone, 'yyyy-MM-dd');

  let summary = [];
  let details = [];
  let topMachines = {};

  configs.forEach(cfg => {
    if (!cfg.sheet) return;
    let lr = cfg.sheet.getLastRow();
    if (lr < 2) return;

    let data = cfg.sheet.getRange(2, 1, lr - 1, cfg.totalCols).getDisplayValues();

    let recent = data.filter(row => {
      let d = getDateStr(row[cfg.dateCol], cfg.dateIsDatetime);
      return d >= cutoffStr && d <= todayStr;
    });

    let getWs = row => cfg.workshopCol >= 0
      ? ((row[cfg.workshopCol] || '').toString().trim())
      : (cfg.defaultWorkshop || 'TB2');

    let tb1Rows = recent.filter(row => getWs(row) === 'TB1');
    let tb2Rows = recent.filter(row => getWs(row) === 'TB2');
    let untrackedAll = recent.filter(row => (row[cfg.statusCol] || '') === '');

    summary.push({
      label:        cfg.label,
      total:        recent.length,
      untracked:    untrackedAll.length,
      hasWs:        cfg.workshopCol >= 0,          // 是否有真实车间列（冷却水无）
      tb1Total:     tb1Rows.length,
      tb1Untracked: tb1Rows.filter(row => (row[cfg.statusCol] || '') === '').length,
      tb2Total:     tb2Rows.length,
      tb2Untracked: tb2Rows.filter(row => (row[cfg.statusCol] || '') === '').length,
    });

    untrackedAll.slice(0, 50).forEach(row => {
      details.push([
        cfg.label,
        getDateStr(row[cfg.dateCol], cfg.dateIsDatetime),
        cfg.workshopCol >= 0 ? (row[cfg.workshopCol] || '-') : (cfg.defaultWorkshop || '-'),
        cfg.machineCol  >= 0 ? (row[cfg.machineCol]  || '-') : '-',
        cfg.typeCol     >= 0 ? (row[cfg.typeCol]     || '-') : '-',
      ]);
    });

    // Top 3 机台（仅IoT/Opera和温控箱）
    if (cfg.machineCol >= 0 &&
        (cfg.label === 'IoT/Opera周期报警' || cfg.label === '热流道温控箱报警')) {
      let mc = {}, mu = {};
      recent.forEach(row => {
        let m = (row[cfg.machineCol] || '-').toString().trim() || '-';
        mc[m] = (mc[m] || 0) + 1;
        if ((row[cfg.statusCol] || '') === '') mu[m] = (mu[m] || 0) + 1;
      });
      topMachines[cfg.label] = Object.entries(mc)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([machine, total]) => ({ machine, total, untracked: mu[machine] || 0 }));
    }
  });

  let trend    = getOosTrend(configs, 14);
  let iotTrend = getIotOperaDailyTrend(14);
  let html     = buildReportHtml(summary, details, trend, iotTrend, topMachines, todayStr, cutoffStr);

  let tomail = getReportRecipients();
  GmailApp.sendEmail(tomail, '【注塑Opera/ IoT工艺报警监控报告】' + todayStr, '', {
    htmlBody: html,
    name:     'GAS报警监控系统',
  });

  const trigger = e && e.triggerType === 'scheduled' ? '定时' : '手动';
  writeLog('sendManagerAlertReport', '成功', '已发送至 ' + tomail, trigger, '');
}

// 获取近N天各日OOS总数
function getOosTrend(configs, days) {
  let now = new Date();
  let dateKeys = [];
  let trendMap = {};

  for (let i = days - 1; i >= 0; i--) {
    let d = Utilities.formatDate(new Date(now.getTime() - i * 24 * 3600 * 1000), currentTimeZone, 'yyyy-MM-dd');
    dateKeys.push(d);
    trendMap[d] = 0;
  }
  let cutoff = dateKeys[0];
  let today  = dateKeys[dateKeys.length - 1];

  configs.forEach(cfg => {
    if (!cfg.sheet) return;
    let lr = cfg.sheet.getLastRow();
    if (lr < 2) return;
    let data = cfg.sheet.getRange(2, cfg.dateCol + 1, lr - 1, 1).getDisplayValues();
    data.forEach(row => {
      let d = getDateStr(row[0], cfg.dateIsDatetime);
      if (d >= cutoff && d <= today && trendMap.hasOwnProperty(d)) trendMap[d]++;
    });
  });

  return dateKeys.map(d => [d, trendMap[d]]);
}

// 从 Mail sheet 读取收件人，排除 Tommy、Grant 和 Ken
function getReportRecipients() {
  const EXCLUDE = ['tommy_xu@colpal.com', 'grant_zhang@colpal.com', 'ken_yang@colpal.com'];
  let ss = SpreadsheetApp.openById(ALERT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Mail');
  let lr = sheet.getLastRow();
  if (lr < 2) return ALERT_MANAGER_EMAIL;
  let all = sheet.getRange(2, 1, lr - 1, 1).getDisplayValues()
    .map(r => r[0].trim())
    .filter(email => email !== '' && EXCLUDE.indexOf(email) === -1);
  return all.join(',');
}

// 提取日期字符串（兼容纯日期和datetime字符串）
function getDateStr(raw, isDatetime) {
  let s = (raw || '').toString().trim();
  return isDatetime ? s.substring(0, 10) : s;
}

// 获取近N天 IoT/Opera 周期报警日趋势（每日总数 + 超下限数）
function getIotOperaDailyTrend(days) {
  let sheet = SpreadsheetApp.openById(ALERT_SPREADSHEET_ID).getSheetByName("IOT/Opera邮件反馈格式");
  if (!sheet) return [];

  let now = new Date();
  let dateKeys = [], totalMap = {}, subLimitMap = {};
  for (let i = days - 1; i >= 0; i--) {
    let d = Utilities.formatDate(new Date(now.getTime() - i * 24 * 3600 * 1000), currentTimeZone, 'yyyy-MM-dd');
    dateKeys.push(d);
    totalMap[d]    = 0;
    subLimitMap[d] = 0;
  }
  let cutoff = dateKeys[0], today = dateKeys[dateKeys.length - 1];

  let lr = sheet.getLastRow();
  if (lr < 2) return dateKeys.map(function(d) { return [d, 0, 0]; });

  // 读 dateCol(E=idx4) 和 diffCol(M=idx12)，取前13列即可
  let data = sheet.getRange(2, 1, lr - 1, 13).getDisplayValues();
  data.forEach(function(row) {
    let d = (row[4] || '').toString().trim();
    if (d >= cutoff && d <= today && totalMap.hasOwnProperty(d)) {
      totalMap[d]++;
      let diff = parseFloat((row[12] || '').toString().replace(/,/g, '').trim());
      if (!isNaN(diff) && diff < 0) subLimitMap[d]++;
    }
  });

  return dateKeys.map(function(d) { return [d, totalMap[d], subLimitMap[d]]; });
}

// 构建 IoT/Opera 双色并排柱状图（总报警数 + 超下限报警数）
function buildIotOperaTrendHtml(trendData) {
  let CHART_H = 80;
  let BAR_W   = 14;
  let maxVal  = Math.max.apply(null, trendData.map(function(r) { return r[1]; }).concat([1]));

  let dayCols = '';
  trendData.forEach(function(r) {
    let totalH = r[1] > 0 ? Math.max(Math.round(r[1] / maxVal * CHART_H), 3) : 0;
    let subH   = r[2] > 0 ? Math.max(Math.round(r[2] / maxVal * CHART_H), 3) : 0;
    let colW   = BAR_W * 2 + 2;

    dayCols +=
      '<td style="padding:0 3px;text-align:center">' +
        // 数值标签行
        '<table style="border-collapse:collapse;width:' + colW + 'px;margin:0 auto"><tr>' +
          '<td style="width:' + BAR_W + 'px;font-size:8px;text-align:center;padding:0;color:#555">' + (r[1] > 0 ? r[1] : '') + '</td>' +
          '<td style="width:2px;padding:0"></td>' +
          '<td style="width:' + BAR_W + 'px;font-size:8px;text-align:center;padding:0;color:#777">' + (r[2] > 0 ? r[2] : '') + '</td>' +
        '</tr></table>' +
        // 柱体行（固定高度容器，vertical-align:bottom 使柱体贴底对齐）
        '<table style="border-collapse:collapse;width:' + colW + 'px;margin:0 auto"><tr>' +
          '<td style="width:' + BAR_W + 'px;height:' + CHART_H + 'px;padding:0;vertical-align:bottom">' +
            (totalH > 0 ? '<div style="background:#E06060;width:' + BAR_W + 'px;height:' + totalH + 'px"></div>' : '') +
          '</td>' +
          '<td style="width:2px;height:' + CHART_H + 'px;padding:0;background:#eee"></td>' +
          '<td style="width:' + BAR_W + 'px;height:' + CHART_H + 'px;padding:0;vertical-align:bottom">' +
            (subH > 0 ? '<div style="background:#F0A060;width:' + BAR_W + 'px;height:' + subH + 'px"></div>' : '') +
          '</td>' +
        '</tr></table>' +
        // 日期标签
        '<div style="font-size:9px;color:#666;padding-top:2px">' + r[0].substring(5) + '</div>' +
      '</td>';
  });

  let legend =
    '<div style="margin-left:10px;margin-bottom:5px;font-size:10px;color:#444">' +
    '<span style="display:inline-block;width:10px;height:10px;background:#E06060;margin-right:3px;vertical-align:middle"></span>总报警数&nbsp;&nbsp;&nbsp;' +
    '<span style="display:inline-block;width:10px;height:10px;background:#F0A060;margin-right:3px;vertical-align:middle"></span>超下限报警数' +
    '</div>';

  return legend +
    '<table style="border-collapse:collapse;margin-left:10px"><tr>' + dayCols + '</tr></table>';
}

// 构建近N天 OOS 趋势 HTML 柱状图（邮件客户端兼容，不用SVG）
function buildTrendHtml(trend) {
  let CHART_H = 80;  // 柱体区域高度 px
  let BAR_W   = 30;  // 每柱宽度 px
  let maxVal  = Math.max.apply(null, trend.map(function(r) { return r[1]; }).concat([1]));

  let barCells   = '';
  let labelCells = '';

  trend.forEach(function(r) {
    let barH   = r[1] > 0 ? Math.max(Math.round(r[1] / maxVal * CHART_H), 4) : 0;
    let emptyH = CHART_H - barH;
    let ratio  = r[1] / maxVal;
    let g      = Math.round(200 - ratio * 160);
    if (g < 0) g = 0;
    let color  = r[1] > 0 ? 'rgb(230,' + g + ',' + g + ')' : '#ddd';

    // 嵌套 table：上半空白行（含数值标签）+ 下半彩色柱
    barCells +=
      '<td style="padding:0 2px;vertical-align:bottom">' +
        '<table style="border-collapse:collapse;width:' + BAR_W + 'px">' +
          '<tr><td style="height:' + emptyH + 'px;padding:0 0 1px;font-size:9px;' +
            'text-align:center;vertical-align:bottom;color:#444">' +
            (r[1] > 0 ? r[1] : '') + '</td></tr>' +
          '<tr><td style="height:' + Math.max(barH, 1) + 'px;padding:0;background:' + color + '"></td></tr>' +
        '</table>' +
      '</td>';

    labelCells +=
      '<td style="padding:2px 2px 0;text-align:center;font-size:9px;' +
        'color:#666;width:' + BAR_W + 'px">' + r[0].substring(5) + '</td>';
  });

  return '<table style="border-collapse:collapse;margin-left:10px">' +
    '<tr>' + barCells   + '</tr>' +
    '<tr>' + labelCells + '</tr>' +
    '</table>';
}

// 构建HTML邮件正文
function buildReportHtml(summary, details, trend, iotTrend, topMachines, todayStr, cutoffStr) {
  let totalAlarms    = summary.reduce(function(s, r) { return s + r.total; }, 0);
  let totalUntracked = summary.reduce(function(s, r) { return s + r.untracked; }, 0);
  let totalRate      = totalAlarms > 0 ? Math.round((1 - totalUntracked / totalAlarms) * 100) : 100;

  // —— 1. 报警汇总表（带车间分列）——
  let summaryHtml = `
    <h3 style="color:#E60012;margin:16px 0 4px">一、报警汇总</h3>
    <p style="margin:0 0 6px 10px;font-size:12px;color:#555">数据范围：<b>${cutoffStr}</b> 至 <b>${todayStr}</b>（近30天）</p>
    <table border="1" style="border-collapse:collapse;font-size:11px;margin-left:10px">
      <tr style="background:#E60012;color:white;text-align:center">
        <th rowspan="2" style="padding:4px 10px;vertical-align:middle">报警类型</th>
        <th colspan="3" style="padding:3px 8px;border-bottom:1px solid rgba(255,255,255,0.4)">TB1</th>
        <th colspan="3" style="padding:3px 8px;border-bottom:1px solid rgba(255,255,255,0.4)">TB2</th>
        <th colspan="3" style="padding:3px 8px;border-bottom:1px solid rgba(255,255,255,0.4)">合计</th>
      </tr>
      <tr style="background:#C00010;color:white;text-align:center">
        <th style="padding:2px 8px">总数</th><th style="padding:2px 8px">未跟进</th><th style="padding:2px 8px">跟进率</th>
        <th style="padding:2px 8px">总数</th><th style="padding:2px 8px">未跟进</th><th style="padding:2px 8px">跟进率</th>
        <th style="padding:2px 8px">总数</th><th style="padding:2px 8px">未跟进</th><th style="padding:2px 8px">跟进率</th>
      </tr>`;

  summary.forEach(function(r) {
    let totalRate = r.total > 0 ? Math.round((1 - r.untracked / r.total) * 100) : 100;
    let tb1Rate   = r.tb1Total > 0 ? Math.round((1 - r.tb1Untracked / r.tb1Total) * 100) : 100;
    let tb2Rate   = r.tb2Total > 0 ? Math.round((1 - r.tb2Untracked / r.tb2Total) * 100) : 100;
    let bg = r.untracked > 0 ? '#FFF0F0' : '#F0FFF0';
    // 冷却水TB1列显示"—"
    let tb1TotalCell     = r.hasWs ? r.tb1Total     : '—';
    let tb1UntrackedCell = r.hasWs ? r.tb1Untracked : '—';
    let tb1RateCell      = r.hasWs ? (tb1Rate + '%') : '—';
    summaryHtml += `
      <tr style="background:${bg};text-align:center">
        <td style="padding:3px 10px;text-align:left">${r.label}</td>
        <td style="padding:3px 8px">${tb1TotalCell}</td>
        <td style="padding:3px 8px">${r.hasWs ? '<b>' + r.tb1Untracked + '</b>' : '—'}</td>
        <td style="padding:3px 8px">${tb1RateCell}</td>
        <td style="padding:3px 8px">${r.tb2Total}</td>
        <td style="padding:3px 8px"><b>${r.tb2Untracked}</b></td>
        <td style="padding:3px 8px">${tb2Rate}%</td>
        <td style="padding:3px 8px">${r.total}</td>
        <td style="padding:3px 8px"><b>${r.untracked}</b></td>
        <td style="padding:3px 8px">${totalRate}%</td>
      </tr>`;
  });

  summaryHtml += `
      <tr style="background:#f0f0f0;font-weight:bold;text-align:center">
        <td style="padding:3px 10px;text-align:left">合计</td>
        <td colspan="3" style="padding:3px 8px">—</td>
        <td colspan="3" style="padding:3px 8px">—</td>
        <td style="padding:3px 8px">${totalAlarms}</td>
        <td style="padding:3px 8px">${totalUntracked}</td>
        <td style="padding:3px 8px">${totalRate}%</td>
      </tr>
    </table>`;

  // —— 2. OOS趋势柱状图（两张）——
  let trendHtml = `
    <h3 style="color:#E60012;margin:20px 0 6px">二、近14天报警趋势</h3>
    <p style="margin:0 0 4px 10px;font-size:11px;color:#555">（一）全部报警总数</p>
    ${buildTrendHtml(trend)}
    <p style="margin:12px 0 4px 10px;font-size:11px;color:#555">（二）IoT/Opera 周期报警：总数 vs 超下限</p>
    ${buildIotOperaTrendHtml(iotTrend)}`;

  // —— 3. 未跟进明细 ——
  let detailHtml = `<h3 style="color:#E60012;margin:20px 0 6px">三、未跟进报警明细（每类最多50条）</h3>`;
  if (details.length > 0) {
    detailHtml += `
      <table border="1" style="border-collapse:collapse;font-size:11px;margin-left:10px">
        <tr style="background:#E60012;color:white;text-align:center">
          <th style="padding:3px 10px">报警类型</th>
          <th style="padding:3px 10px">日期</th>
          <th style="padding:3px 10px">车间</th>
          <th style="padding:3px 10px">机台</th>
          <th style="padding:3px 10px">参数/类型</th>
        </tr>`;
    details.forEach(function(row) {
      detailHtml += '<tr style="text-align:center">' +
        row.map(function(cell) { return '<td style="padding:2px 10px">' + cell + '</td>'; }).join('') +
        '</tr>';
    });
    detailHtml += '</table>';
  } else {
    detailHtml += '<p style="color:green;margin-left:10px">✅ 近30天所有报警均已跟进，状态良好！</p>';
  }

  // —— 4. 重点机台 TOP 3 ——
  let TOP_LABELS = ['IoT/Opera周期报警', '热流道温控箱报警'];
  let topHtml = `<h3 style="color:#E60012;margin:20px 0 6px">四、重点机台报警 TOP 3（近30天）</h3>`;
  TOP_LABELS.forEach(function(label) {
    let machines = topMachines[label] || [];
    topHtml += `<p style="margin:8px 0 4px 10px;font-weight:bold;font-size:12px">${label}</p>`;
    if (machines.length === 0) {
      topHtml += '<p style="margin-left:10px;font-size:11px;color:#999">无数据</p>';
      return;
    }
    topHtml += `<table border="1" style="border-collapse:collapse;font-size:11px;margin-left:10px;margin-bottom:12px">
      <tr style="background:#E60012;color:white;text-align:center">
        <th style="padding:3px 10px">排名</th>
        <th style="padding:3px 16px">机台</th>
        <th style="padding:3px 12px">报警次数</th>
        <th style="padding:3px 12px">未跟进</th>
        <th style="padding:3px 12px">跟进率</th>
      </tr>`;
    machines.forEach(function(m, i) {
      let rate = m.total > 0 ? Math.round((1 - m.untracked / m.total) * 100) : 100;
      let bg   = m.untracked > 0 ? '#FFF0F0' : '#F0FFF0';
      topHtml += `<tr style="background:${bg};text-align:center">
        <td style="padding:3px 10px">${i + 1}</td>
        <td style="padding:3px 16px">${m.machine}</td>
        <td style="padding:3px 12px">${m.total}</td>
        <td style="padding:3px 12px"><b>${m.untracked}</b></td>
        <td style="padding:3px 12px">${rate}%</td>
      </tr>`;
    });
    topHtml += '</table>';
  });

  return `<div style="font-family:Arial,sans-serif;font-size:13px;color:#333;padding:10px">
    <p>Dear Kelland,</p>
    <p>以下是截至 <b>${todayStr}</b> 的注塑Opera/ IoT工艺报警监控报告：</p>
    ${summaryHtml}
    ${trendHtml}
    ${detailHtml}
    ${topHtml}
    <p style="color:#aaa;font-size:11px;margin-top:24px">— GAS 报警监控系统自动发送</p>
  </div>`;
}
