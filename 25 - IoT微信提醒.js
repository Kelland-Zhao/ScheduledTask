// ============================================================
// 29 - 微信提醒.js — 企业微信群机器人班次提醒（IoT报警未跟进）
// 入口：sendWechatShiftReminder
// 依赖：30 - 报警监控.js（getAlertSheetsConfig）
// 菜单：发送IoT企业微信班次提醒
// ============================================================

// 主函数：统计当班时间段内未跟进报警，推送到企业微信群（按车间分组）
function sendWechatShiftReminder(e) {
  let configs = getAlertSheetsConfig();
  let now     = new Date();
  let window  = getShiftWindow(now);

  let totalUntracked = 0;
  let workshopMap = { TB1: {}, TB2: {} };

  configs.forEach(cfg => {
    if (!cfg.sheet || cfg.statusCol < 0) return;
    let lr = cfg.sheet.getLastRow();
    if (lr < 2) return;

    let data = cfg.sheet.getRange(2, 1, lr - 1, cfg.totalCols).getDisplayValues();

    let untracked = data.filter(row => {
      let dt = getRowDatetimeStr(row, cfg);
      return dt >= window.startStr && dt < window.endStr && (row[cfg.statusCol] || '') === '';
    });

    untracked.forEach(row => {
      let ws = cfg.workshopCol >= 0
        ? ((row[cfg.workshopCol] || '').toString().trim() || 'TB1')
        : (cfg.defaultWorkshop || 'TB2');
      if (!workshopMap[ws]) workshopMap[ws] = {};
      workshopMap[ws][cfg.label] = (workshopMap[ws][cfg.label] || 0) + 1;
      totalUntracked++;
    });
  });

  let nowStr     = Utilities.formatDate(now, currentTimeZone, 'MM-dd HH:mm');
  let statusLine = totalUntracked === 0
    ? '✅ **本班次所有报警均已跟进，状态良好！**'
    : `⚠️ **共 ${totalUntracked} 条报警未跟进，请及时处理！**`;

  let contentParts = [
    '## 🔔 IoT报警跟进提醒',
    `**时间**：${nowStr}　　**班次**：${window.shiftLabel}`,
    `**数据范围**：${window.rangeStr}`,
    '',
    statusLine,
    '',
  ];

  ['TB1', 'TB2'].forEach(ws => {
    let entries = Object.entries(workshopMap[ws]);
    let wsTotal = entries.reduce((s, [, v]) => s + v, 0);
    if (wsTotal === 0) {
      contentParts.push(`**🏭 ${ws}**：✅ 无未跟进`);
    } else {
      contentParts.push(`**🏭 ${ws}（${wsTotal}条）**`);
      entries.forEach(([label, count]) => {
        contentParts.push(`> ${label}：${count} 条`);
      });
    }
    contentParts.push('');
  });

  contentParts.push('> 请当班组长在后端 Sheet 中填写**原因分析**和**解决措施**');
  if (totalUntracked > 0) contentParts.push('<@所有人>');

  sendWechatMessage(contentParts.join('\n'));
  const trigger = e && e.triggerType === 'scheduled' ? '定时' : '手动';
  writeLog('sendWechatShiftReminder', '成功',
    `${window.shiftLabel} 未跟进: ${totalUntracked}`, trigger, '');
}

// 根据当前时间计算班次时间窗口
// 08:00 → 夜班（前一天23:00 ~ 今日07:00）
// 16:00 → 白班（今日07:00 ~ 今日15:00）
// 00:00 → 中班（前一天15:00 ~ 前一天23:00）
function getShiftWindow(now) {
  let h   = now.getHours();
  let mid = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0); // 今日0点

  let startDt, endDt, shiftLabel;

  if (h >= 6 && h <= 10) {         // 08:00 触发
    startDt    = new Date(mid.getTime() - 1 * 3600 * 1000);   // 昨日 23:00
    endDt      = new Date(mid.getTime() + 7 * 3600 * 1000);   // 今日 07:00
    shiftLabel = '夜班';
  } else if (h >= 14 && h <= 18) { // 16:00 触发
    startDt    = new Date(mid.getTime() + 7 * 3600 * 1000);   // 今日 07:00
    endDt      = new Date(mid.getTime() + 15 * 3600 * 1000);  // 今日 15:00
    shiftLabel = '白班';
  } else {                          // 00:00 触发（h=0 或 h>=22）
    let refMid = (h <= 2)
      ? new Date(mid.getTime() - 24 * 3600 * 1000)  // 昨日0点
      : mid;
    startDt    = new Date(refMid.getTime() + 15 * 3600 * 1000); // 昨日/今日 15:00
    endDt      = new Date(refMid.getTime() + 23 * 3600 * 1000); // 昨日/今日 23:00
    shiftLabel = '中班';
  }

  let fmt = t => Utilities.formatDate(t, currentTimeZone, 'yyyy-MM-dd HH:mm:ss');
  let fmtShort = t => Utilities.formatDate(t, currentTimeZone, 'MM-dd HH:mm');

  return {
    startDt,
    endDt,
    startStr:   fmt(startDt),
    endStr:     fmt(endDt),
    shiftLabel,
    rangeStr:   `${fmtShort(startDt)} ~ ${fmtShort(endDt)}`,
  };
}

// 将各Sheet行数据转换为可比较的 datetime 字符串（yyyy-MM-dd HH:mm:ss）
function getRowDatetimeStr(row, cfg) {
  let dateRaw = (row[cfg.dateCol] || '').toString().trim();

  if (cfg.dateIsDatetime) {
    // 漏料：时间点已是 "yyyy-MM-dd HH:mm:ss"
    return dateRaw.length >= 19 ? dateRaw.substring(0, 19) : dateRaw + ' 00:00:00';
  }

  if (cfg.timeCol === 1 && cfg.dateCol === 0) {
    // 冷却水：Date + Hour（整数列，无分钟）
    let hourRaw = (row[cfg.timeCol] || '0').toString().trim();
    let hh = hourRaw.padStart(2, '0');
    return `${dateRaw} ${hh}:00:00`;
  }

  if (cfg.timeCol >= 0) {
    // IOT/Opera、温控箱、炮筒温度：日期 + 报警时间
    let timeRaw = (row[cfg.timeCol] || '00:00:00').toString().trim();
    return `${dateRaw} ${timeRaw}`;
  }

  return `${dateRaw} 00:00:00`;
}

// 发送 Markdown 消息到企业微信群
function sendWechatMessage(content) {
  let payload = JSON.stringify({
    msgtype:  'markdown',
    markdown: { content: content },
  });
  UrlFetchApp.fetch(ALERT_WECHAT_WEBHOOK_URL, {
    method:      'POST',
    contentType: 'application/json',
    payload:     payload,
  });
}
