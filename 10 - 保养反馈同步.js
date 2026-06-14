// V20260608.1 — 保养反馈同步模块（适配合并表 Shift_Records）
// 功能：从 MasterData 表读取「后续措施 - 保养（反馈）」数据，同步写入 Shift_Records 合并表

// ========== 保养反馈同步配置 ==========
const FEEDBACK_CONFIG = {
  MASTER_SHEET_ID: '1HZHz5wN8sXeP5S7ub041bqklk0Rm2Jsmh3Ovd7ZKeJE',
  MASTER_SHEET_NAME: 'MasterData',
  TARGET_SHEET_ID: '10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w',
  TARGET_SHEET_NAME: 'Shift_Records',

  ADMIN_EMAIL: 'kelland_zhao@colpal.com'
};

/** 同步保养反馈数据：从 MasterData 读取反馈值，匹配编号写入 Shift_Records（定时 e 存在 / 手动 e 为 undefined） */
function syncMaintenanceFeedbackData(e) {
  const fnName = 'syncMaintenanceFeedbackData';
  const trigger = e ? '定时' : '手动';
  let updateCount = 0;
  let notFoundCount = 0;
  let skipCount = 0;

  try {
    console.log('开始同步保养反馈数据...');

    // 1. 读取 MasterData
    const masterSS = SpreadsheetApp.openById(FEEDBACK_CONFIG.MASTER_SHEET_ID);
    const masterSheet = masterSS.getSheetByName(FEEDBACK_CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('找不到源数据表: ' + FEEDBACK_CONFIG.MASTER_SHEET_NAME);

    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];

    const processCol = masterHeaders.indexOf('工序');
    const idCol = masterHeaders.indexOf('编号');
    const feedbackCol = masterHeaders.indexOf('后续措施 - 保养（反馈）');

    if (processCol === -1 || idCol === -1 || feedbackCol === -1) {
      throw new Error('MasterData 表缺少必要的列：工序、编号或后续措施 - 保养（反馈）');
    }

    // 2. 读取 Shift_Records 合并表，建立 编号→行号 索引
    const targetSS = SpreadsheetApp.openById(FEEDBACK_CONFIG.TARGET_SHEET_ID);
    const targetSheet = targetSS.getSheetByName(FEEDBACK_CONFIG.TARGET_SHEET_NAME);
    if (!targetSheet) throw new Error('找不到目标表: ' + FEEDBACK_CONFIG.TARGET_SHEET_NAME);

    const targetData = targetSheet.getDataRange().getValues();
    const targetHeaders = targetData[0];
    const targetIdCol = targetHeaders.indexOf('编号');
    const targetFeedbackCol = targetHeaders.indexOf('后续措施 - 保养（反馈）');

    if (targetIdCol === -1 || targetFeedbackCol === -1) {
      throw new Error('Shift_Records 表缺少必要的列：编号或后续措施 - 保养（反馈）');
    }

    // 建立 编号→行号 映射（0-based row index）
    const lookup = {};
    for (let j = 1; j < targetData.length; j++) {
      const id = String(targetData[j][targetIdCol] || '').trim();
      if (id) lookup[id] = j;
    }

    console.log('Shift_Records 已缓存 ' + Object.keys(lookup).length + ' 条编号索引');

    // 3. 遍历 MasterData，按编号匹配写入
    for (let i = 1; i < masterData.length; i++) {
      const row = masterData[i];
      const processType = String(row[processCol] || '').trim();
      const itemId = String(row[idCol] || '').trim();
      const feedbackValue = row[feedbackCol];

      if (!processType || !itemId) {
        skipCount++;
        continue;
      }

      const rowIdx = lookup[itemId];
      if (rowIdx !== undefined) {
        targetSheet.getRange(rowIdx + 1, targetFeedbackCol + 1).setValue(feedbackValue);
        updateCount++;
      } else {
        notFoundCount++;
      }
    }

    console.log('同步完成：成功更新 ' + updateCount + ' 条，未找到 ' + notFoundCount + ' 条');
    writeLog(fnName, '成功', '更新: ' + updateCount + ' | 未找到: ' + notFoundCount + ' | 跳过: ' + skipCount + ' | 总处理: ' + (masterData.length - 1),
      trigger, '');

    return { success: true, updateCount: updateCount, notFoundCount: notFoundCount, skipCount: skipCount };

  } catch (error) {
    console.error('同步保养反馈数据时出错:', error);
    writeLog(fnName, '失败', error.message, trigger, error.stack || '');

    try {
      GmailApp.sendEmail(FEEDBACK_CONFIG.ADMIN_EMAIL,
        '[系统错误] 保养反馈同步',
        '保养反馈同步出错:\n\n错误: ' + error.message + '\n时间: ' + Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd HH:mm:ss')
      );
    } catch (mailError) {
      console.error('发送错误通知时出错:', mailError);
    }

    throw error;
  }
}
