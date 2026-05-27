// V20260527.1 — 保养反馈同步模块
// 功能：从 MasterData 表读取「后续措施 - 保养（反馈）」数据，同步写入各工序车间表

// ========== 保养反馈同步配置 ==========
const FEEDBACK_CONFIG = {
  MASTER_SHEET_ID: '1HZHz5wN8sXeP5S7ub041bqklk0Rm2Jsmh3Ovd7ZKeJE',
  MASTER_SHEET_NAME: 'MasterData',
  TARGET_SHEET_ID: '10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w',

  // 工序 → 目标车间表映射
  PROCESS_SHEET_MAPPING: {
    'INJ': ['Shift_INJ_TB1', 'Shift_INJ_TB2'],
    'TF': ['Shift_TF_TB1', 'Shift_TF_TB2'],
    'PK': ['Shift_PK_TB1', 'Shift_PK_TB2']
  }
};

/** 同步保养反馈数据：从 MasterData 读取反馈值，匹配编号写入各工序车间表 */
function syncMaintenanceFeedbackData() {
  const fnName = 'syncMaintenanceFeedbackData';
  let updateCount = 0;
  let notFoundCount = 0;
  let skipCount = 0;

  try {
    console.log('开始同步保养反馈数据...');

    // 1. 打开源表 MasterData
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

    // 2. 打开目标表
    const targetSS = SpreadsheetApp.openById(FEEDBACK_CONFIG.TARGET_SHEET_ID);

    // 3. 遍历 MasterData 每一行
    for (let i = 1; i < masterData.length; i++) {
      const row = masterData[i];
      const processType = row[processCol];
      const itemId = row[idCol];
      const feedbackValue = row[feedbackCol];

      if (!processType || !itemId) {
        console.log('跳过第 ' + (i + 1) + ' 行：工序或编号为空');
        skipCount++;
        continue;
      }

      const targetSheets = FEEDBACK_CONFIG.PROCESS_SHEET_MAPPING[processType];
      if (!targetSheets) {
        console.log('跳过第 ' + (i + 1) + ' 行：不支持的工序类型 ' + processType);
        skipCount++;
        continue;
      }

      let updated = false;

      for (let s = 0; s < targetSheets.length; s++) {
        const sheetName = targetSheets[s];
        const targetSheet = targetSS.getSheetByName(sheetName);
        if (!targetSheet) {
          console.log('警告：找不到目标表 ' + sheetName);
          continue;
        }

        const targetData = targetSheet.getDataRange().getValues();
        const targetHeaders = targetData[0];

        const targetIdCol = targetHeaders.indexOf('编号');
        const targetFeedbackCol = targetHeaders.indexOf('后续措施 - 保养（反馈）');

        if (targetIdCol === -1 || targetFeedbackCol === -1) {
          console.log('警告：表 ' + sheetName + ' 缺少必要的列');
          continue;
        }

        for (let j = 1; j < targetData.length; j++) {
          if (targetData[j][targetIdCol] === itemId) {
            targetSheet.getRange(j + 1, targetFeedbackCol + 1).setValue(feedbackValue);
            console.log('成功更新 ' + sheetName + ' 中编号 ' + itemId);
            updateCount++;
            updated = true;
            break;
          }
        }

        if (updated) break;
      }

      if (!updated) {
        console.log('未找到编号 ' + itemId + ' (工序: ' + processType + ')');
        notFoundCount++;
      }
    }

    console.log('同步完成：成功更新 ' + updateCount + ' 条，未找到 ' + notFoundCount + ' 条');
    writeLog(fnName, '成功', '更新: ' + updateCount + ' | 未找到: ' + notFoundCount + ' | 跳过: ' + skipCount + ' | 总处理: ' + (masterData.length - 1),
      '定时', '');

    return { success: true, updateCount: updateCount, notFoundCount: notFoundCount, skipCount: skipCount };

  } catch (error) {
    console.error('同步保养反馈数据时出错:', error);
    writeLog(fnName, '失败', error.message, '定时', error.stack || '');

    try {
      GmailApp.sendEmail(FAULT_CONFIG.ADMIN_EMAIL,
        '[系统错误] 保养反馈同步',
        '保养反馈同步出错:\n\n错误: ' + error.message + '\n时间: ' + Utilities.formatDate(new Date(), currentTimeZone, 'yyyy-MM-dd HH:mm:ss'),
        { from: getFaultEmailSender() }
      );
    } catch (mailError) {
      console.error('发送错误通知时出错:', mailError);
    }

    throw error;
  }
}
