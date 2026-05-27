// V20260528.1 — 保养反馈同步模块（性能优化版）
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
  },

  ADMIN_EMAIL: 'kelland_zhao@colpal.com'
};

/** 同步保养反馈数据：从 MasterData 读取反馈值，匹配编号写入各工序车间表 */
function syncMaintenanceFeedbackData() {
  const fnName = 'syncMaintenanceFeedbackData';
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

    // 2. 预读所有目标车间表，建立 编号→行号 索引（只读一次）
    const targetSS = SpreadsheetApp.openById(FEEDBACK_CONFIG.TARGET_SHEET_ID);
    const sheetCache = {};

    for (var processKey in FEEDBACK_CONFIG.PROCESS_SHEET_MAPPING) {
      var sheetNames = FEEDBACK_CONFIG.PROCESS_SHEET_MAPPING[processKey];
      for (var s = 0; s < sheetNames.length; s++) {
        var sheetName = sheetNames[s];
        var sheet = targetSS.getSheetByName(sheetName);
        if (!sheet) continue;

        var data = sheet.getDataRange().getValues();
        var headers = data[0];
        var targetIdCol = headers.indexOf('编号');
        var targetFeedbackCol = headers.indexOf('后续措施 - 保养（反馈）');

        if (targetIdCol === -1 || targetFeedbackCol === -1) continue;

        // 建立 编号→行号 映射（0-based row index）
        var lookup = {};
        for (var j = 1; j < data.length; j++) {
          var id = String(data[j][targetIdCol] || '').trim();
          if (id) lookup[id] = j;
        }

        sheetCache[sheetName] = {
          sheet: sheet,
          lookup: lookup,
          feedbackCol: targetFeedbackCol
        };
      }
    }

    console.log('已缓存 ' + Object.keys(sheetCache).length + ' 个车间表的索引');

    // 3. 遍历 MasterData，通过索引快速匹配
    var numRows = masterData.length;
    for (var i = 1; i < numRows; i++) {
      var row = masterData[i];
      var processType = String(row[processCol] || '').trim();
      var itemId = String(row[idCol] || '').trim();
      var feedbackValue = row[feedbackCol];

      if (!processType || !itemId) {
        skipCount++;
        continue;
      }

      var targetSheets = FEEDBACK_CONFIG.PROCESS_SHEET_MAPPING[processType];
      if (!targetSheets) {
        skipCount++;
        continue;
      }

      var found = false;
      for (var t = 0; t < targetSheets.length; t++) {
        var tgtSheetName = targetSheets[t];
        var cache = sheetCache[tgtSheetName];
        if (!cache) continue;

        var rowIdx = cache.lookup[itemId];
        if (rowIdx !== undefined) {
          // 写入（使用索引定位，直接写单元格）
          cache.sheet.getRange(rowIdx + 1, cache.feedbackCol + 1).setValue(feedbackValue);
          updateCount++;
          found = true;
          break;
        }
      }

      if (!found) notFoundCount++;
    }

    console.log('同步完成：成功更新 ' + updateCount + ' 条，未找到 ' + notFoundCount + ' 条');
    writeLog(fnName, '成功', '更新: ' + updateCount + ' | 未找到: ' + notFoundCount + ' | 跳过: ' + skipCount + ' | 总处理: ' + (numRows - 1),
      '定时', '');

    return { success: true, updateCount: updateCount, notFoundCount: notFoundCount, skipCount: skipCount };

  } catch (error) {
    console.error('同步保养反馈数据时出错:', error);
    writeLog(fnName, '失败', error.message, '定时', error.stack || '');

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
