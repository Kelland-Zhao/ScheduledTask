// V20260606.2 — 项目跟踪进度同步至 Kaizen & CI（项目总表 E列 → Master Data N列）
// 格式: 序号. 任务描述  责任人  计划完成时间  状态

var SYNC_PT_SS_ID = "1aoQDjeWU9Xa9clloyTwiXL6WS62tYVbB0-VOavpAgAM";
var SYNC_PT_SHEET = "项目总表";
var SYNC_KC_SS_ID = "1FXeW4EoofbPkeLOB796ALZBoNsujAwGln-Zaam5OOKA";
var SYNC_KC_SHEET = "Master Data";

/** 从 JSON name 提取中文显示名（取 / 前部分） */
function _sp_extractName(fullName) {
  if (!fullName) return "";
  var idx = fullName.indexOf("/");
  if (idx > -1) return fullName.substring(0, idx).trim();
  return fullName.trim();
}

/**
 * 项目跟踪进度同步
 * 读取「项目总表」E列 Milestones_JSON → 格式化为「序号. 任务描述  责任人  计划完成时间  状态」
 * 按 G列项目编号匹配写入「Master Data」N列
 */
function syncProjectProgress(e) {
  var trigger = e ? "定时" : "手动";
  try {
    // 1. 读取项目总表
    var ptSS = SpreadsheetApp.openById(SYNC_PT_SS_ID);
    var ptSheet = ptSS.getSheetByName(SYNC_PT_SHEET);
    var ptData = ptSheet.getDataRange().getValues();

    if (ptData.length <= 1) {
      writeLog("syncProjectProgress", "跳过", "项目总表无数据", trigger, "");
      return;
    }

    // 2. 解析 E列 JSON，构建 项目编号 → 格式化进度 映射
    var progressMap = {};  // { "P001": "1|设计|2026-06-15|张三\n2|采购|2026-07-01|李四" }
    var parseErrors = 0;

    for (var i = 1; i < ptData.length; i++) {
      var row = ptData[i];
      var projectNumber = String(row[6] || "").trim();  // G列：项目编号
      var msJsonRaw = row[4];                            // E列：Milestones_JSON

      if (!projectNumber) continue;

      var progressLines = [];

      if (msJsonRaw) {
        var milestones;
        try {
          milestones = JSON.parse(String(msJsonRaw));
        } catch (err) {
          writeLog("syncProjectProgress", "失败", "JSON解析失败: " + projectNumber, trigger, err.message);
          parseErrors++;
          continue;
        }

        var seq = 0;
        milestones.forEach(function (ms) {
          var planned = String(ms.planned || "").trim();
          if (!planned || planned === "NA") return;

          var displayName = _sp_extractName(String(ms.name || ""));
          var owner = String(ms.owner || "").trim();
          var status = String(ms.status || "").trim();

          seq++;
          progressLines.push(seq + ". " + displayName + "  " + owner + "  " + planned + "  " + status);
        });
      }

      progressMap[projectNumber] = progressLines.join("\n");
    }

    if (Object.keys(progressMap).length === 0) {
      writeLog("syncProjectProgress", "跳过", "无有效项目数据" + (parseErrors > 0 ? "（" + parseErrors + " 个解析失败）" : ""), trigger, "");
      return;
    }

    // 3. 读取 Master Data，按 A列项目编号匹配写入 N列
    var masterSS = SpreadsheetApp.openById(SYNC_KC_SS_ID);
    var masterSheet = masterSS.getSheetByName(SYNC_KC_SHEET);
    var masterData = masterSheet.getDataRange().getValues();

    var updatedCount = 0;
    var unmatchedList = [];

    for (var m = 1; m < masterData.length; m++) {
      var masterProjectNumber = String(masterData[m][0] || "").trim(); // A列
      if (!masterProjectNumber) continue;

      if (progressMap.hasOwnProperty(masterProjectNumber)) {
        var progressStr = progressMap[masterProjectNumber];
        if (progressStr) {
          // N列 = 第14列（index 13），行号 = m + 1（1-indexed）
          masterSheet.getRange(m + 1, 14).setValue(progressStr);
          updatedCount++;
        }
        // 进度为空则保持原值不变
        delete progressMap[masterProjectNumber];
      }
    }

    // 收集未匹配的项目（项目总表有但 Master Data 没有的）
    for (var pn in progressMap) {
      unmatchedList.push(pn);
    }

    var detail = "更新 " + updatedCount + " 行";
    if (unmatchedList.length > 0) {
      detail += "，未匹配 " + unmatchedList.length + " 个项目: " + unmatchedList.join(", ");
    }
    if (parseErrors > 0) {
      detail += "，" + parseErrors + " 个JSON解析失败";
    }

    writeLog("syncProjectProgress", "成功", detail, trigger, "");

  } catch (err) {
    try { writeLog("syncProjectProgress", "失败", err.message, trigger, err.stack || ""); } catch (e2) {}
    console.error(err.stack || err.message);
  }
}
