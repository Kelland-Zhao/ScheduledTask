// V20260524.1 — SmartMeeting Alert 全局常量 + 菜单

// ========== 配置表（GAS 项目所属表）==========
const CONFIG_SPREADSHEET_ID = "1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM";
const saas = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
const sbnMenuSet = saas.getSheetByName("菜单设置");
const sbnTimingSet = saas.getSheetByName("定时设置");
const sbnLog = saas.getSheetByName("Log");
const currentTimeZone = saas.getSpreadsheetTimeZone();

// ========== IoT 报警监控（29/30号脚本）==========
const ALERT_SPREADSHEET_ID = "1FtljG58B1uzcgNU_vKEYUSeW39X23ljUYtngEYmWl_A";
const ALERT_WECHAT_WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=fbbae99b-8237-4ba9-beb2-27b6d8845f06";
const ALERT_MANAGER_EMAIL = "kelland_zhao@colpal.com";

// ========== SmartMeeting 业务数据 ==========
const SMARTMEETING_SPREADSHEET_ID = "1Rmf_IJMHNzXv0cfm5AMwZrjfSii955j-vAEGOnA7KP8";
const SMARTMEETING_SHEET_NAME = "会议记录";
const MEETING_TOPIC_FILTER = "INJ SDM";   // 仅处理该会议主题
const DATA_MONTHS = 24;                    // 数据回溯月数
const DONE_STATUS = "Done";                // 完成状态判定值

// ========== 企微推送 ==========
const WECHAT_WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=04a68f66-dff9-47fa-b294-cc709876d6b4";
const MAX_ITEMS_PER_PERSON = 5;            // 企微每人最多展示条数

// ========== Gmail ==========
const GMAIL_CC = "kelland_zhao@colpal.com"; // 所有 Gmail 周报抄送给 Kelland

// ========== 权限表 ==========
const PERMISSION_SPREADSHEET_ID = "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM";
const PERMISSION_SHEET_NAME = "userID";
const PERMISSION_EMAIL_COL_IDX = 9;        // J列（0-indexed）
const PERMISSION_ALERT_HEADER = "SmartMeeting_Alert"; // 第2行表头查找

// ========== onOpen 动态菜单 ==========
function onOpen(e) {
  if (!sbnMenuSet) {
    console.warn("未找到『菜单设置』工作表");
    return;
  }
  const lastRow = sbnMenuSet.getLastRow();
  if (lastRow <= 2) {
    console.log("菜单设置数据不足，跳过菜单生成");
    return;
  }
  const menuData = sbnMenuSet.getRange(3, 1, lastRow - 2, 3).getDisplayValues();
  const enabledMenus = menuData.filter(function(row) { return row[0] && row[2] === "是"; });
  if (enabledMenus.length === 0) {
    console.log("菜单设置中没有启用的菜单项");
    return;
  }

  // 用 Ui.createMenu 替代废弃的 addMenu，installable trigger 上下文更稳定
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("专项菜单");
  enabledMenus.forEach(function(item, idx) {
    if (idx > 0) menu.addSeparator();
    menu.addItem(item[0], item[1]);
  });
  menu.addToUi();
}

/** 手动安装 onOpen 可安装触发器（standalone 脚本首次部署用） */
function installOnOpenTrigger() {
  // 删除已有的 onOpen 触发器
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "onOpen" && t.getEventType() === ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  // 新建针对配置表的 onOpen 触发器
  ScriptApp.newTrigger("onOpen").forSpreadsheet(saas).onOpen().create();
  console.log("✅ onOpen 触发器已安装（删除旧的 " + removed + " 个）");
  console.log("现在刷新配置表，菜单应自动出现");
  try { writeLog("installOnOpenTrigger", "成功", "onOpen 触发器已安装", "手动", ""); } catch (e) {}
}

/** 诊断：列出当前所有触发器 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log("==== 当前触发器列表 (" + triggers.length + " 个) ====");
  if (triggers.length === 0) {
    console.log("⚠️ 没有任何触发器，需要先运行 installOnOpenTrigger 和 timeExec");
    return;
  }
  triggers.forEach(function(t, i) {
    console.log((i + 1) + ". 函数=" + t.getHandlerFunction() +
      " | 类型=" + t.getEventType() +
      " | 来源=" + t.getTriggerSource() +
      " | ID=" + t.getUniqueId());
  });
}
