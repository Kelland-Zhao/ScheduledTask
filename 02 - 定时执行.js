// V20260524.1 — 定时执行模块

/** 取消所有时间驱动触发器 */
function cancelAllTimeDrivenTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (let i = triggers.length - 1; i > -1; i--) {
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  console.log("已删除 " + count + " 个时钟触发器");
  try { writeLog("cancelAllTimeDrivenTriggers", "成功", "删除 " + count + " 个时钟触发器", "手动", ""); } catch (e) {}
}

/** 每天 12:30 自动维护触发器（注册自身 + 立即创建当日触发器） */
function timeExec() {
  cancelAllTimeDrivenTriggers();
  ScriptApp.newTrigger("createSpecialTimeDrivenTriggers")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(30)
    .create();
  createSpecialTimeDrivenTriggers();
  try { writeLog("timeExec", "成功", "已重建触发器（12:30 自维护 + 当日定时）", "手动", ""); } catch (e) {}
}

/** 根据「定时设置」sheet 创建当日（或次日）精准触发器 */
function createSpecialTimeDrivenTriggers() {
  if (!sbnTimingSet) {
    console.warn("未找到『定时设置』工作表");
    return;
  }
  const lastRow = sbnTimingSet.getLastRow();
  if (lastRow <= 2) return;

  const arrSetTime = sbnTimingSet.getRange(3, 1, lastRow - 2, 7).getDisplayValues()
    .filter(function(v) {
      return v[0] !== "" && v[1] !== "" && v[2] !== "" && v[3] !== "" && v[4] !== "" && v[5] !== "" && v[6] !== "";
    });

  if (arrSetTime.length === 0) return;

  const now = new Date();
  const nowYear = now.getFullYear();
  const nowMonth = now.getMonth();
  const nowDay = now.getDate();
  const nowOver30Minute = new Date(now.getTime() + 0.5 * 3600 * 1000);

  for (let i = 0; i < arrSetTime.length; i++) {
    const arrMonth = JSON.parse(arrSetTime[i][0]);
    const arrWeek = JSON.parse(arrSetTime[i][1]);
    const arrWeekInMonth = JSON.parse(arrSetTime[i][2]);
    const arrDay = JSON.parse(arrSetTime[i][3]);
    const strDes = arrSetTime[i][5];
    const strFunction = arrSetTime[i][6];

    cancelAssignstrFunction(strFunction);

    const arrTimes = arrSetTime[i][4].split("|");
    for (let j = 0; j < arrTimes.length; j++) {
      const parts = arrTimes[j].split(":");
      const hr = Number(parts[0]);
      const mn = Number(parts[1]);
      if (hr < 0 || mn < 0 || isNaN(hr) || isNaN(mn)) continue;

      let specificTime = new Date(nowYear, nowMonth, nowDay, hr, mn, 0);
      if (nowOver30Minute > specificTime) {
        specificTime = new Date(nowYear, nowMonth, nowDay + 1, hr, mn, 0);
      }

      const m = specificTime.getMonth() + 1;
      const w = specificTime.getDay();
      const wim = getWeekInMonth(specificTime);
      const d = specificTime.getDate();

      console.log({
        程序: strDes,
        函数: strFunction,
        触发时间: formatVariableAsDateHms(specificTime),
        月匹配: arrMonth.indexOf(m) !== -1,
        周匹配: arrWeek.indexOf(w) !== -1,
        月内周匹配: arrWeekInMonth.indexOf(wim) !== -1,
        日匹配: arrDay.indexOf(d) !== -1
      });

      if (arrMonth.indexOf(m) !== -1 && arrWeek.indexOf(w) !== -1 &&
          arrWeekInMonth.indexOf(wim) !== -1 && arrDay.indexOf(d) !== -1) {
        ScriptApp.newTrigger(strFunction).timeBased().at(specificTime).create();
      }
    }
  }
}

/** 取消指定函数名的所有时钟触发器 */
function cancelAssignstrFunction(strFunction) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = triggers.length - 1; i > -1; i--) {
    if (triggers[i].getHandlerFunction() === strFunction &&
        triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
