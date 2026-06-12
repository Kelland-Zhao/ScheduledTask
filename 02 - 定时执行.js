// V20260613.1 — 定时执行模块（每5分钟轮询调度）

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

/** 每天 12:30 自动维护触发器（每5分钟调度器 + 12:30 自维护） */
function timeExec() {
  cancelAllTimeDrivenTriggers();
  // 每5分钟调度触发器
  ScriptApp.newTrigger("dispatchScheduledTasks")
    .timeBased()
    .everyMinutes(5)
    .create();
  // 每日 12:30 自维护触发器
  ScriptApp.newTrigger("timeExec")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(30)
    .create();
  // 立即执行一次当前窗口的任务
  dispatchScheduledTasks();
  try { writeLog("timeExec", "成功", "已重建触发器（每5分钟调度 + 12:30 自维护）", "手动", ""); } catch (e) {}
}

/** 每5分钟轮询：匹配 (now-5min, now] 窗口内的定时规则，依次执行 */
function dispatchScheduledTasks() {
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
  // 半开区间 (now-5min, now]：上界包含、下界排除，自然去重
  const windowStart = new Date(now.getTime() - 5 * 60 * 1000);

  for (let i = 0; i < arrSetTime.length; i++) {
    const arrMonth = JSON.parse(arrSetTime[i][0]);
    const arrWeek = JSON.parse(arrSetTime[i][1]);
    const arrWeekInMonth = JSON.parse(arrSetTime[i][2]);
    const arrDay = JSON.parse(arrSetTime[i][3]);
    const strDes = arrSetTime[i][5];
    const strFunction = arrSetTime[i][6];

    const arrTimes = arrSetTime[i][4].split("|");
    for (let j = 0; j < arrTimes.length; j++) {
      const parts = arrTimes[j].split(":");
      const hr = Number(parts[0]);
      const mn = Number(parts[1]);
      if (hr < 0 || mn < 0 || isNaN(hr) || isNaN(mn)) continue;

      const ruleTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hr, mn, 0);

      // 仅当规则时间落在 (windowStart, now] 区间内才执行
      if (ruleTime <= windowStart || ruleTime > now) continue;

      const m = ruleTime.getMonth() + 1;
      const w = ruleTime.getDay();
      const wim = getWeekInMonth(ruleTime);
      const d = ruleTime.getDate();

      if (arrMonth.indexOf(m) !== -1 && arrWeek.indexOf(w) !== -1 &&
          arrWeekInMonth.indexOf(wim) !== -1 && arrDay.indexOf(d) !== -1) {
        console.log("[" + formatVariableAsDateHms(now) + "] 执行: " + strDes + " (" + strFunction + ")");
        try {
          if (typeof this[strFunction] === "function") {
            this[strFunction]({ triggerType: "scheduled" });
          }
        } catch (err) {
          console.error(strFunction + " 执行失败: " + err.message);
          try { writeLog("dispatchScheduledTasks", "异常", strDes + " (" + strFunction + ") 失败: " + err.message, "定时", ""); } catch (e) {}
        }
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
