# 34 - 备件到货提醒 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新建 GAS 定时模块「备件到货提醒」——每天 08:30 邮件提醒未领出备件的申请人领用，直至领出自动停止。

**Architecture:** 单文件模块 `34 - 备件到货提醒.js`，遵循项目 33 个既有模块的骨架模式：配置常量 → 主入口（e 参数区分定时/手动）→ writeLog 全记录 → userID 表姓名→邮箱解析 → 按申请人聚合发送 HTML 邮件。只读后端公式表，不做任何写回；无需去重状态——报告 tab 的存在性即状态。

**Tech Stack:** Google Apps Script（V8，clasp 3.3.0 管理）、GmailApp、SpreadsheetApp。无本地测试框架——验证方式为 `clasp push` + `clasp run` 测试入口 + 收件箱/Log 检查。

**Spec:** `docs/需求文档_34.md`（本计划从该 spec 展开；执行者两篇对照阅读）

## Global Constraints

- 文件名必须为 `34 - 备件到货提醒.js`；内部函数前缀 `_sar_`（已确认与现有前缀无冲突）
- 邮件遵循 `docs/邮件UI规范.md`：模式 C（白色卡片+左侧红色色条）、渐变表头卡片表格、发件人名称「备件到货提醒系统」、全角 `【】` 标签、所有动态内容经 `escapeHtml()`、内联样式、DOCTYPE+UTF-8 meta、纯文本回退、中英双语
- 收件人：申请人 TO（userID J列 GMail）+ 上级 CC（userID BI列）+ `kelland_zhao@colpal.com` CC；姓名匹配不到 → 汇总一封发 Kelland + Log 警告
- 每次运行必须 `writeLog()`（成功/失败/警告全记录）
- 后端表 `1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU` 只读，绝不写回
- userID 表：`PERMISSION_SPREADSHEET_ID`/`PERMISSION_SHEET_NAME`（00 文件全局常量），数据从第 3 行起，读 A~BI 61 列；B列(1)=姓名、J列(9)=GMail、BI列(60)=直线上级邮箱
- 定时配置经 MCP 写入配置表（`CONFIG_SPREADSHEET_ID = 1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM`）：「定时设置」每天 08:30 执行 `sendSpareArrivalReminder`，「菜单设置」新增「备件到货提醒-手动」（启用=是）
- 生产调度由既有 5 分钟轮询器驱动（02 文件），**无需重新部署**

---

### Task 1: 模块骨架 + 报告数据读取

**Files:**
- Create: `34 - 备件到货提醒.js`

**Interfaces:**
- Produces:
  - `_sar_normalizeName(name) → String`（小写去后缀键）
  - `_sar_readReportRows(trigger) → Array<Object>`，元素字段：`{code, name, po, applicant, reason, poQty, arrivalDate, daysInRoom, overdueNote, unitPrice, totalPrice, unpickedQty}`（全为字符串）
  - 临时调试入口 `_sar_debugRead()`（Task 4 删除）

- [ ] **Step 1: 创建模块文件（常量 + 姓名预处理 + 报告读取）**

```js
// V20260817.01 — 备件到货提醒（每天 08:30 提醒未领出备件的申请人，直至领出自动停止）
// 入口：sendSpareArrivalReminder（定时/手动）；测试入口：testSpareArrivalReminder（仅发 kelland_zhao@colpal.com）
// 数据源：备件项目采购未领出报告（1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU）→ 未领出备件报告
// 逻辑：报告 tab 的 FILTER 公式只含未领用数量>0 的记录 → 每天提醒，领出后记录消失自动停止
// 收件人：申请人（userID 姓名→J列GMail）TO + 直线上级（BI列）CC + kelland_zhao CC
// 兜底：姓名匹配不到 → 汇总一封发 kelland_zhao@colpal.com + Log 警告
// 邮件：遵循 docs/邮件UI规范.md 模式 C

// ========== 配置常量 ==========
const SAR_CONFIG = {
  REPORT_SPREADSHEET_ID: "1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU",
  REPORT_SHEET_NAME: "未领出备件报告",
  SENDER_NAME: "备件到货提醒系统",
  ADMIN_EMAIL: "kelland_zhao@colpal.com",
};

const SAR_FUNC_NAME = "sendSpareArrivalReminder";
const SAR_EMAIL_COL = 9;      // userID J列 GMail
const SAR_SUPERVISOR_COL = 60; // userID BI列 直线上级邮箱

// ========== 姓名预处理（用于 userID 匹配，同 08 号模块模式）==========
function _sar_normalizeName(name) {
  let s = String(name || '').trim();
  s = s.replace(/-[A-Za-z0-9]+$/, ''); // 去掉 "-Z" 类后缀
  s = s.replace(/_[A-Za-z0-9_]+$/, ''); // 去掉 "_QC" 类后缀
  return s.trim().toLowerCase();
}

// ========== 读取未领出备件报告 ==========
function _sar_readReportRows(trigger) {
  const ss = SpreadsheetApp.openById(SAR_CONFIG.REPORT_SPREADSHEET_ID);
  const ws = ss.getSheetByName(SAR_CONFIG.REPORT_SHEET_NAME);
  if (!ws) {
    writeLog(SAR_FUNC_NAME, "失败", "找不到工作表: " + SAR_CONFIG.REPORT_SHEET_NAME, trigger, "");
    return [];
  }
  const lastRow = ws.getLastRow();
  if (lastRow < 2) {
    writeLog(SAR_FUNC_NAME, "成功", "报告无数据行，跳过", trigger, "");
    return [];
  }
  // 列：A备件# B名称 C-PO D申请人 E采购原因 F采购数量 G到货日期 H-Stock I-Safety Stock J未领用数量 K在备件房天数 L超期提醒 M备注 N单价 O合价
  const data = ws.getRange(2, 1, lastRow - 1, 15).getDisplayValues();
  const rows = [];
  data.forEach(function(r) {
    const code = String(r[0] || "").trim();
    const applicant = String(r[3] || "").trim();
    const qty = Number(String(r[9] || "0").replace(/,/g, ""));
    if (!code || !applicant || !(qty > 0)) return; // 防御性过滤（报告公式已保证，双保险）
    rows.push({
      code: code,
      name: String(r[1] || "").trim(),
      po: String(r[2] || "").trim(),
      applicant: applicant,
      reason: String(r[4] || "").trim(),
      poQty: String(r[5] || "").trim(),
      arrivalDate: String(r[6] || "").trim(),
      daysInRoom: String(r[10] || "").trim(),
      overdueNote: String(r[11] || "").trim(),
      unitPrice: String(r[13] || "").trim(),
      totalPrice: String(r[14] || "").trim(),
      unpickedQty: String(r[9] || "").trim(),
    });
  });
  writeLog(SAR_FUNC_NAME, "成功", "读取到 " + rows.length + " 条未领出记录", trigger, "");
  return rows;
}

// ========== 临时调试入口（Task 4 删除）==========
function _sar_debugRead() {
  const rows = _sar_readReportRows("手动");
  console.log("共 " + rows.length + " 条未领出记录");
  if (rows.length > 0) {
    console.log(JSON.stringify(rows[0]));
    console.log(JSON.stringify(rows[rows.length - 1]));
  }
}
```

- [ ] **Step 2: 推送并运行调试入口**

Run: `clasp push` → 输出 `Pushed N files.`（无报错）
Run: `clasp run _sar_debugRead`
Expected: 输出 `共 72 条未领出记录`（数据量可能随时间变化，>0 即可）+ 首末两条 JSON（字段齐全，arrivalDate 为 `2026-08-17` 格式字符串）+ 配置表 Log sheet 新增 `_sar_debugRead` 相关记录

- [ ] **Step 3: Commit**

```bash
git add "34 - 备件到货提醒.js"
git commit -m "feat(34): 备件到货提醒骨架+报告读取"
```

---

### Task 2: userID 姓名→邮箱映射

**Files:**
- Modify: `34 - 备件到货提醒.js`（追加函数，扩展 `_sar_debugRead`）

**Interfaces:**
- Consumes: 全局常量 `PERMISSION_SPREADSHEET_ID`/`PERMISSION_SHEET_NAME`（00 文件）、`_sar_normalizeName`（Task 1）
- Produces: `_sar_buildUserMap(trigger) → Object`，键=规范化姓名，值=`{email, supervisorEmail}`（小写邮箱字符串）

- [ ] **Step 1: 实现映射函数并扩展调试入口**

```js
// ========== 构建 userID 姓名→邮箱映射 ==========
function _sar_buildUserMap(trigger) {
  const sheet = SpreadsheetApp.openById(PERMISSION_SPREADSHEET_ID).getSheetByName(PERMISSION_SHEET_NAME);
  if (!sheet) {
    writeLog(SAR_FUNC_NAME, "失败", "找不到权限表: " + PERMISSION_SHEET_NAME, trigger, "");
    return {};
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    writeLog(SAR_FUNC_NAME, "失败", "权限表数据不足（lastRow=" + lastRow + "）", trigger, "");
    return {};
  }
  // 读取 A~BI（61列），从第3行开始（跳过2行表头），同 08 号模块模式
  const data = sheet.getRange(3, 1, lastRow - 2, 61).getValues();
  const map = {};
  data.forEach(function(row) {
    const key = _sar_normalizeName(row[1]); // B列 姓名
    if (!key) return;
    const email = String(row[SAR_EMAIL_COL] || "").trim().toLowerCase();
    const supervisorEmail = String(row[SAR_SUPERVISOR_COL] || "").trim().toLowerCase();
    if (!map[key]) {
      map[key] = { email: email, supervisorEmail: supervisorEmail };
    } else if (!map[key].email && email) {
      map[key].email = email; // 首条无邮箱时补
    }
  });
  writeLog(SAR_FUNC_NAME, "成功", "userID 映射构建完成，共 " + Object.keys(map).length + " 人", trigger, "");
  return map;
}
```

同时把 `_sar_debugRead` 改为（保留原输出，追加映射检查）：

```js
function _sar_debugRead() {
  const rows = _sar_readReportRows("手动");
  console.log("共 " + rows.length + " 条未领出记录");
  if (rows.length > 0) {
    console.log(JSON.stringify(rows[0]));
  }
  const userMap = _sar_buildUserMap("手动");
  console.log("userID 映射人数: " + Object.keys(userMap).length);
  const sampleNames = {};
  rows.forEach(function(r) {
    const key = _sar_normalizeName(r.applicant);
    if (!sampleNames[key]) sampleNames[key] = { applicant: r.applicant, u: userMap[key] };
  });
  Object.keys(sampleNames).forEach(function(key) {
    console.log(key + " → " + JSON.stringify(sampleNames[key].u));
  });
}
```

- [ ] **Step 2: 推送并运行**

Run: `clasp push`
Run: `clasp run _sar_debugRead`
Expected: `userID 映射人数: >0`；报告中每个申请人（韩善明/丁志伟/付彪/张嘉骏/齐兵/袁建兵/李华/许强/罗红兵/范朝锋/万苏华/林飞/赵凯 等）都有映射条目且 `email` 非空；个别缺失的（若有）记为 Task 4 兜底验证素材

- [ ] **Step 3: Commit**

```bash
git add "34 - 备件到货提醒.js"
git commit -m "feat(34): userID姓名邮箱映射"
```

---

### Task 3: 邮件正文生成

**Files:**
- Modify: `34 - 备件到货提醒.js`（追加 `_sar_buildEmailBody`、`_sar_buildUnmatchedEmailBody`）

**Interfaces:**
- Consumes: `escapeHtml`（01 文件全局函数）、Task 1 的行对象结构
- Produces:
  - `_sar_buildEmailBody(applicant, items, email) → String`（HTML 邮件正文）
  - `_sar_buildUnmatchedEmailBody(list) → String`，`list` 为 `[{applicant, items}]` 数组

- [ ] **Step 1: 实现正文生成函数**

```js
// ========== 邮件正文（模式 C：白色卡片+色条 + 渐变表头卡片表格）==========
function _sar_buildEmailBody(applicant, items, email) {
  var tableRows = "";
  items.forEach(function(it) {
    var daysCell;
    if (String(it.daysInRoom).trim() === "0") {
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:#27ae60;color:white;font-weight:600;font-size:12px;">今日到货<br><span style="font-size:10px;opacity:0.9;">Arrived Today</span></span>';
    } else {
      daysCell = '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:linear-gradient(135deg,#f39c12,#e67e22);color:white;font-weight:600;font-size:12px;">' + escapeHtml(it.daysInRoom) + ' 天<br><span style="font-size:10px;opacity:0.9;">Days</span></span>';
    }
    var overdueCell = it.overdueNote
      ? '<span style="display:inline-block;padding:4px 10px;border-radius:16px;background:linear-gradient(135deg,#e74c3c,#c0392b);color:white;font-weight:600;font-size:12px;">' + escapeHtml(it.overdueNote) + '</span>'
      : '—';
    var reasonCell = it.reason === "紧急采购"
      ? '<span style="display:inline-block;padding:4px 8px;border-radius:12px;background:#f39c12;color:white;font-size:12px;">紧急采购</span>'
      : escapeHtml(it.reason || "—");
    var priceCell = it.unitPrice ? (escapeHtml(it.unitPrice) + " / " + escapeHtml(it.totalPrice || "—")) : "—";
    tableRows += '<tr style="background-color:#ffffff;">' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + escapeHtml(it.code) + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:260px;word-wrap:break-word;">' + escapeHtml(it.name || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.po || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + reasonCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.arrivalDate || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + daysCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">' + overdueCell + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.unpickedQty || "—") + '</td>' +
      '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + priceCell + '</td>' +
      '</tr>';
  });

  var htmlBody =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">' +

    // 模式 C：白色卡片 + 左侧红色色条
    '<div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;border-left:5px solid #E60012;">' +
    '<h2 style="color:#d32f2f;margin:0 0 8px 0;">【备件到货提醒】您有备件待领用</h2>' +
    '<p style="color:#c62828;margin:0;font-size:14px;">Spare Parts Pickup Reminder — ' + items.length + ' 件待领用</p>' +
    '</div>' +

    // 正文卡片
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;">' +
    '<p style="font-size:15px;color:#34495e;line-height:1.8;">' +
    escapeHtml(applicant) + '，您好！以下备件已到货/仍在备件房，请尽快到备件房领用：<br>' +
    '<span style="font-size:13px;opacity:0.8;">Hello! The following spare parts have arrived / are waiting in the spare parts room. Please pick them up as soon as possible.</span>' +
    '</p>' +

    // 渐变表头卡片表格
    '<div style="overflow-x:auto;margin:24px 0;">' +
    '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.1);font-size:13px;">' +
    '<thead><tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">' +
    '<th style="padding:10px;text-align:left;font-weight:600;">备件编码<br><span style="font-size:11px;opacity:0.85;">Code</span></th>' +
    '<th style="padding:10px;text-align:left;font-weight:600;">备件名称<br><span style="font-size:11px;opacity:0.85;">Name</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">PO</th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">采购原因<br><span style="font-size:11px;opacity:0.85;">Reason</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">到货日期<br><span style="font-size:11px;opacity:0.85;">Arrival</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">在备件房<br><span style="font-size:11px;opacity:0.85;">Days</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">超期提醒<br><span style="font-size:11px;opacity:0.85;">Overdue</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">未领数量<br><span style="font-size:11px;opacity:0.85;">Qty</span></th>' +
    '<th style="padding:10px;text-align:center;font-weight:600;">单价/合价<br><span style="font-size:11px;opacity:0.85;">Price</span></th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table></div>' +
    '</div>' +

    // 尾部 模式 B：双语
    '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;">' +
    '<p style="margin-bottom:10px;">请及时到备件房领用相关备件。<br>' +
    '<span style="font-size:0.9em;opacity:0.8;">Please pick up the spare parts at the spare parts room promptly.</span></p>' +
    '<p style="margin:0;font-style:italic;">此邮件由备件到货提醒系统自动发送，请勿回复。<br>' +
    '<span style="font-size:0.8em;opacity:0.8;">This email is automatically sent by the Spare Parts Arrival Reminder System, please do not reply.</span></p>' +
    '</div></div></body></html>';

  return htmlBody;
}

// ========== 未匹配人员汇总邮件正文 ==========
function _sar_buildUnmatchedEmailBody(list) {
  var tableRows = "";
  list.forEach(function(g) {
    g.items.forEach(function(it) {
      tableRows += '<tr style="background-color:#ffffff;">' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">' + escapeHtml(it.code) + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:260px;word-wrap:break-word;">' + escapeHtml(it.name || "—") + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.po || "—") + '</td>' +
        '<td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;color:#34495e;">' + escapeHtml(it.applicant) + '</td>' +
        '</tr>';
    });
  });

  var htmlBody =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:900px;margin:0 auto;">' +
    '<div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);' +
    'padding:30px;margin-bottom:20px;border-left:5px solid #E60012;">' +
    '<h2 style="color:#d32f2f;margin:0 0 8px 0;">【备件到货提醒】未匹配人员汇总</h2>' +
    '<p style="color:#c62828;margin:0;font-size:14px;">以下申请人在 userID 表中未找到，无法发送提醒，请人工转告</p>' +
    '</div>' +
    '<div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">' +
    '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<thead><tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">' +
    '<th style="padding:10px;text-align:left;">备件编码</th><th style="padding:10px;text-align:left;">备件名称</th>' +
    '<th style="padding:10px;text-align:center;">PO</th><th style="padding:10px;text-align:center;">申请人</th>' +
    '</tr></thead><tbody>' + tableRows + '</tbody></table></div></div>' +
    '<div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;margin-top:20px;">' +
    '<p style="margin:0;font-style:italic;">此邮件由备件到货提醒系统自动发送，请勿回复。</p></div>' +
    '</div></body></html>';

  return htmlBody;
}
```

同时把 `_sar_debugRead` 末尾追加正文预览（放在映射输出之后）：

```js
  const firstKey = Object.keys(sampleNames)[0];
  if (firstKey) {
    const items = rows.filter(function(r) {
      return _sar_normalizeName(r.applicant) === firstKey;
    });
    const body = _sar_buildEmailBody(sampleNames[firstKey].applicant, items, "debug");
    console.log("正文长度: " + body.length);
    console.log(body.substring(0, 500));
  }
```

- [ ] **Step 2: 推送并运行**

Run: `clasp push`
Run: `clasp run _sar_debugRead`
Expected: 输出 `正文长度: >3000` 且预览含 `【备件到货提醒】`、渐变表头、`今日到货` badge（若有 K=0 行）；无报错

- [ ] **Step 3: Commit**

```bash
git add "34 - 备件到货提醒.js"
git commit -m "feat(34): 邮件正文生成（模式C卡片+双语）"
```

---

### Task 4: 主入口 + 测试入口 + 端到端验证

**Files:**
- Modify: `34 - 备件到货提醒.js`（追加 `_sar_run`/`sendSpareArrivalReminder`/`testSpareArrivalReminder`，**删除** `_sar_debugRead`）

**Interfaces:**
- Consumes: Task 1/2/3 全部函数、`writeLog`/`GmailApp`
- Produces: `sendSpareArrivalReminder(e)`（定时入口）、`testSpareArrivalReminder()`（测试入口，仅发 kelland_zhao@colpal.com）

- [ ] **Step 1: 实现主入口并删除调试入口**

```js
// ========== 主入口 ==========
function sendSpareArrivalReminder(e) {
  _sar_run(e ? "定时" : "手动", false);
}

/** 测试入口：仅发送至 kelland_zhao@colpal.com，主题加【测试】前缀 */
function testSpareArrivalReminder() {
  _sar_run("手动", true);
}

function _sar_run(trigger, testMode) {
  const startTime = new Date();
  try {
    writeLog(SAR_FUNC_NAME, "成功", "开始扫描备件未领出报告（testMode=" + testMode + "）", trigger, "");

    const rows = _sar_readReportRows(trigger);
    if (rows.length === 0) {
      writeLog(SAR_FUNC_NAME, "成功", "无未领出备件记录，跳过发送", trigger, "");
      return;
    }

    const userMap = _sar_buildUserMap(trigger);

    // 按申请人分组（保留首个原始姓名用于显示）
    const byName = {};
    rows.forEach(function(r) {
      const key = _sar_normalizeName(r.applicant);
      if (!byName[key]) byName[key] = { applicant: r.applicant, items: [] };
      byName[key].items.push(r);
    });

    let sentCount = 0;
    const unmatched = [];
    Object.keys(byName).forEach(function(key) {
      const group = byName[key];
      const u = userMap[key];
      if (!u || !u.email) {
        unmatched.push(group);
        return;
      }
      const htmlBody = _sar_buildEmailBody(group.applicant, group.items, u.email);
      const subject = (testMode ? "【测试】" : "") +
        "【备件到货提醒】" + group.applicant + "，您申请的 " + group.items.length + " 件备件待领用";
      const ccList = [];
      if (u.supervisorEmail && u.supervisorEmail !== u.email) ccList.push(u.supervisorEmail);
      if (ccList.indexOf(SAR_CONFIG.ADMIN_EMAIL) === -1) ccList.push(SAR_CONFIG.ADMIN_EMAIL);
      const toEmail = testMode ? SAR_CONFIG.ADMIN_EMAIL : u.email;
      try {
        GmailApp.sendEmail(toEmail, subject, "请使用支持 HTML 的邮件客户端查看此邮件。", {
          htmlBody: htmlBody,
          name: SAR_CONFIG.SENDER_NAME,
          cc: testMode ? "" : ccList.join(","),
        });
        sentCount++;
        writeLog(SAR_FUNC_NAME, "成功", "已发送提醒至 " + u.email + "（" + group.items.length + " 件）", trigger, "");
      } catch (mailErr) {
        writeLog(SAR_FUNC_NAME, "失败", "发送邮件至 " + u.email + " 失败: " + mailErr.message, trigger, "");
      }
    });

    // 未匹配兜底：汇总一封发 Kelland
    if (unmatched.length > 0) {
      const names = unmatched.map(function(g) { return g.applicant; }).join("、");
      const unmatchedHtml = _sar_buildUnmatchedEmailBody(unmatched);
      const unmatchedSubject = (testMode ? "【测试】" : "") + "【备件到货提醒】" + unmatched.length + " 位申请人未匹配到 userID，请人工转告";
      try {
        GmailApp.sendEmail(SAR_CONFIG.ADMIN_EMAIL, unmatchedSubject, "请使用支持 HTML 的邮件客户端查看此邮件。", {
          htmlBody: unmatchedHtml,
          name: SAR_CONFIG.SENDER_NAME,
        });
        writeLog(SAR_FUNC_NAME, "警告", "以下申请人未匹配到 userID，已汇总发 Kelland: " + names, trigger, "");
      } catch (mailErr) {
        writeLog(SAR_FUNC_NAME, "失败", "未匹配汇总邮件发送失败: " + mailErr.message, trigger, "");
      }
    }

    const endTime = new Date();
    const duration = ((endTime.getTime() - startTime.getTime()) / 1000).toFixed(1);
    writeLog(SAR_FUNC_NAME, "成功", "备件到货提醒完成，耗时 " + duration + " 秒，发送 " + sentCount + " 封（未匹配 " + unmatched.length + " 组）", trigger, "");
  } catch (err) {
    writeLog(SAR_FUNC_NAME, "失败", err.message, trigger, err.stack || "");
  }
}
```

删除 `_sar_debugRead` 整个函数及其注释。

- [ ] **Step 2: 推送并端到端测试**

Run: `clasp push`
Run: `clasp run testSpareArrivalReminder`
Expected:
- 输出无报错
- `kelland_zhao@colpal.com` 收到多封 `【测试】【备件到货提醒】...` 邮件（每位申请人一封，内容为其名下全部未领出件；今日到货行绿色 badge、超期行红色 badge、紧急采购橙色 badge 正确显示；单价缺失行显示 `—`）
- 配置表 Log sheet 新增该次运行记录：开始/读取 N 条/映射人数/每封发送/完成汇总
- 若有未匹配申请人 → 收到 `【测试】【备件到货提醒】N 位申请人未匹配...` 汇总邮件 + Log 警告行

- [ ] **Step 3: Commit**

```bash
git add "34 - 备件到货提醒.js"
git commit -m "feat(34): 主入口+测试入口+未匹配兜底"
```

---

### Task 5: 配置表登记（经 MCP）

**Files:**
- 无代码改动；修改 Google Sheet 配置表 `1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM`

**Interfaces:**
- Consumes: 02 文件调度器对「定时设置」的解析格式：A=月(JSON数组) B=星期(JSON数组) C=当月第几周(JSON数组) D=日(JSON数组) E=时间("HH:MM") F=描述 G=函数名；00 文件 onOpen 对「菜单设置」的解析格式：A=显示名 B=函数名 C=是否启用("是")

- [ ] **Step 1: 读现有行确认格式**

用 MCP（google-sheet-mcp CLI，service account 凭据，env `GOOGLE_SPREADSHEET_ID=1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM`）读「定时设置」与「菜单设置」前几行：

```bash
npx -y google-sheet-mcp read -s "定时设置" -r "A3:G5" --raw
npx -y google-sheet-mcp read -s "菜单设置" -r "A3:C5" --raw
```

Expected: 确认 7 列/3 列格式与现有 08:30 行（如 `sendDailyProcessReport`）的 JSON 数组写法一致。
若 service account 无配置表访问权限（403/404）→ 停止，报告用户：需将配置表分享给 service account（Editor）。

- [ ] **Step 2: 追加「定时设置」行**

用 MCP append 追加一行（值对齐全列）：

```
A(月): [1,2,3,4,5,6,7,8,9,10,11,12]
B(星期): [0,1,2,3,4,5,6]
C(当月第几周): [1,2,3,4,5]
D(日): [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
E(时间): 08:30
F(描述): 备件到货提醒
G(函数名): sendSpareArrivalReminder
```

- [ ] **Step 3: 追加「菜单设置」行**

```
A(显示名): 备件到货提醒-手动
B(函数名): sendSpareArrivalReminder
C(是否启用): 是
```

- [ ] **Step 4: 读回验证**

Run: 再次 `read` 两个 sheet 末几行
Expected: 新行存在且格式与其他行一致；无重复追加。验证后告知用户：次日 08:30 起 5 分钟轮询器将自动执行（无需重新部署）。

---

### Task 6: 变更日志 + 提交推送 GitHub

**Files:**
- Modify: `docs/项目记录.md`（顶部插入新版本条目）

**Interfaces:**
- Consumes: 无

- [ ] **Step 1: 在 docs/项目记录.md 顶部插入变更条目**

在 `# SmartMeeting_Alert 项目记录` 标题之后插入：

```markdown
## V20260817.01 — 新增备件到货提醒（34号项目）

**改动内容**：
- 新增文件 `34 - 备件到货提醒.js`，入口 `sendSpareArrivalReminder(e)`，测试入口 `testSpareArrivalReminder()`
- **功能**：每天 08:30 扫描备件项目采购未领出报告的「未领出备件报告」tab（`1I56UtusP9E2JHh1LsRvPY_VMmkv5_9t_Q_yDkf86MUU`），提醒未领出备件的申请人到备件房领用，直至领出（记录消失）自动停止
- **触发逻辑**：报告 tab 公式已只含未领用数量>0 的记录；不做日期窗口过滤，用 K列（在备件房天数，外部程序每日更新）展示天数，0=今日到货
- **收件人**：申请人（userID 姓名→J列GMail）TO + 直线上级（BI列）CC + kelland_zhao@colpal.com CC；姓名未匹配 → 汇总发 Kelland + Log 警告
- **邮件**：遵循 邮件UI规范 模式 C，发件人「备件到货提醒系统」，中英双语；今日到货绿色 badge、超期提醒红色 badge、紧急采购橙色 badge
- 配置表 `定时设置` 新增每天 08:30 触发（`sendSpareArrivalReminder`，经 MCP 配置）
- 配置表 `菜单设置` 新增菜单项 `备件到货提醒-手动`（经 MCP 配置）
- 内部函数统一 `_sar_` 前缀
```

- [ ] **Step 2: 提交并推送 GitHub（push-to-github skill）**

```bash
git add -A
git commit -m "V2026-08-17.01_新增备件到货提醒模块"
git push origin master
```

Expected: push 成功，GitHub 与本地一致。
