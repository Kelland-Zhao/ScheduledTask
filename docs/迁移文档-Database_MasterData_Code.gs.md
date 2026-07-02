# 迁移文档：Database_MasterData code.gs → ScheduledScripts

> **源脚本 ID**：`168P7bXzcwBS9b9Fx9u6e-m52zAusUwMED4aOPeCGZLbsuHVIlKT9rpT1`  
> **关联 Google Sheet**：[Database_MasterData](https://docs.google.com/spreadsheets/d/1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U/edit)  
> **源脚本类型**：Container-bound（绑定到 Database_MasterData 表格）  
> **目标项目**：[ScheduledScripts](https://github.com/../ScheduledScripts)（Standalone 脚本，通过 openById 访问表格）  
> **迁移日期**：2026-07-02  
> **版本**：V20260702.01

---

## 一、源脚本概览

| 文件 | 行数 | 用途 |
|------|------|------|
| `0 - 专项菜单.js` | ~80 | onOpen 菜单渲染 + 测试物料邮件发送 |
| `1 - Common.js` | ~150 | 日期格式化 + HTML表格生成 + 邮件发送工具 |
| `2 - 定时执行.js` | ~120 | 定时触发器管理（精确时间触发框架） |
| `3 - Code.js` | ~470 | **核心业务逻辑**：故障代码合并 + 机台汇总 + 逾期催办 |

---

## 二、code.gs（3 - Code.js）功能详解

### 2.1 故障代码合并更新 — `errorCode_update()`

**触发方式**：每周四 10:30 定时执行（定时设置表配置）

**数据流向**：
```
IM源 (1uIEoS1T...)  ─┐
TF源 (1VrXakMRB...)  ─┤→ 合并清洗 → "errorCode 2.0" 工作表
PK源 (13IWm6K59Q...)  ─┘
```

**处理逻辑**：
1. 从 3 个数据源的 `4.Line_UPDT` 工作表读取 A-D 列
2. 根据 Machine_Type 推断 Process 字段（IM/TF/PK）
3. 过滤空行和含"已删除"标记的记录
4. 追加 3 条手动"转规格/Changeover"记录
5. 全量写入目标表 `errorCode 2.0`（清空后重写）

**输出字段**：`Process | Machine_Type | UPDT (English) | UPDT (Local Language) | Downtime Driver`

---

### 2.2 机台汇总更新 — `workcenter_update()`

**触发方式**：每周四 10:30 定时执行

**数据流向**：
```
IM源 (11zyH65MhC...) → "1. Line Database"    D列(Workcenter) B列(Workshop)
TF源 (1Wm_-6j60ZD...) → "1. Active Machine"   D列(Workcenter) B列(Workshop)
PK源 (1736u5O21vH...) → "1. Active Machine"   C列(Workcenter) A列(Workshop)
         ↓ 合并
   "Workcenter_V202602" 工作表
```

**输出字段**：`Workcenter | Workshop | Process`

**⚠️ 与现有脚本的关系**：
- 本项目的 `13 - Workcenter数据同步.js` 是**不同功能**——它从 Line Database 同步到 Equipment_Number_EAM 做字典匹配并写入 Workcenter 表
- `workcenter_update()` 是从三个工序数据源汇总所有活跃机台清单，写入 Database_MasterData 的汇总表
- 两者**功能互补**，不冲突

---

### 2.3 逾期催办扫描 — `scanAndRemindOverdue()`

**触发方式**：周一至周五 10:30 定时执行

**数据流向**：
```
PM_DB (Database_MasterData) → Tasklist_history + Database for Web
点检DB (1RQql-PrcBWi...)     → Tasklist_history + Database for Web
         ↓
   扫描逾期项 → 匹配审批人 → 发送 HTML 邮件
```

**核心逻辑**：
| 状态 | 逾期阈值 |
|------|----------|
| 待审批/ Pending | 3 天 |
| 待发放/ Wait for Dissminater | 5 天 |

**邮件颜色编码**：
- >7 天：红色加粗
- >3 天：橙色
- ≤3 天：默认

**审批链路判定**（`determineResponsibleEmail`）：
```
待发放 → Mail_Disseninate
待审批 + 生产审批=Y + 审批人空 → Mail_Production
待审批 + approver1 空 → Mail_Approve1
待审批 + approver1 有值 + approver2 空 → Mail_Approve2
```

**邮件模板**：HTML 表格 + Tasklist MoC Web App 链接

---

### 2.4 辅助函数

| 函数 | 来源 | 用途 |
|------|------|------|
| `writeLog()` | `3 - Code.js` | 写入 Log 工作表（Database_MasterData） |
| `writeErrorLog()` | `3 - Code.js` | 写入错误日志 |
| `extractTimestampFromOperator()` | `3 - Code.js` | 从操作者字符串中提取时间戳 |
| `determineResponsibleEmail()` | `3 - Code.js` | 根据审批状态确定责任人邮箱 |
| `formatVariableAsDateHms()` | `1 - Common.js` | 日期→字符串格式化 |
| `General_Htmltable()` | `1 - Common.js` | 数据数组→HTML表格 |
| `Mail_HTML_TXT()` | `1 - Common.js` | 发送 HTML 邮件 |

---

## 三、关联 Google Sheet 结构

| 工作表 | 行数 | 用途 |
|--------|------|------|
| Temp | 1000 | 临时数据 |
| HistoryTaskInfo | 11 | 任务历史 |
| errorCode | 483 | 旧版故障代码（已废弃?） |
| **errorCode 2.0** | 1000 | ⭐ 合并后的故障代码（errorCode_update 输出） |
| **Workcenter_V202602** | 1000 | ⭐ 合并后的机台清单（workcenter_update 输出） |
| Workcenter_手动维护 | 1000 | 手动维护的机台数据 |
| Log | 1144 | ⭐ 执行日志 |
| 定时设置 | 5 | ⭐ 定时任务配置表 |
| 菜单设置 | 7 | ⭐ 自定义菜单配置 |
| Workcenter & Mold Matrix | 1187 | 机台模具矩阵 |
| PM Tasklist | 1411 | 保养任务清单 |
| Tasklist_history | 184 | 任务变更历史（催办扫描源） |
| Tasklist Details Bencmark | 9639 | 保养基准明细 |
| Tasklist MoC Report InfomNameList | 996 | MoC 通知名单 |
| Database for Web | 70 | ⭐ Web 应用数据库（含审批人邮箱配置） |
| Authorization settings | 1001 | 权限设置 |
| Production Approval List | 1000 | 生产审批名单 |
| InformNameList | 30 | 通知名单 |

---

## 四、迁移计划

### 4.1 新建脚本

| 序号 | 文件名 | 对应源函数 | 说明 |
|------|--------|-----------|------|
| **32** | `32 - 故障代码合并.js` | `errorCode_update()` | 故障代码从3源合并写入 errorCode 2.0 |
| **33** | `33 - 机台汇总更新.js` | `workcenter_update()` | 机台从3源汇总写入 Workcenter_V202602 |
| **34** | `34 - 逾期催办提醒.js` | `scanAndRemindOverdue()` | PM/点检逾期扫描+邮件提醒 |

### 4.2 需要合并到 `01 - Common.js` 的函数

| 函数 | 优先级 | 说明 |
|------|--------|------|
| `extractTimestampFromOperator()` | 中 | 34 号脚本需要；检查现有 Common.js 是否有类似函数 |
| `formatVariableAsDateHms()` | **低** | 源 Common.js 已有类似函数，检查本项目 `01 - Common.js` 是否已覆盖 |

### 4.3 改造要点

#### 4.3.1 Container-bound → Standalone 改造

所有 `getActiveSpreadsheet()` 改为 `SpreadsheetApp.openById(id)`：
- 源脚本用 `const saas = SpreadsheetApp.getActiveSpreadsheet()` → 改为硬编码 spreadsheet ID
- `currentTimeZone` 来源 `getSpreadsheetTimeZone()` → 改为 `"Asia/Shanghai"`

#### 4.3.2 全局变量处理

源脚本的全局变量：
```javascript
const saas = SpreadsheetApp.getActiveSpreadsheet(); // 顶层调用，需移到函数内
const scriptRunOwner = Session.getActiveUser().getEmail(); // 需移到函数内
```

**注意**：CLAUDE.md 明确要求「避免顶层调用需 OAuth scope 的 API」，这些必须移到函数体内。

#### 4.3.3 writeLog 改造（日志体系统一）

**源脚本日志签名**（Database_MasterData，2个函数）：
```javascript
writeLog(message)        // → Log 表: [时间, 日志信息]          2列
writeErrorLog(message)   // → Log 表: [时间, "", 错误信息]      3列
```

**本项目日志签名**（ScheduledScripts，`01 - Common.js` L53，全局共用）：
```javascript
writeLog(funcName, status, detail, trigger, remark)
// → sbnLog 表: [时间, 函数名, 状态, 详情, 触发方式, 备注]  6列
// sbnLog 定义在 00 - 专项菜单.js，指向 ScheduledScripts 配置表的 Log sheet
```

**迁移映射规则**：

| 源调用 | 迁移后调用 |
|--------|-----------|
| `writeLog("普通日志")` | `writeLog(funcName, "成功", "普通日志", trigger, "")` |
| `writeErrorLog("错误信息")` | `writeLog(funcName, "失败", "错误信息", trigger, "")` |

**关键改造点**：
- **删除**源脚本的 `writeLog()` 和 `writeErrorLog()` 函数定义（约 40 行）
- **复用**项目已有的 `writeLog(funcName, status, detail, trigger, remark)`（已定义在 `01 - Common.js`）
- 每个迁移函数定义 `const funcName = "函数名"` 作为第一个参数
- `trigger` 参数通过入口函数 `e` 判断：`e ? "定时" : "手动"`
- 源脚本中 `[时间戳] 前缀` 格式的日志文本 → 去掉时间前缀，`formatVariableAsDateHms()` 已由本项目的 writeLog 自动追加
- **不再写入** Database_MasterData 的 Log 表（原项目停用，日志统一在 ScheduledScripts）

#### 4.3.4 定时触发配置（✅ 已通过 MCP 配置）

已在 ScheduledScripts 配置表的定时设置中追加 3 条规则：

| 行 | 函数 | 月 | 星期 | 时间 | 程序名 |
|----|------|-----|------|------|------|
| 34 | `errorCode_update` | 1-12 | 4(周四) | 10:30 | 【故障代码】更新 |
| 35 | `workcenter_update` | 1-12 | 4(周四) | 10:31 | 【机台号】更新 |
| 36 | `scanAndRemindOverdue` | 1-12 | 1-5(周一~五) | 10:33 | 【逾期催办】更新 |

菜单设置已同步追加 3 项（故障代码合并更新 / 机台汇总更新 / 逾期催办提醒），均可手动执行。

#### 4.3.5 邮件模板重新设计（follow 邮件UI规范）

源脚本的逾期催办邮件使用简单拼接 HTML，不符合 ScheduledScripts 的邮件 UI 规范。
迁移后 `34 - 逾期催办提醒.js` 的邮件模板需重新设计。

**规范对齐**：参考 `docs/邮件UI规范.md` → 提醒类邮件 → 模式 C + 卡片表格 + 尾部模式 B

**发件人名称**：`逾期催办提醒系统`

**邮件主题**：
```
【逾期提醒】保养/ PM — N 项待处理 / Overdue PM Items — N Pending
【逾期提醒】点检/ Inspection — N 项待处理 / Overdue Inspection Items — N Pending
```

**邮件正文模板**：
```html
<!DOCTYPE html>
<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head>
<body>
<div style="font-family:Arial,'Microsoft YaHei','Helvetica Neue',sans-serif;max-width:900px;margin:0 auto;">

  <!-- 模式 C：白色卡片 + 左侧红色色条 -->
  <div style="background:#ffebee;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);
              padding:30px;margin-bottom:20px;border-left:5px solid #f44336;">
    <h2 style="color:#d32f2f;margin:0 0 8px 0;">
      【逾期提醒】{dbName} 待审批/待发放项
    </h2>
    <p style="color:#c62828;margin:0;font-size:14px;">
      Overdue Approval / Dissemination Items — {N} 项待处理
    </p>
  </div>

  <!-- 正文卡片 -->
  <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);
              padding:30px;margin-bottom:20px;">
    <p style="font-size:15px;color:#34495e;line-height:1.8;">
      您好！以下 {dbName} 待审批/待发放项已逾期，请尽快处理：<br>
      <span style="font-size:13px;opacity:0.8;">Hello! The following {dbName} items are overdue.
      Please process as soon as possible.</span>
    </p>

    <!-- 渐变表头卡片表格 -->
    <div style="overflow-x:auto;margin:24px 0;">
      <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;
                    box-shadow:0 1px 3px rgba(0,0,0,0.1);font-size:13px;">
        <thead>
          <tr style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;">
            <th style="padding:10px;text-align:center;font-weight:600;">
              机型<br><span style="font-size:11px;opacity:0.85;">Machine</span></th>
            <th style="padding:10px;text-align:center;font-weight:600;">
              状态<br><span style="font-size:11px;opacity:0.85;">Status</span></th>
            <th style="padding:10px;text-align:center;font-weight:600;">
              等待天数<br><span style="font-size:11px;opacity:0.85;">Days</span></th>
            <th style="padding:10px;text-align:center;font-weight:600;">
              变更原因<br><span style="font-size:11px;opacity:0.85;">Reason</span></th>
            <th style="padding:10px;text-align:center;font-weight:600;">
              申请人<br><span style="font-size:11px;opacity:0.85;">Applier</span></th>
          </tr>
        </thead>
        <tbody>
          {tableRows}
        </tbody>
      </table>
    </div>

    <p style="margin-top:24px;font-size:14px;color:#34495e;">
      请点击下方链接登录系统处理：<br>
      <span style="font-size:12px;opacity:0.8;">Please click the link below to log in and process:</span>
    </p>
    <a href="{webAppUrl}"
       style="display:inline-block;padding:10px 20px;background:#E60012;color:white;
              text-decoration:none;border-radius:4px;font-size:14px;margin-top:8px;">
      任务清单变更管理 / Tasklist MoC
    </a>
  </div>

  <!-- 尾部 模式 B：双语 -->
  <div style="text-align:center;color:#7f8c8d;font-size:14px;line-height:1.6;">
    <p style="margin-bottom:10px;">
      请及时查看并处理相关逾期事项。<br>
      <span style="font-size:0.9em;opacity:0.8;">
        Please review and handle related overdue items promptly.
      </span>
    </p>
    <p style="margin:0;font-style:italic;">
      此邮件由逾期催办提醒系统自动发送，请勿回复。<br>
      <span style="font-size:0.8em;opacity:0.8;">
        This email is automatically sent by the Overdue Reminder System, please do not reply.
      </span>
    </p>
  </div>

</div>
</body>
</html>
```

**数据行（tableRows）生成规则**：

| 条件 | 行背景 | 天数列 Badge |
|------|--------|-------------|
| daysElapsed > 7 | `#fff5f5` | 渐变红：`linear-gradient(135deg,#e74c3c,#c0392b)`，白色文字，圆角 `16px` |
| daysElapsed 4-7 | `#fffbf0` | 渐变橙：`linear-gradient(135deg,#f39c12,#e67e22)`，白色文字，圆角 `16px` |
| daysElapsed ≤ 3 | `#ffffff`（默认） | 默认文字，无 badge |

**邮件发送参数**：
```javascript
GmailApp.sendEmail(recipientEmail, subject,
  "请使用支持 HTML 的邮件客户端查看此邮件。", {
    htmlBody: htmlBody,
    name: "逾期催办提醒系统"
  });
```

**与规范的对照**：

| 规范项 | 本模块取值 | 规范来源 |
|--------|-----------|---------|
| 邮件宽度 | `900px` | §2.2 提醒类 |
| 头部模式 | C（白色卡片+左色条） | §3 模式 C |
| 表格表头 | 渐变红 `#f44336→#d32f2f` | §6.5 卡片表格 |
| 中英双语表头 | 中文在上、英文在下 | §4.4 |
| 发件人名称 | `逾期催办提醒系统` | §11（中文命名） |
| 主题标签 | `【逾期提醒】` | §10.1 |
| 尾部 | B（双语） | §12.2 |
| 纯文本回退 | `"请使用支持 HTML 的邮件客户端查看此邮件。"` | §14.4 |

### 4.4 依赖的 Google Sheet

| 表格 | ID | 用途 |
|------|-----|------|
| Database_MasterData | `1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U` | 主数据库 |
| IM Line Database | `11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc` | 机台数据(IM) |
| IM 故障代码源 | `1uIEoS1T1Evw0rXa3IUrew91mCffgwf-X_yQ4K4dKDYA` | 故障代码(IM) |
| TF Active Machine | `1Wm_-6j60ZD8KrUMa1gjzuUlRzky1zmc3hrX_u0ZFxT0` | 机台数据(TF) |
| TF 故障代码源 | `1VrXakMRBVcsDO9fAN9JlqLyWq3Sr61dNLqQK8x8OjaM` | 故障代码(TF) |
| PK Active Machine | `1736u5O21vH6Qhw6Uesqlg-uyZ955blUbSRST3xtQGZ8` | 机台数据(PK) |
| PK 故障代码源 | `13IWm6K59Q86Yhw_X84ekTsa8i_fvrR8g6H0QuJtJvkE` | 故障代码(PK) |
| 点检数据库 | `1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY` | 点检 MoC 审批 |

### 4.5 不迁移的部分

| 文件 | 原因 |
|------|------|
| `0 - 专项菜单.js` (onOpen) | ScheduledScripts 已有自己的菜单体系 |
| `0 - 专项菜单.js` (getDataAndSend) | 测试物料领用功能属于 Database_MasterData 专属，与 ScheduledScripts 业务无关 |
| `2 - 定时执行.js` (timeExec/cancelAll) | ScheduledScripts 已有 `02 - 定时执行.js`，定时框架相同 |
| `1 - Common.js` 中的 `Mail_HTML_TXT` | ScheduledScripts 已有自己的邮件发送体系 |

---

## 五、验收标准

- [ ] `32 - 故障代码合并.js` 独立运行成功，errorCode 2.0 数据正确
- [ ] `33 - 机台汇总更新.js` 独立运行成功，Workcenter_V202602 数据正确
- [ ] `34 - 逾期催办提醒.js` 独立运行成功，邮件发送正确（先用测试邮箱验证）
- [ ] 定时触发配置正确，3 个新函数按时执行
- [ ] Log 日志写入正常
- [ ] 源脚本可安全废弃（保留备份）

---

## 五.5、源脚本已知 Bug 及修复方案

### Bug #1：邮箱重复拼接（2026-07-02 确认）

**症状**：`yue_cao@colpal.com@colpal.com` — 收件人邮箱被重复拼接 `@colpal.com`

**根因**：`scanSingleDatabase()` 中：
```javascript
var recipientEmail = recipientPrefix + "@colpal.com";
```
`Database for Web` 工作表的邮件列（Mail_Approve1/2、Mail_Disseninate 等）存储的是**完整邮箱**（`yue_cao@colpal.com`），但代码假设它是用户名前缀，又拼接了一次域名。

**修复方案**：拼接前判断是否已含 `@`：
```javascript
var recipientEmail = recipientPrefix.includes("@") 
  ? recipientPrefix 
  : recipientPrefix + "@colpal.com";
```

**影响范围**：`34 - 逾期催办提醒.js` 迁移时必须应用此修复。

---

## 六、风险提示

1. **逾期催办会真实发送邮件**：测试时需将收件人临时替换为 `kelland_zhao@colpal.com`
2. **errorCode 2.0 和 Workcenter_V202602 被全量覆盖**：确认无其他系统依赖这两个表的增量数据
3. **writeLog 双写**：如果改为 ScheduledScripts 的日志表，Database_MasterData 的 Log 表将不再更新——需确认是否有其他系统读取该 Log 表
4. **6 个外部表格依赖**：需确保 ScheduledScripts 的 Service Account 对这些表格有读取权限
