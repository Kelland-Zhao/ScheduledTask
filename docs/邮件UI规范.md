# 邮件 UI 规范 · ScheduledScripts

> 适用项目：ScheduledScripts（GAS 定时脚本集）  
> 适用范围：所有通过 `GmailApp.sendEmail()` 发送的 HTML 邮件  
> 版本：V20260620.02  
> 维护人：Kelland Zhao

---

## 目录

1. [品牌标识](#1-品牌标识)
2. [邮件整体结构](#2-邮件整体结构)
3. [头部（Header）模式](#3-头部header模式)
4. [字体排版](#4-字体排版)
5. [色彩体系](#5-色彩体系)
6. [表格样式](#6-表格样式)
7. [状态指示器](#7-状态指示器)
8. [卡片组件](#8-卡片组件)
9. [标题行（Section Title）](#9-标题行section-title)
10. [邮件主题行](#10-邮件主题行)
11. [发件人名称](#11-发件人名称)
12. [尾部（Footer）](#12-尾部footer)
13. [测试模式](#13-测试模式)
14. [编码规范](#14-编码规范)
15. [模块对照表](#15-模块对照表)

---

## 1. 品牌标识

### 1.1 品牌色

| 属性 | 值 | 用途 |
|------|-----|------|
| **主色（Colgate Red）** | `#E60012` | 头部背景、表格表头、分区标题边框、强调色 |
| **深红** | `#C00010` | 多层表头第二层（IoT 报警报告） |
| **浅红背景** | `#FFF9F9` | 车间表格外层容器 |
| **浅红行背景** | `#fde0e0` | 表头下方副标题行 |
| **浅粉行背景** | `#fff8f8` | 数据行交替色 |
| **合计行背景** | `#fff3f3` | 汇总/小计行 |
| **警告背景** | `#fff5f5` | 异常统计条背景 |

### 1.2 Logo

```javascript
var _dr_COLGATE_LOGO = "https://upload.wikimedia.org/wikipedia/commons/thumb/8/86/Colgate-Palmolive_%282025%29.svg/500px-Colgate-Palmolive_%282025%29.svg.png";
```

**使用位置**：[22 - 注塑工序日报](22%20-%20注塑工序日报.js)、[27 - 注塑工序日报](27%20-%20注塑工序日报.js)、[32 - 注塑排版自动提醒](32%20-%20注塑排版自动提醒.js) 邮件头部右侧（36px 高）。

**规范**：
- 日报/周报/排班提醒等正式对外邮件**必须包含 Logo**
- Logo 高度统一为 `36px`，`display:block`
- Logo 放在头部右侧，与左侧文字对齐（参考 §3 模式A 含Logo变体）
- 简单提醒类（Milestone 提醒、故障报告提醒等）**不包含 Logo**

---

## 2. 邮件整体结构

### 2.1 标准模板

```html
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  <div style="font-family: Arial, 'Microsoft YaHei', 'Helvetica Neue', sans-serif; max-width: 960px; margin: 0 auto;">
    
    <!-- 头部 -->
    <div style="background: #E60012; color: white; padding: 16px 24px;">
      ...
    </div>
    
    <!-- 正文 -->
    <div style="padding: 24px;">
      ...
    </div>
    
    <!-- 尾部 -->
    <p style="color: #bdc3c7; font-size: 11px; margin-top: 32px;">
      此邮件由 [系统名称] 自动发送
    </p>
  </div>
</body>
</html>
```

### 2.2 布局规则

| 属性 | 规则 |
|------|------|
| 最大宽度 | `960px`（日报类）/ `900px`（提醒类）/ `1200px`（宽表类） |
| 居中方式 | `margin: 0 auto` |
| 内边距 | 头部 `16px 24px`，正文 `20px 28px` 或 `24px` |
| 字体族 | `Arial, 'Microsoft YaHei', 'Helvetica Neue', sans-serif` |
| 背景色 | 透明（默认白色邮件背景） |

### 2.3 模块邮件分类

| 类型 | 宽度 | 代表模块 |
|------|------|---------|
| **日报** | 960px | 27-注塑工序日报 |
| **周报/月报** | 900px | 06-SmartMeeting月报, 22-故障报告周报 |
| **汇总报告** | 900~960px | 07-故障报告通知, 17-质量巡场, 14-点检核对 |
| **宽表日报** | 1200px | 31-注塑测试日报 |
| **简单提醒** | 无 max-width | 15-Milestone提醒, 21-超期日报, 23-跟进项提醒 |

---

## 3. 头部（Header）模式

### 模式 A：标准红底白字（推荐，最常用）

**适用**：日报、月报、汇总类邮件（模块 06/14/17/27）

```html
<div style="background: #E60012; color: white; padding: 16px 24px;">
  <h2 style="margin: 0;">邮件标题</h2>
  <p style="margin: 8px 0 0; opacity: 0.95; font-size: 14px;">副标题/说明</p>
  <p style="margin: 4px 0 0; opacity: 0.7; font-size: 12px;">发送时间：2026-06-20</p>
</div>
```

**规格**：
- 背景：`#E60012`
- 主标题：`h2`，`font-size: 20px`，`margin: 0`
- 副标题：`opacity: 0.95`，`font-size: 14px`
- 时间戳：`opacity: 0.7`，`font-size: 12px`

**含 Logo 变体**（模块 22/27/32 使用）：

```html
<div style="background: #E60012; color: white; padding: 14px 28px;">
  <table border="0" cellpadding="0" cellspacing="0" style="width: 100%;">
    <tr>
      <td style="vertical-align: middle;">
        <h2 style="margin: 0; font-size: 20px;">标题</h2>
        <p style="margin: 4px 0 0; opacity: 0.85; font-size: 12px;">发送时间：...</p>
      </td>
      <td style="width: 50px; vertical-align: middle; text-align: right; padding-left: 16px;">
        <img src="[LOGO_URL]" style="height: 36px; display: block;" alt="Colgate">
      </td>
    </tr>
  </table>
</div>
```

### 模式 B：渐变红底 + 圆角（仅模块 07）

```html
<div style="background: linear-gradient(135deg, #E60012 0%, #FF6B6B 100%); 
            color: white; padding: 30px; border-radius: 15px; text-align: center; 
            box-shadow: 0 8px 25px rgba(230, 0, 18, 0.3);">
  ...
</div>
```

### 模式 C：白色卡片 + 左侧色条（提醒类）

**适用**：Milestone 提醒、故障报告提醒（模块 15/21/23）

```html
<div style="background: #ffebee; border-radius: 8px; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; 
            margin-bottom: 20px; border-left: 5px solid #f44336;">
  ...
</div>
```

### 模式 D：大圆角红底 + 统计摘要条（模块 31）

```html
<div style="background: #E60012; color: #fff; padding: 20px; 
            border-radius: 8px 8px 0 0;">
  <h1 style="margin: 0; font-size: 24px;">标题 / English Title</h1>
  <div style="margin-top: 8px; font-size: 14px;">日报日期 / Report Date: ...</div>
</div>
<div style="padding: 14px 20px; background: #fff5f5; border: 1px solid #f3cccc;">
  昨日 N · 明日 M · 异常 K
</div>
```

---

## 4. 字体排版

### 4.1 字体族

**统一字体栈**（优先级从高到低）：

```
font-family: Arial, 'Microsoft YaHei', 'Helvetica Neue', sans-serif;
```

> **注意**：`Microsoft YaHei`（微软雅黑）用于中文显示。部分模块省略了 `'Helvetica Neue'`，统一建议使用完整字体栈。

### 4.2 字号层级

| 层级 | 字号 | 用途 |
|------|------|------|
| H1 | `24px` | 日报主标题（仅模块 31） |
| H2 | `20px` | 邮件头部标题 |
| H3 | `17px` | 正文分区标题 |
| H4 | `15px` | 子分区标题 |
| 正文 | `13px~14px` | 表格内容、说明文字 |
| 辅助 | `12px` | 双列表头英文行（模块 31）、时间戳 |
| 小字 | `11px` | 尾部声明 |
| 大数 | `28px~36px` | 摘要卡片数值 |

### 4.3 字重

| 元素 | font-weight |
|------|-------------|
| 表格表头 | `bold` / `600` |
| 卡片数值 | `bold` |
| 合计行 | `bold` |
| 异常文本 | `bold` |
| 分区标题 | `700`（模块 31）/ 默认 bold |
| 正文 | `normal`（默认） |

### 4.4 中英文双行表头（模块 31 专用）

表头单元格内中文在上、英文在下：

```html
<th style="...">
  [中文列名]<br>
  <span style="font-weight: 400;">[ENGLISH COLUMN NAME]</span>
</th>
```

分区标题同样双行：

```html
<div style="...">
  [中文标题]<br>
  <span style="font-weight: 400;">[ENGLISH TITLE]</span> · 2026-06-20
</div>
```

---

## 5. 色彩体系

### 5.1 品牌色

| 色值 | 名称 | 用途 |
|------|------|------|
| `#E60012` | Colgate Red | 表头背景、头部背景、分区标题左边框/下边框、强调文字 |
| `#C00010` | Dark Red | 多层表头第二层 |

### 5.2 语义色

| 色值 | 语义 | 使用场景 |
|------|------|---------|
| `#27ae60` | 成功/绿色 | 完成率 ≥ 80%、0 异常、"全部完成"提示 |
| `#e67e22` | 警告/橙色 | 完成率 60~79%、临期提醒、差异类型1 |
| `#f39c12` | 次警告/琥珀 | 待审核 badge 背景、中度过期 |
| `#e74c3c` | 危险/红色 | 完成率 < 60%、超期 badge、差异类型2 |
| `#d32f2f` | 深红 | 超期 badge 渐变深色端 |
| `#FFC107` | 高亮黄 | 排班缺失预警单元格、备件库存不足行 |

### 5.3 中性色

| 色值 | 用途 |
|------|------|
| `#2c3e50` | 卡片数值深色、表格主文字 |
| `#34495e` | 正文段落 |
| `#7f8c8d` | 辅助说明文字、卡片标签 |
| `#bdc3c7` | 尾部声明文字 |
| `#95a5a6` | 次级辅助文字 |
| `#6c757d` | 分区标题装饰文字（模块 31） |

### 5.4 背景色

| 色值 | 用途 |
|------|------|
| `#f8f9fa` | 偶数行交替背景（buildHtmlTable）、模块 22 邮件外层 |
| `#ffffff` | 奇数行背景、卡片背景 |
| `#f9f9f9` | 故障表格偶数行（模块 07） |
| `#FFF0F0` | 故障行背景（红色调）— 存在未跟进项 |
| `#F0FFF0` | 正常行背景（绿色调）— 全部跟进 |
| `#fce8e6` | 测试异常行背景（模块 31） |
| `#e6f4ea` | 测试正常行背景（模块 31） |
| `#fff5f5` | 异常统计条背景（模块 31） |
| `#ffebee` | 提醒卡片背景（模块 15/21/23） |
| `#fffbf0` | 待审核表格偶数行（模块 07） |
| `#fff8e1` | 待审核表格悬停行（模块 07） |
| `#f0f0f0` | 合计行背景（IoT 报告） |

### 5.5 渐变

| 渐变 | 用途 |
|------|------|
| `linear-gradient(135deg, #E60012 0%, #FF6B6B 100%)` | 模块 07 头部 |
| `linear-gradient(135deg, #f44336, #d32f2f)` | 超期 badge |
| `linear-gradient(135deg, #f39c12, #e67e22)` | 一般超期 badge |
| `linear-gradient(135deg, #e74c3c, #c0392b)` | 严重超期 badge |
| `linear-gradient(135deg, #667eea 0%, #764ba2 100%)` | 周报表头（模块 22） |
| `linear-gradient(135deg, #3f51b5, #283593)` | 验证提醒表头（模块 23） |

### 5.6 行背景状态映射

| 条件 | 行背景色 | 说明 |
|------|---------|------|
| 有异常 / 有问题 | `#fce8e6` | 测试日报异常行 |
| 正常 / 无问题 | `#e6f4ea` | 测试日报正常行 |
| 有未跟进 | `#FFF0F0` | IoT 报警分类行 |
| 全部跟进 | `#F0FFF0` | IoT 报警分类行 |
| 无人排班预警 | `#FFC107`（单元格） | 工序日报特定列 |
| 奇数行 | `#ffffff` | 标准表格 |
| 偶数行 | `#f8f9fa` | 标准表格 |

---

## 6. 表格样式

### 6.1 标准表格（`buildHtmlTable`）

**位置**：[01 - Common.js](../01%20-%20Common.js) 第 181 行  
**使用模块**：06, 09, 14, 17

```javascript
function buildHtmlTable(headers, rows, headerBg) {
  // headerBg 统一传入 "#E60012"
}
```

**样式规格**：

| 属性 | 值 |
|------|-----|
| 边框 | `1px solid`，`border-collapse: collapse` |
| 宽度 | `100%` |
| 字号 | `13px` |
| 表头背景 | `#E60012` |
| 表头文字 | 白色，`text-align: left`，`padding: 10px` |
| 奇数行背景 | `#ffffff` |
| 偶数行背景 | `#f8f9fa` |
| 单元格内边距 | `10px` |
| HTML 转义 | 所有值经过 `escapeHtml()` |

### 6.2 故障报告双表格（模块 07）

**故障表格**（`.fault-table`）：
- 表头：`background: #E60012`，`border: 1px solid #d32f2f`，`padding: 12px 8px`
- 偶数行：`#f9f9f9`
- 悬停行：`#f0f0f0`

**待审核表格**（`.review-table`）：
- 表头：`background: #f39c12`，`border: 1px solid #e67e22`，`padding: 12px 8px`
- 偶数行：`#fffbf0`
- 悬停行：`#fff8e1`

### 6.3 车间排班表格（模块 27）

**外层容器**：`background: #FFF9F9; border: 1px solid #f0d0d0; border-radius: 4px`

**车间名表头**：`background: #E60012; text-align: center; padding: 8px; font-size: 15px`

**列标题行**：`background: #fde0e0; border-bottom: 1px solid #f0d0d0`

**合计行**：`background: #fff3f3; border-top: 2px solid #E60012; color: #E60012; font-weight: bold`

**预警单元格**：`background: #FFC107`（开机数 > 0 但人数 = 0 时）

### 6.4 宽表（模块 31）

**表头**：`padding: 5px 6px; border: 1px solid #ddd; background: #E60012; color: #fff; font-size: 12px; text-align: center`

**数据行**：`padding: 5px 6px; border: 1px solid #ddd; font-size: 12px; text-align: center`

**行背景**：
- 有问题：`#fce8e6`
- 正常：`#e6f4ea`

### 6.5 卡片式表格（模块 15/22/23）

**外层容器**：`background: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px`

**表头**：渐变背景 + `border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1)`

**单元格**：`padding: 12px; border-bottom: 1px solid #e9ecef`

### 6.6 IoT 报警汇总表（模块 30）

**特点**：`rowspan`/`colspan` 多层表头
- 第一层：报警类型（rowspan="2"），TB1 / TB2 / 合计（各 colspan="3"）
- 第二层：总数 / 未跟进 / 跟进率 × 3
- 第一层背景：`#E60012`
- 第二层背景：`#C00010`
- 行背景：`#FFF0F0`（有未跟进）/ `#F0FFF0`（全部跟进）
- 合计行背景：`#f0f0f0`

### 6.7 表格通用规范

| 规范项 | 规则 |
|--------|------|
| 表头颜色 | 统一使用 `#E60012`（特殊语义表头除外） |
| 表头文字颜色 | `white` / `#fff` |
| 奇数/偶数行 | 必须有交替色，提升可读性 |
| 表格内容转义 | 所有动态内容必须经 `escapeHtml()` |
| 表格响应式 | 使用 `overflow-x: auto` 包裹 |
| 空值显示 | 显示 `-` 或 `—`（全角破折号） |
| 合计行 | 位于表格末尾，`font-weight: bold`，顶部 `2px solid #E60012` 分隔线 |

---

## 7. 状态指示器

### 7.1 Badge（故障报告超期）

**超期 ≥ 7 天**（严重）：
```html
<div style="background: linear-gradient(135deg, #e74c3c, #c0392b); color: white; 
            padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; 
            box-shadow: 0 2px 6px rgba(231, 76, 60, 0.3); display: inline-block; 
            min-width: 80px; text-align: center;">
  <span style="display: block;">[超期] Nd天</span>
  <span style="display: block; font-size: 10px; opacity: 0.9;">Days</span>
</div>
```

**一般超期**：
```html
<div style="background: linear-gradient(135deg, #f39c12, #e67e22); color: white; 
            padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; 
            box-shadow: 0 2px 6px rgba(243, 156, 18, 0.3); display: inline-block; 
            min-width: 80px; text-align: center;">
  <span style="display: block;">Nd天</span>
  <span style="display: block; font-size: 10px; opacity: 0.9;">Days</span>
</div>
```

### 7.2 状态 Badge（模块 07）

```html
<!-- 故障报告状态 -->
<span style="background: #FF6B6B; color: white; padding: 4px 8px; 
             border-radius: 12px; font-size: 12px;">待判断</span>

<!-- 待审核 -->
<span style="background: #f39c12; color: white; padding: 4px 8px; 
             border-radius: 12px; font-size: 12px;">待审核</span>
```

### 7.3 测试检查结果 Badge（模块 31）

```html
<!-- 有问题 -->
<span style="display: inline-block; padding: 3px 8px; border-radius: 10px; 
             color: #b3261e; background: #fce8e6; font-size: 12px;">
  ⚠ N项
</span>

<!-- 正常 -->
<span style="display: inline-block; padding: 3px 8px; border-radius: 10px; 
             color: #137333; background: #e6f4ea; font-size: 12px;">
  ✓ 正常
</span>
```

### 7.4 状态图标/文字

| 场景 | 显示 | 颜色 |
|------|------|------|
| 0 异常 / 全部完成 | `★ [描述]` | `#27ae60` |
| 有未解决问题 | 数量 badge | `#e67e22` 或 `#e74c3c` |
| 无数据 | `无` | `#999` / `#bdc3c7` |
| 不适用 | `—`（全角破折号） | 默认 |

---

## 8. 卡片组件

### 8.1 摘要卡片（模块 06/14）

**用途**：在邮件正文顶部以大字展示核心指标（总数、完成数、完成率等）。

```html
<td style="text-align: center; padding: 16px; border: 1px solid #ecf0f1;">
  <div style="font-size: 28px; font-weight: bold; color: [语义色];">[数值]</div>
  <div style="color: #7f8c8d; font-size: 12px; margin-top: 4px;">[标签]</div>
</td>
```

**布局**：4 卡片并排（`width: 100%` 的 `<table>` 包含一个 `<tr>`）

| 卡片 | 颜色 |
|------|------|
| 总跟进项 | `#2c3e50`（深色） |
| 已完成 | `#27ae60`（绿色） |
| 未完成 | `#e74c3c`（红色） |
| 完成率 | 动态（≥80% 绿 / 60~79% 橙 / <60% 红） |

**完成率色阶**：
- ≥ 80% → `#27ae60`
- 60%~79% → `#e67e22`
- < 60% → `#e74c3c`

### 8.2 双卡片（模块 14）

```html
<td style="text-align: center; padding: 16px; border: 1px solid #ecf0f1; width: 50%;">
  <div style="font-size: 32px; font-weight: bold; color: [语义色];">[数值]</div>
  <div style="color: #7f8c8d; font-size: 13px; margin-top: 6px; line-height: 1.5;">[标签]</div>
</td>
```

---

## 9. 标题行（Section Title）

### 9.1 标准分区标题（模式 A）

```html
<h3 style="color: #E60012; border-left: 4px solid #E60012; 
           padding-left: 8px; border-bottom: 2px solid #E60012; 
           padding-bottom: 8px; margin-top: 32px;">
  [标题] (N项)
</h3>
```

**使用模块**：06, 14, 17

### 9.2 简化分区标题

```html
<h3 style="color: #333; border-bottom: 2px solid #E60012; 
           padding-bottom: 8px; margin-top: 0; font-size: 17px;">
  一、[标题]
</h3>
```

**使用模块**：27

### 9.3 IoT 报告分区标题

```html
<h3 style="color: #E60012; margin: 16px 0 4px;">
  一、[标题]
</h3>
```

**使用模块**：30

### 9.4 中英双语分区标题（模块 31）

```html
<div style="font-size: 13px; font-weight: 700; color: #6c757d; 
            border-left: 3px solid #E60012; padding-left: 10px; 
            margin: 18px 0 10px; letter-spacing: 0.5px;">
  [中文标题]<br>
  <span style="font-weight: 400;">[ENGLISH TITLE]</span> · [日期]
</div>
```

---

## 10. 邮件主题行

### 10.1 格式规范

**统一格式**：`【标签】【空格】【具体信息】`

**强制规则**：所有邮件主题和正文中的标签/前缀**必须使用全角方括号 `【...】`**，禁止使用半角 `[...]`。

| 模块 | 标签 | 格式 |
|------|------|------|
| 06 | `【SmartMeeting Alert】` | `【SmartMeeting Alert】 INJ SDM 跟进完成率月报 - {monthLabel}` |
| 07 | `【故障报告提醒】` | `【故障报告提醒】 {date} - {process}工序发现N个故障待判断和M个故障报告待审核` |
| 09 | `【辅导与反馈提醒】` | `【辅导与反馈提醒】 {process} 工序 - 待提交 N 份 \| Coaching Form Reminder - {process} - N Pending` |
| 14 | `【点检核对】` | `【点检核对】 注塑机台差异报告 {date}` |
| 15 | `【项目逾期】/【项目临期】` | `【项目逾期】有 N 个Milestone已逾期` |
| 17 | `【质量巡场】` | `【质量巡场】 未解决问题汇总 {date}` |
| 21 | `【紧急】` | `【紧急】超期故障报告提醒 - {displayName}({process})` |
| 22 | `【提醒】` | `【提醒】 故障报告定期提醒 - {process} / Failure Report Weekly Reminder - {process}` |
| 23 | `【逾期提醒】/【临期提醒】/【验证提醒】` | `【逾期提醒】故障报告跟进项目逾期 / Follow-up Items Overdue` |
| 26 | `【排班缺失提醒】` | `【排班缺失提醒】 注塑工序人员排班未填写` |
| 27 | `【注塑工序日报】` | `【注塑工序日报】 {date} 注塑工序日报` |
| 30 | `【注塑Opera/ IoT工艺报警监控报告】` | `【注塑Opera/ IoT工艺报警监控报告】{date}` |
| 31 | `【注塑测试日报】` | `【注塑测试日报】{date} 昨日复盘 & 明日提醒` |

### 10.2 标签规范

| 标签类型 | 符号 | 含义 |
|---------|------|------|
| 日报 | `【...日报】` | 每日定时发送的常规日报 |
| 紧急提醒 | `【紧急】` | 超期 ≥ 7 天的严重事项 |
| 逾期提醒 | `【逾期提醒】` | 已逾期的跟进项 |
| 临期提醒 | `【临期提醒】` | 即将到期的项目 |
| 普通提醒 | `【提醒】` / `【xxx提醒】` | 例行周报/通知 |
| 系统错误 | `【系统错误】` | 脚本执行异常通知 |
| 测试 | `【测试】`（前缀） | 测试模式邮件主题前缀 |

### 10.3 双语主题

部分模块（09, 21, 22, 23）使用中英双语主题，格式：`中文 / English`。

---

## 11. 发件人名称

| 发件人名称 | 使用模块 |
|-----------|---------|
| `SmartMeeting Alert` | 01, 04, 06, 09 |
| `PointCheck Alert` | 14 |
| `Quality Tour Alert` | 17 |
| `项目Milestone提醒系统` | 15 |
| `故障报告提醒系统` | 21, 22, 23 |
| `工时数据汇总系统` | 26 |
| `工序日报系统` | 27 |
| `GAS报警监控系统` | 30 |
| `注塑测试日报` | 31 |

**规范**：
- 发件人名称使用中文（仅 SmartMeeting/PointCheck/Quality Tour 沿用英文历史命名）
- 新模块统一使用中文名称
- 模块 31 额外配置了 `from` 别名（`CSX_PlantSystem@colpal.com`）

---

## 12. 尾部（Footer）

### 12.1 标准尾部（模式 A）

```html
<p style="color: #bdc3c7; font-size: 11px; margin-top: 32px;">
  此邮件由 [系统名称] 自动发送
</p>
```

**使用模块**：06, 14, 17, 27

### 12.2 双语尾部（模式 B）

```html
<div style="text-align: center; color: #7f8c8d; font-size: 14px; line-height: 1.6;">
  <p style="margin-bottom: 10px;">
    请及时查看并处理相关[事项名称]。<br>
    <span style="font-size: 0.9em; opacity: 0.8;">Please review and handle related [items] promptly.</span>
  </p>
  <p style="margin: 0; font-style: italic;">
    此邮件由系统自动发送，请勿回复。<br>
    <span style="font-size: 0.8em; opacity: 0.8;">
      This email is automatically sent by the system, please do not reply.
    </span>
  </p>
</div>
```

**使用模块**：22（部分翻译为英文的提醒模块）

### 12.3 简洁双语尾部（模式 C）

```html
<div style="margin-top: 20px; color: #777; font-size: 11px;">
  此邮件由 ScheduledScripts 自动生成 / Generated automatically.
</div>
```

**使用模块**：31

### 12.4 尾部规范

| 规范项 | 规则 |
|--------|------|
| 颜色 | `#bdc3c7` 或 `#777` |
| 字号 | `11px` |
| 上边距 | `32px`（标准）/ `20px`（紧凑） |
| 内容 | 必须包含"自动发送 / 请勿回复"语义 |
| 国际化 | 中文优先，英文可加注 |

---

## 13. 测试模式

### 13.1 通用测试模式

```javascript
var _dr_TEST_MODE = false;  // 生产环境
var _dr_TEST_EMAIL = "kelland_zhao@colpal.com";
```

**行为**：
- 测试模式下邮件仅发送至 `TEST_EMAIL`
- 邮件主题前加 `[测试]` 前缀
- 生产部署前必须将 `TEST_MODE` 设为 `false`

### 13.2 模块 31 测试模式

```javascript
var ITR_CONFIG = {
  TEST_EMAIL: "kelland_zhao@colpal.com",
  FROM_EMAIL: "CSX_PlantSystem@colpal.com"
};
// 测试模式通过 options.testMode 参数控制
```

---

## 14. 编码规范

### 14.1 HTML 转义

所有来自数据的动态内容必须经过 `escapeHtml()` 转义：

```javascript
function escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
```

**模块 31 使用增强版**（额外转义 `"` 和 `'`）：

```javascript
function _itr_escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
```

### 14.2 DOCTYPE

所有 HTML 邮件必须包含：

```html
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
  ...
</body>
</html>
```

### 14.3 邮件发送

**标准发送模式**：

```javascript
GmailApp.sendEmail(to, subject, "", {
  htmlBody: htmlBody,
  name: "发件人名称"
});
```

**带 CC**：

```javascript
GmailApp.sendEmail(to, subject, "", {
  htmlBody: htmlBody,
  name: "发件人名称",
  cc: ccList.join(",")
});
```

**带 from 别名**（需先验证别名可用）：

```javascript
var aliases = GmailApp.getAliases() || [];
if (aliases.indexOf(ITR_CONFIG.FROM_EMAIL) >= 0) {
  options.from = ITR_CONFIG.FROM_EMAIL;
}
```

### 14.4 纯文本回退

`sendEmail` 第三个参数（`body`）提供纯文本回退内容，供不支持 HTML 的邮件客户端使用：

```javascript
GmailApp.sendEmail(to, subject, "请使用支持 HTML 的邮件客户端查看。", options);
```

### 14.5 CC 规范

| 场景 | CC 对象 |
|------|--------|
| 所有 Gmail 报告 | `kelland_zhao@colpal.com`（全局 GMAIL_CC） |
| 故障报告通知 | 管理员 + 审批人直线上级（supervisor） |
| 辅导提醒 | 管理员 + 直线上级 |
| Milestone 提醒 | 项目管理员 + 直线上级 |

### 14.6 内联样式

**强制要求**：所有 CSS 必须使用内联样式（`style="..."`），禁止使用 `<style>` 标签或外部 CSS。

> Google Gmail 会剥离 `<style>` 标签和外部样式表，仅保留内联样式。

少数模块（07）使用 `<style>` 定义了 class 选择器（如 `.fault-table`、`.review-table`），这在新模块中**不再推荐**，应统一改为内联样式。

---

## 15. 模块对照表

| 文件 | 系统名称 | 邮件频率 | 头部模式 | 表格样式 | 尾部模式 |
|------|---------|---------|---------|---------|---------|
| 06-Gmail月报 | SmartMeeting Alert | 每月 | A（红底白字） | buildHtmlTable + 摘要卡片 | A |
| 07-故障报告通知 | 故障报告提醒（无 sender name） | 每 5 分钟 | B（渐变红底） | 故障表格 + 待审核表格 | — |
| 09-辅导与反馈提醒 | SmartMeeting Alert | 每天 | 内联简单头部 | buildHtmlTable | A |
| 14-点检机台核对 | PointCheck Alert | 每天 | A（红底白字） | buildHtmlTable + 双卡片 | A |
| 15-新品自动化CI项目邮件提醒 | 项目Milestone提醒系统 | 每天 | C（白色卡片+色条） | 渐变表头卡片表格 | — |
| 17-质量巡场问题提醒 | Quality Tour Alert | 每天 | A（红底白字） | buildHtmlTable | A |
| 21-故障报告7天未完每日提醒 | 故障报告提醒系统 | 每天 | C（白色卡片+色条） | 渐变表头卡片表格 | B（双语） |
| 22-故障报告周报 | 故障报告提醒系统 | 每周 | C（白色卡片+色条） | 渐变表头卡片表格 | B（双语） |
| 23-故障报告跟进项提醒 | 故障报告提醒系统 | 每天 | C（白色卡片+色条） | 渐变表头卡片表格 | B（双语） |
| 26-工时数据汇总 | 工时数据汇总系统 | 每天（按需） | 内联简单头部 | 简单表格 | — |
| 27-注塑工序日报 | 工序日报系统 | 每天 | A（含 Logo） | 车间排班表 + 备件表 | A |
| 30-IoT报警监控 | GAS报警监控系统 | 每天 | A（渐变表头） | 多层表头 + 柱状图 | — |
| 31-注塑测试日报 | 注塑测试日报 | 每天 | D（大圆角+统计条） | 中英双语宽表 | C（双语简洁） |
| 32-注塑排版自动提醒 | 考勤排班提醒系统 | 每天（周一~周五） | A（含 Logo） | 标准表格（§6.1） | A |

---

## 变更记录

| 日期 | 版本 | 变更内容 |
|------|------|---------|
| 2026-06-20 | V20260620.01 | 初始版本，基于现有 13 个邮件模块归纳编写 |
| 2026-06-20 | V20260620.02 | 统一规范：所有邮件标签强制使用全角方括号 `【...】`，全量替换半角 `[...]` |
