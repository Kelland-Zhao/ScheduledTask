# 点检机台核对与主数据双向同步 实施计划

> **注意（V20261001.28）**：本文件是**按当时设计记录的历史文档**，描述的是"当初打算怎么建"，不是当前状态。
> 其中提到的**预演入口 `dryRunPointCheckSync`（菜单「点检同步预演」）已于 V20261001.27 移除**——原因是它唯一的输出是邮件，每次点击都会给 S&C 全组发信（详见 `docs/项目记录.md` 与 `docs/需求文档_13.md` §16）。
> **实现以代码与 `docs/需求文档_13.md` 为准。**

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 `13 - 点检机台核对.js` 从"只比对报警"升级为"比对后自动双向同步 MachineList 主数据"，判定依据收敛为表11 的 `无需检查Y/N` 单一开关。

**Architecture:** 全部逻辑在现有文件 `13 - 点检机台核对.js` 内扩展（遵循本仓库"一个模块号一个文件"的约定）。把所有**决策逻辑**抽成纯函数（可单测），**IO 操作**（删行/追加/标黄/发信）保持薄壳，由真实表验证覆盖。纯函数用 `node --test` + `eval(源码)` 测试，沿用现有 `tests/pointcheck-workcenter.test.mjs` 的写法。

**Tech Stack:** Google Apps Script（ES5 风格 JS，`const` + `function`）、`node:test` + `node:assert/strict`、clasp。

**Spec:** `docs/需求文档_13.md`

## Global Constraints

以下约束来自 spec，**每个任务都隐含包含**，逐条照抄不改动：

- **安全不变量**：只有 `工序 = INJ` 的行可被写入（值和背景色）。PK / TF 等其他工序的行零触碰。
- 删除按 **(工序=INJ, 机台号)** 匹配，**不可只按机台号**——表内 `T1GR0301` 在 TF 下有两行，只按机台号会跨工序误删。
- 追加一律 `insertRowsAfter(最后一行, n)`，**不插到任何现有行中间**。
- 标黄**只写「目标色 ≠ 当前色」的连续区间**；颜色没变的行一概不写。
- 纳入集 = `Workcenter` 全部行 − `无需检查Y/N = Y`。
- 机台号比对归一：`trim()` + `toUpperCase()`；**写入时保留表11 原值**。
- 车间：机台号**从左数第 2 个字符**，`0`/`1` → `TB1`，`2` → `TB2`，其他 → 留空并计入待补全（不猜）。
- `机型(C列) = Plasma` 的 INJ 行**不删除**。
- 安全阀：拟删机台数 **> 10** → 删除不执行（追加照常执行）；手动运行（`e === undefined`）绕过安全阀。
- 安全阀计数单位是**机台号个数**，不是行数——`H2FCS506` 占 2 行只计 1 台。
- Workcenter 列定位一律**按表头名**，不硬编码列号。
- 变量与函数前缀 `_pc_`。
- 入口函数名 `checkPointCheckMachines` **不变**（`定时设置` 无需改动）；新增预演入口 `dryRunPointCheckSync`。
- 定时 08:25；发件人名称 `PointCheck Alert`。
- 提交信息遵循本仓库约定：`V<YYYYMMDD>.<nn>_<描述>`。**版本号在提交时取「当前未被占用的下一个编号」，不要照抄计划里其他任务示例中的号** —— 修复轮同样消耗编号，一个任务可能要提交多次。计划里的示例号仅示格式。

## Review Focus

以下 5 类输入/失效模式是 spec 隐含要求、但任务测试最容易漏掉的，**每条都已在对应任务里配了测试**：

1. **INJ 行的机台号为空或仅空格** → 合理预期：不参与增删、不报错、不被静默删除（不能因为匹配不上就删）。
2. **机台号长度 < 2**（如 `H`）→ 合理预期：车间留空 + 计入待补全，**不抛异常**（`charAt(1)` 返回空串）。
3. **Workcenter 同一机台号出现两次** → 合理预期：取第一条 + 邮件告警，且**只追加一次**。
4. **`无需检查Y/N` 值为小写 `y` 或 `是`** → 合理预期：**不当作 Y**（不排除），但必须出现在邮件告警里，不静默失效。
5. **背景色为 `null` 与具体色值混合** → 合理预期：`null` 视为无色，需要标黄时能正常标；已标黄的能正常清除。

---

## File Structure

| 文件 | 动作 | 职责 |
|---|---|---|
| `13 - 点检机台核对.js` | 修改 | 全部逻辑：判定、同步计划、安全阀、写入、邮件 |
| `tests/pointcheck-workcenter.test.mjs` | 修改 | Workcenter 列定位回归 + 纳入集判定（原有 3 条断言中 1 条随语义变更） |
| `tests/pointcheck-sync.test.mjs` | 新建 | 车间推导、同步计划、安全阀、待补全、标黄区间、邮件 |
| `docs/项目记录.md` | 修改 | 追加本次变更记录 |

**不新建 .js 文件**——本仓库约定一个模块号一个文件。

---

## Task 1: 纳入集判定改造为「无需检查Y/N」

替换掉 v1 的「`机器性能` 关键字匹配」判据。这是本次变更的地基：后面所有增删都建立在"哪些机台在纳入集里"之上。

**Files:**
- Modify: `13 - 点检机台核对.js`（新增常量与两个函数；删除 `_pc_EXCLUDE_KEYWORD`、`_pc_buildWorkcenterMap`；改写 `checkPointCheckMachines` 中读取 Workcenter 的那一段）
- Test: `tests/pointcheck-workcenter.test.mjs`（第 49-72 行的两条断言随语义变更；新增归一化与重复键用例）

**Interfaces:**
- Consumes: 无（起点）
- Produces:
  - `_pc_normalizeMachineNo(value: any) → string` —— `trim()` + `toUpperCase()`，非字符串输入安全
  - `_pc_buildIncludedSet(dataWC: Array<Array>) → { included: {[machineNo]: {machineType: string, machineModel: string}}, flagged: {[machineNo]: true}, missing: string[], duplicates: string[], badFlags: string[] }`
  - 常量 `_pc_HEADER_FLAG = "无需检查Y/N"`、`_pc_FLAG_YES = "Y"`

- [ ] **Step 1: 写失败测试（改 `tests/pointcheck-workcenter.test.mjs`）**

**先改第 25-47 行的两条列定位测试**——它们调用的是即将删除的 `_pc_buildWorkcenterMap`，必须改成新函数名与新返回结构（这两条是 19 列结构的回归守卫，保留价值不变）：

```javascript
test('按表头名定位：machineModel 取 Final Machine Type 列', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', 'Machine Type': 'HT160', 'Final Machine Type': '6AX', 'New Formed Cell': 'NFC1' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.equal(r.included['M1'].machineType, 'HT160', 'Machine Type 列');
  assert.equal(r.included['M1'].machineModel, '6AX', 'Final Machine Type 列');
  assert.deepEqual(r.missing, []);
});

test('machineModel 不取自 New Formed Cell 列', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', 'New Formed Cell': '6AX', 'Final Machine Type': 'DP' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.equal(r.included['M1'].machineModel, 'DP', '按 D 列误读的话这里会是 6AX');
});
```

**然后把第 49-72 行两条 `机器性能` 相关断言替换为以下内容**，并在文件末尾追加后四条：

```javascript
test('无需检查Y/N = Y 的行被排除出纳入集', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'H2FTA001', '无需检查Y/N': 'Y' }),
    row({ 'Workcenter': 'H2FTA002', '无需检查Y/N': '' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['H2FTA002']);
  assert.equal(r.flagged['H2FTA001'], true, '被排除的机台要能分辨是「因Y」');
});

test('无需检查Y/N 小写 y / 中文「是」不当作 Y，但计入告警', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'H2FTA001', '无需检查Y/N': 'y' }),
    row({ 'Workcenter': 'H2FTA002', '无需检查Y/N': '是' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included).sort(), ['H2FTA001', 'H2FTA002'], '不能被当成Y排除掉');
  assert.deepEqual(r.badFlags.sort(), ['y', '是'], '必须告警，不能静默失效');
});

test('表头缺少 无需检查Y/N → 返回空并报出缺失字段', () => {
  const broken = WC_HEADERS.filter(h => h !== '无需检查Y/N');
  const data = [broken, row({ 'Workcenter': 'H2FTA001' })];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), [], '宁可空手而归，也不能读错列');
  assert.deepEqual(r.missing, ['无需检查Y/N']);
});

test('Workcenter 重复机台号：取第一条并告警，只出现一次', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'H2FTA001', 'Machine Type': 'HT160', 'Final Machine Type': '6AX' }),
    row({ 'Workcenter': 'H2FTA001', 'Machine Type': 'HT250', 'Final Machine Type': 'FT400' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['H2FTA001'], 'key 只有一个');
  assert.equal(r.included['H2FTA001'].machineType, 'HT160', '取第一条');
  assert.deepEqual(r.duplicates, ['H2FTA001']);
});

test('机台号归一：前后空格与大小写差异视为同一台', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': ' h2fta001 ', '无需检查Y/N': 'Y' }),
    row({ 'Workcenter': 'h2fta001' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), [], '两条归一后都是 H2FTA001，且第一条已标Y，第二条属重复键');
  assert.equal(r.flagged['H2FTA001'], true);
  assert.deepEqual(r.duplicates, ['H2FTA001'], '第二条是重复键（不是 Y 候选，因为已在 flagged 里）');
});

test('机台号为空/仅空格的 Workcenter 行被跳过', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': '' }),
    row({ 'Workcenter': '   ' }),
    row({ 'Workcenter': 'H2FTA001' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['H2FTA001']);
  assert.deepEqual(r.duplicates, [], '空机台号不算重复');
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-workcenter.test.mjs`
Expected: FAIL —— `TypeError: globalThis._pc_buildIncludedSet is not a function`

- [ ] **Step 3: 改 `13 - 点检机台核对.js`——新增常量**

把第 8-23 行「数据源配置」区块里这两行删掉：

```javascript
const _pc_SHEET_WORKCENTER = "Workcenter";
const _pc_EXCLUDE_KEYWORD = "闲置";       // Workcenter 机器性能列排除关键字
```

改为：

```javascript
const _pc_SHEET_WORKCENTER = "Workcenter";

// ========== 纳入集判定的表头名（按名定位，不硬编码列号）==========
const _pc_HEADER_WORKCENTER = "Workcenter";
const _pc_HEADER_MACHINE_TYPE = "Machine Type";
const _pc_HEADER_FINAL_TYPE = "Final Machine Type";
const _pc_HEADER_FLAG = "无需检查Y/N";
const _pc_FLAG_YES = "Y";
```

- [ ] **Step 4: 改 `13 - 点检机台核对.js`——用新函数替换 `_pc_buildWorkcenterMap`**

删除第 35-60 行的 `_pc_buildWorkcenterMap` 整个函数（含 jsdoc），替换为：

```javascript
/** 机台号比对归一：去空格 + 转大写。非字符串输入安全返回 "" */
function _pc_normalizeMachineNo(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim().toUpperCase();
}

/**
 * 解析 Workcenter，产出纳入集与数据质量问题
 * 纳入集 = 全部行 − (「无需检查Y/N」= Y)；判据按表头名定位列
 * @param {Array<Array>} dataWC getDataRange().getValues() 全量（含表头行）
 * @returns {{included: Object, flagged: Object, missing: Array<string>, duplicates: Array<string>, badFlags: Array<string>}}
 *   missing 非空时调用方应跳过，不要用空 included 继续
 */
function _pc_buildIncludedSet(dataWC) {
  const required = [_pc_HEADER_WORKCENTER, _pc_HEADER_MACHINE_TYPE, _pc_HEADER_FINAL_TYPE, _pc_HEADER_FLAG];
  const cols = _pc_headerIndex(dataWC[0] || []);
  const missing = required.filter(function (n) { return cols[n] === undefined; });
  if (missing.length > 0) return { included: {}, flagged: {}, missing: missing, duplicates: [], badFlags: [] };

  const included = {};
  const flagged = {};
  const duplicates = [];
  const badFlags = [];

  for (let i = 1; i < dataWC.length; i++) {
    const machineNo = _pc_normalizeMachineNo(dataWC[i][cols[_pc_HEADER_WORKCENTER]]);
    if (!machineNo) continue;

    const rawFlag = String(dataWC[i][cols[_pc_HEADER_FLAG]] || "").trim();

    // 精确匹配 `Y`，**不做大小写归一**。
    // 用小写 y 去「排除」一台机是危险方向：静默漏检，而本模块的全部目的就是确保点检覆盖。
    // 无法解析的值一律当作「未标 Y」→ 机台被纳入，同时进 badFlags 让邮件喊出来。
    if (rawFlag !== "" && rawFlag !== _pc_FLAG_YES && badFlags.indexOf(rawFlag) < 0) {
      badFlags.push(rawFlag);
    }

    if (rawFlag === _pc_FLAG_YES) {
      flagged[machineNo] = true;
      delete included[machineNo];   // Y 优先：同号任一行标 Y，整台机就不点检
      continue;
    }

    // 已被标 Y 的号优先于其他行，不能因为后面还有一行没标 Y 就把它捞回纳入集
    if (Object.prototype.hasOwnProperty.call(flagged, machineNo)) {
      duplicates.push(machineNo);
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(included, machineNo)) {
      duplicates.push(machineNo);   // 重复键取第一条（沿用模块 12 做法）
      continue;
    }

    included[machineNo] = {
      machineType: String(dataWC[i][cols[_pc_HEADER_MACHINE_TYPE]] || "").trim(),
      machineModel: String(dataWC[i][cols[_pc_HEADER_FINAL_TYPE]] || "").trim(),
    };
  }

  return { included: included, flagged: flagged, missing: [], duplicates: duplicates, badFlags: badFlags };
}
```

- [ ] **Step 5: 改 `checkPointCheckMachines`——接上新的判定函数**

把第 107-114 行这段：

```javascript
    const built = _pc_buildWorkcenterMap(dataWC, _pc_EXCLUDE_KEYWORD);
    if (built.missing.length > 0) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 表头缺少字段: " + built.missing.join(", "), trigger, "");
      return;
    }
    const wcMap = built.map;  // Workcenter → { machineType, machineModel }

    console.log("Workcenter 有效行数(排除闲置): " + Object.keys(wcMap).length);
```

改为：

```javascript
    const built = _pc_buildIncludedSet(dataWC);
    if (built.missing.length > 0) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 表头缺少字段: " + built.missing.join(", "), trigger, "");
      return;
    }
    const wcMap = built.included;  // Workcenter → { machineType, machineModel }

    console.log("Workcenter 纳入集行数(排除无需检查): " + Object.keys(wcMap).length);
    if (built.badFlags.length > 0) console.warn("无需检查Y/N 异常取值: " + built.badFlags.join(", "));
    if (built.duplicates.length > 0) console.warn("Workcenter 重复机台号: " + built.duplicates.join(", "));
```

同一函数里第 122-138 行的差异计算暂不改动（Task 3、Task 6 会整体替换），但它引用的 `wcSet` 现在来自纳入集，语义已经正确。

- [ ] **Step 6: 跑测试确认通过**

Run: `node --test tests/pointcheck-workcenter.test.mjs`
Expected: PASS，8 条全绿

- [ ] **Step 7: 确认没有残留引用**

Run: `grep -n "_pc_buildWorkcenterMap\|_pc_EXCLUDE_KEYWORD" "13 - 点检机台核对.js" tests/*.mjs`
Expected: 无输出（全部清干净）

- [ ] **Step 8: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-workcenter.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.07_点检同步_纳入集改为「无需检查Y/N」单一判据

废弃 机器性能 关键字匹配与 6AX/DP/HS 豁免名单；
新增 badFlags（标了却没生效的值）与 duplicates（重复机台号）告警。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: 车间推导

**Files:**
- Modify: `13 - 点检机台核对.js`（追加一个函数）
- Test: `tests/pointcheck-sync.test.mjs`（新建）

**Interfaces:**
- Consumes: `_pc_normalizeMachineNo`（Task 1）
- Produces: `_pc_deriveWorkshop(machineNo: string) → "TB1" | "TB2" | ""`

- [ ] **Step 1: 新建 `tests/pointcheck-sync.test.mjs` 并写失败测试**

```javascript
// 模块 13 点检机台核对 — 主数据双向同步逻辑测试
// 运行：node --test tests/pointcheck-sync.test.mjs
// fixture 来源均标注在生产表取样位置，不得改为编造数据
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// 模块依赖 01 - Common.js 的公共函数，但 Common.js 顶层会 SpreadsheetApp.openById()，
// 在 Node 里跑不了。测试只 eval 模块本身，所以这里按需打桩。
globalThis.buildHtmlTable = function (headers, rows) {
  return '<table><tr>' + headers.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr>'
    + rows.map(function (r) {
        return '<tr>' + r.map(function (c) { return '<td>' + c + '</td>'; }).join('') + '</tr>';
      }).join('')
    + '</table>';
};
globalThis.formatVariableAsDate = function (d) { return String(d); };

const code = fs.readFileSync(new URL('../13 - 点检机台核对.js', import.meta.url), 'utf8');
(0, eval)(code);

test('车间推导：第2位 0/1 → TB1，2 → TB2', () => {
  assert.equal(globalThis._pc_deriveWorkshop('H2FCS506'), 'TB2');  // 取样 MachineList!D1051
  assert.equal(globalThis._pc_deriveWorkshop('V1FTA958'), 'TB1');  // 取样 MachineList!D1058
  assert.equal(globalThis._pc_deriveWorkshop('E0EN0001'), 'TB1');  // 取样 Workcenter!A9，第2位为 0
  assert.equal(globalThis._pc_deriveWorkshop('S1HS0001'), 'TB1');  // 取样 Workcenter!A2
});

test('车间推导：机台号长度 < 2 → 留空，不抛异常', () => {
  assert.equal(globalThis._pc_deriveWorkshop('H'), '');
  assert.equal(globalThis._pc_deriveWorkshop(''), '');
});

test('车间推导：意外值 → 留空（不猜）', () => {
  assert.equal(globalThis._pc_deriveWorkshop('H3FTA001'), '');
  assert.equal(globalThis._pc_deriveWorkshop('HXFTA001'), '');
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: FAIL —— `TypeError: globalThis._pc_deriveWorkshop is not a function`

- [ ] **Step 3: 在 `13 - 点检机台核对.js` 的 `_pc_buildIncludedSet` 之后追加实现**

```javascript
/**
 * 由机台号推导车间：从左数第 2 个字符
 * 实测 Workcenter 全部 326 台机台第二位只有 0/1/2；其他值一律留空交人工（不猜）
 * @param {string} machineNo
 * @returns {string} "TB1" | "TB2" | ""
 */
function _pc_deriveWorkshop(machineNo) {
  const c = String(machineNo || "").charAt(1);
  if (c === "0" || c === "1") return "TB1";
  if (c === "2") return "TB2";
  return "";
}
```

- [ ] **Step 4: 跑测试确认通过**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: PASS，3 条全绿

- [ ] **Step 5: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-sync.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.08_点检同步_新增车间推导

由机台号第2位推导：0/1→TB1，2→TB2，其他留空交人工。
意外值不猜、不抛异常。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: 同步计划计算

把"读到的两张表"算成"要追加什么、要删什么、什么只报告"。这是整个模块的决策核心。

**Files:**
- Modify: `13 - 点检机台核对.js`（追加常量与一个函数）
- Test: `tests/pointcheck-sync.test.mjs`（追加）

**Interfaces:**
- Consumes: `_pc_normalizeMachineNo`、`_pc_deriveWorkshop`（Task 1、2）；`_pc_buildIncludedSet` 的返回结构（`included` / `flagged`）
- Produces:
  - `_pc_computeSyncPlan(injRows, built) → { append: Array<{machineNo, workshop}>, toDelete: Array<{machineNo, rowIndexes: number[], reason: "FLAG"|"ABSENT", snapshot: Array<Array>}>, plasmaKept: Array<{machineNo, rowIndexes: number[]}> }`
  - `injRows` 沿用现有结构：`Array<{rowIndex: number, machineNo: string, rowData: Array}>`（`rowIndex` 为 1-based 表行号）
  - 常量 `_pc_EXEMPT_PLASMA = "Plasma"`

- [ ] **Step 1: 写失败测试（追加到 `tests/pointcheck-sync.test.mjs`）**

```javascript
// —— 取样自生产表 `Database_PointCheck-点检后台数据` → `MachineList` 的 INJ 段 ——
// 工序/车间/机型/机台号/点检人均为该表真实值；rowIndex 只是测试用行号占位，
// 不绑定具体行（该表会被人工编辑，行号在漂移：本次取样后又上移了 2 行）
const ML_HEADERS = ['工序', '车间', '机型', '机台号', '点检人'];
const mlRow = (a, b, c, d, e) => [a, b, c, d, e];

const INJ_ROWS_FIXTURE = [
  { rowIndex: 1051, machineNo: 'H2FCS506', rowData: mlRow('INJ', 'TB2', '6AX(自动化部分 for FCS)', 'H2FCS506', '技术员') },
  { rowIndex: 1052, machineNo: 'H2FCS506', rowData: mlRow('INJ', 'TB2', 'FCS', 'H2FCS506', 'OPC') },
  { rowIndex: 1053, machineNo: 'H2HTA520', rowData: mlRow('INJ', 'TB2', '6AX', 'H2HTA520', 'OPC') },
  { rowIndex: 1058, machineNo: 'V1FTA958', rowData: mlRow('INJ', 'TB1', 'OMNI', 'V1FTA958', '技术员') },
];

const builtWith = (machines, flagged) => ({
  included: machines, flagged: flagged || {}, missing: [], duplicates: [], badFlags: [],
});

test('同机台号多行：整组一起删除，snapshot 含全部行', () => {
  const plan = globalThis._pc_computeSyncPlan(
    INJ_ROWS_FIXTURE,
    builtWith({}, { H2FCS506: true })   // 501/502 被标Y，其余在 included 里
  );
  // 注意：本用例只标了 H2FCS506，故其它机会进 ABSENT 删除集，见下条断言
  const d = plan.toDelete.find(x => x.machineNo === 'H2FCS506');

  assert.ok(d, 'H2FCS506 应进删除集');
  assert.deepEqual(d.rowIndexes, [1051, 1052], '两行整组一起删');
  assert.equal(d.reason, 'FLAG');
  assert.equal(d.snapshot.length, 2, '快照要含两行，只记机台号无法恢复');
});

test('删除原因区分：表11标Y → FLAG；表11完全没有 → ABSENT', () => {
  const plan = globalThis._pc_computeSyncPlan(
    INJ_ROWS_FIXTURE,
    builtWith({ V1FTA958: {} }, { H2FCS506: true })
  );

  const reasons = plan.toDelete.map(x => [x.machineNo, x.reason]).sort();
  assert.deepEqual(reasons, [
    ['H2FCS506', 'FLAG'],     // 表11 有这行，但标了 Y
    ['H2HTA520', 'ABSENT'],   // 表11 里根本没有
  ]);
});

test('Plasma 的 INJ 行不删除，进 plasmaKept', () => {
  // 真实数据：生产表 MachineList 中唯一的 Plasma 行 —— 机型与机台号**都是** `Plasma`。
  // 机台号是裸值 `Plasma`，所以断言归一化后的 `PLASMA` —— 顺带锁住 raw → 归一 这一步
  // （若改用大写且无空格的编造号，这个断言就抓不住归一化回归）
  const rows = [
    { rowIndex: 844, machineNo: 'Plasma', rowData: mlRow('INJ', 'TB1', 'Plasma', 'Plasma', '技术员') },
  ];
  const plan = globalThis._pc_computeSyncPlan(rows, builtWith({}));

  assert.deepEqual(plan.toDelete, [], 'Plasma 不删');
  assert.deepEqual(plan.plasmaKept.map(x => x.machineNo), ['PLASMA']);
});

test('纳入集有、MachineList 无 → 进追加集，字段按规则推导', () => {
  const plan = globalThis._pc_computeSyncPlan(
    INJ_ROWS_FIXTURE,
    builtWith({ H2FCS506: {}, H2HTA520: {}, V1FTA958: {}, E0EN0001: {} })
  );

  assert.deepEqual(plan.toDelete, [], '四台都在纳入集里，没有要删的');
  assert.deepEqual(plan.append, [{ machineNo: 'E0EN0001', workshop: 'TB1' }]);
});

test('机台号为空/仅空格的 INJ 行：不参与增删，也不报错', () => {
  // 防御性用例：生产表当前**没有**机台号为空的行。这里守住的是
  // 「读表异常/列错位导致机台号读空」时，不能把这批行当成「表11查不到」而删掉
  const rows = INJ_ROWS_FIXTURE.concat([
    { rowIndex: 901, machineNo: '', rowData: mlRow('INJ', 'TB2', '', '', '技术员') },
    { rowIndex: 902, machineNo: '   ', rowData: mlRow('INJ', 'TB2', '', '   ', '技术员') },
  ]);
  const plan = globalThis._pc_computeSyncPlan(rows, builtWith({ H2FCS506: {}, H2HTA520: {}, V1FTA958: {} }));

  assert.deepEqual(plan.toDelete, [], '空机台号不能被当成"表11查不到"而删掉');
  assert.deepEqual(plan.append, []);
});

test('追加集按车间留空计入待补全（意外第2位）', () => {
  const rows = [{ rowIndex: 1058, machineNo: 'V1FTA958', rowData: mlRow('INJ', 'TB1', 'OMNI', 'V1FTA958', '技术员') }];
  const plan = globalThis._pc_computeSyncPlan(rows, builtWith({ V1FTA958: {}, H3FTA001: {} }));

  assert.deepEqual(plan.append, [{ machineNo: 'H3FTA001', workshop: '' }]);
});

test('防御：MachineList 的 INJ 段读空 → 不追加任何机台', () => {
  // 若这里返回整个纳入集，INJ 段一旦读取异常就会把几百台计划账机台灌进点检表
  const plan = globalThis._pc_computeSyncPlan([], builtWith({ H2FTA001: {}, H2FTA002: {}, H2FTA003: {} }));

  assert.deepEqual(plan, { append: [], toDelete: [], plasmaKept: [] });
});

test('防御：INJ 行非空但机台号列全读空 → 同样不追加任何机台', () => {
  // 与上一条同源、不同扳机：列错位/读表异常时 injRows 非空而机台号全空，
  // mlGroups 会是空的 —— 不守的话整个纳入集照样被当成「待追加」批量写进主数据
  const rows = [
    { rowIndex: 901, machineNo: '', rowData: mlRow('INJ', 'TB2', '', '', '技术员') },
    { rowIndex: 902, machineNo: '   ', rowData: mlRow('INJ', 'TB2', '', '   ', '技术员') },
  ];
  const plan = globalThis._pc_computeSyncPlan(rows, builtWith({ H2FTA001: {}, H2FTA002: {} }));

  assert.deepEqual(plan, { append: [], toDelete: [], plasmaKept: [] });
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: FAIL —— `TypeError: globalThis._pc_computeSyncPlan is not a function`

- [ ] **Step 3: 在 `13 - 点检机台核对.js` 的 `_pc_deriveWorkshop` 之后追加实现**

先在第 15 行附近（`_pc_FLAG_YES` 之后）追加常量：

```javascript
const _pc_EXEMPT_PLASMA = "Plasma";       // 点检侧独有、表11 完全没有的机型，不纳入删除
```

然后追加函数：

```javascript
/**
 * 算出本次同步计划（纯函数，不做任何 IO）
 * 比对键 = 机台号；一台机在 MachineList 中可能占多行，整组一起进、一起出
 * @param {Array<{rowIndex: number, machineNo: string, rowData: Array}>} injRows MachineList 中 工序=INJ 的行
 * @param {{included: Object, flagged: Object}} built _pc_buildIncludedSet 的结果
 * @returns {{append: Array, toDelete: Array, plasmaKept: Array}}
 */
function _pc_computeSyncPlan(injRows, built) {
  // 防御：MachineList 的 INJ 段读空时，绝不能把整个纳入集当成「待追加」灌进去
  if (!injRows || injRows.length === 0) return { append: [], toDelete: [], plasmaKept: [] };

  const included = built.included || {};
  const flagged = built.flagged || {};

  // 按机台号聚拢 MachineList 的行（保留首次出现顺序，供追加位置与稳定性判断）
  const mlGroups = {};
  injRows.forEach(function (r) {
    const no = _pc_normalizeMachineNo(r.machineNo);
    if (!no) return;                       // 机台号为空的行不参与比对
    if (!mlGroups[no]) mlGroups[no] = { rows: [], rowData: [] };
    mlGroups[no].rows.push(r.rowIndex);
    mlGroups[no].rowData.push(r.rowData);
  });

  // 二次防御：injRows 非空但机台号列全读空（列错位/读表异常）→ mlGroups 为空，
  // 不守的话整个纳入集照样会被当成「待追加」批量写进主数据。同一灾难、不同扳机
  if (Object.keys(mlGroups).length === 0) return { append: [], toDelete: [], plasmaKept: [] };

  const toDelete = [];
  const plasmaKept = [];

  Object.keys(mlGroups).forEach(function (no) {
    const g = mlGroups[no];
    if (Object.prototype.hasOwnProperty.call(included, no)) return;   // 两侧都有 → 不动

    // 表11 完全没有 / 已被标Y → 删；但 Plasma 机型豁免
    const isPlasma = g.rowData.some(function (d) {
      return String(d[2] || "").trim() === _pc_EXEMPT_PLASMA;
    });
    if (isPlasma) {
      plasmaKept.push({ machineNo: no, rowIndexes: g.rows });
      return;
    }

    toDelete.push({
      machineNo: no,
      rowIndexes: g.rows,
      reason: Object.prototype.hasOwnProperty.call(flagged, no) ? "FLAG" : "ABSENT",
      snapshot: g.rowData,
    });
  });

  const append = [];
  Object.keys(included).forEach(function (no) {
    if (Object.prototype.hasOwnProperty.call(mlGroups, no)) return;
    append.push({ machineNo: no, workshop: _pc_deriveWorkshop(no) });
  });

  return { append: append, toDelete: toDelete, plasmaKept: plasmaKept };
}
```

- [ ] **Step 4: 跑测试确认通过**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: PASS，11 条全绿

- [ ] **Step 5: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-sync.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.09_点检同步_新增同步计划计算

按机台号聚拢，同号多行整组增删；删除原因区分 FLAG/ABSENT；
Plasma 机型豁免不删；空机台号行不参与比对。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: 安全阀

**Files:**
- Modify: `13 - 点检机台核对.js`（追加常量与一个函数）
- Test: `tests/pointcheck-sync.test.mjs`（追加）

**Interfaces:**
- Consumes: `_pc_computeSyncPlan` 的 `toDelete` 数组（Task 3）
- Produces:
  - `_pc_shouldDelete(deleteCount: number, isManual: boolean) → { ok: boolean, reason: string }`
  - 常量 `_pc_DELETE_LIMIT = 10`

- [ ] **Step 1: 写失败测试（追加到 `tests/pointcheck-sync.test.mjs`）**

```javascript
test('安全阀：拟删 10 台 → 执行；11 台 → 拦下', () => {
  assert.equal(globalThis._pc_shouldDelete(10, false).ok, true);
  assert.equal(globalThis._pc_shouldDelete(0, false).ok, true);

  const blocked = globalThis._pc_shouldDelete(11, false);
  assert.equal(blocked.ok, false);
  assert.match(blocked.reason, /11/, '原因里要带上实际台数，方便人判断');
});

test('安全阀：手动运行绕过（人工放行）', () => {
  assert.equal(globalThis._pc_shouldDelete(50, true).ok, true);
});

test('安全阀计数单位是机台数不是行数', () => {
  // H2FCS506 占 2 行，加上 H2HTA520、V1FTA958 → 3 台机共 4 行
  const plan = globalThis._pc_computeSyncPlan(
    INJ_ROWS_FIXTURE,
    builtWith({}, {})
  );
  assert.equal(plan.toDelete.length, 3, '3 台机（不是 4 行）');
  assert.equal(
    plan.toDelete.reduce((n, d) => n + d.rowIndexes.length, 0), 4,
    '对应 4 行 —— 阀按机台数算，3 < 10 放行'
  );
  assert.equal(globalThis._pc_shouldDelete(plan.toDelete.length, false).ok, true);
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: FAIL —— `TypeError: globalThis._pc_shouldDelete is not a function`

- [ ] **Step 3: 追加实现**

在 `_pc_EXEMPT_PLASMA` 之后追加常量：

```javascript
// 安全阀：单次删除超过这个机台数就只报告不删。稳态日差异应为 0-3 台，
// 超阈值几乎必定是上游异常（表头错位、人工误标）而非真有批量变动
const _pc_DELETE_LIMIT = 10;
```

在 `_pc_computeSyncPlan` 之后追加函数：

```javascript
/**
 * 安全阀：是否执行删除。追加不设阀（可逆），只拦删除（不可逆）
 * @param {number} deleteCount 拟删除的**机台数**（不是行数）
 * @param {boolean} isManual 手动运行 = 人工放行，绕过阀
 * @returns {{ok: boolean, reason: string}}
 */
function _pc_shouldDelete(deleteCount, isManual) {
  if (isManual) return { ok: true, reason: "" };
  if (deleteCount > _pc_DELETE_LIMIT) {
    return {
      ok: false,
      reason: "拟删除 " + deleteCount + " 台，超过阈值 " + _pc_DELETE_LIMIT + "，本次仅报告未执行删除",
    };
  }
  return { ok: true, reason: "" };
}
```

- [ ] **Step 4: 跑测试确认通过**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: PASS，14 条全绿

- [ ] **Step 5: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-sync.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.10_点检同步_新增删除安全阀

拟删机台数 > 10 只报告不删（追加照常）；手动运行绕过阀。
计数单位为机台数，非行数。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: 待补全判定与标黄区间

标黄从 v1 的"差异类型1 待确认"改语义为"待补全"。关键约束：只写变色区间，非 INJ 行永不写入。

**Files:**
- Modify: `13 - 点检机台核对.js`（追加常量与两个函数）
- Test: `tests/pointcheck-sync.test.mjs`（追加）

**Interfaces:**
- Consumes: 无（纯输入）
- Produces:
  - `_pc_isIncomplete(rowData: Array) → boolean` —— `机型(C列,index2)` 或 `点检人(E列,index4)` 为空
  - `_pc_computeColorActions(bgData: Array<Array>, injRowIdx: Set<number>, incompleteIdx: Set<number>) → Array<{start: number, count: number, color: string|null}>` —— `start` 为 **0-based** 行下标，供 `getRange(start+1, ...)` 使用
  - 常量 `_pc_HIGHLIGHT = "#FFFF00"`

- [ ] **Step 1: 写失败测试（追加到 `tests/pointcheck-sync.test.mjs`）**

```javascript
test('待补全判定：机型或点检人为空', () => {
  assert.equal(globalThis._pc_isIncomplete(mlRow('INJ', 'TB2', '6AX', 'H2HTA520', 'OPC')), false);
  assert.equal(globalThis._pc_isIncomplete(mlRow('INJ', 'TB2', '', 'H2HTA520', 'OPC')), true, '机型空');
  assert.equal(globalThis._pc_isIncomplete(mlRow('INJ', 'TB2', '6AX', 'H2HTA520', '')), true, '点检人空');
  assert.equal(globalThis._pc_isIncomplete(mlRow('INJ', 'TB2', '  ', 'H2HTA520', '   ')), true, '仅空格也算空');
});

test('标黄：只输出目标色与当前色不同的区间', () => {
  // 4 行数据（0-based）：0=无需标黄且已无色，1=需标黄且无色，
  // 2=需标黄且已是黄色（不变，不该写），3=原为黄但已不需（要清除）
  const bg = [
    ['', '', '', '', ''],
    [null, null, null, null, null],
    ['#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00'],
    ['#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00'],
  ];
  const injIdx = new Set([0, 1, 2, 3]);
  const incomplete = new Set([1, 2]);

  const actions = globalThis._pc_computeColorActions(bg, injIdx, incomplete);

  assert.deepEqual(actions, [
    { start: 1, count: 1, color: '#FFFF00' },   // 0 不变、2 不变 → 不写
    { start: 3, count: 1, color: null },        // 3 要清除
  ]);
});

test('标黄：非 INJ 行即使颜色不同也不产生任何动作', () => {
  const bg = [
    ['#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00'],  // 第0行是 PK，非 INJ
    ['', '', '', '', ''],
  ];
  const injIdx = new Set([1]);
  const incomplete = new Set([1]);

  const actions = globalThis._pc_computeColorActions(bg, injIdx, incomplete);

  assert.deepEqual(actions, [{ start: 1, count: 1, color: '#FFFF00' }], '第0行不得出现');
});

test('null 背景色视为无色，需要时能正常标黄', () => {
  const bg = [[null, null, null, null, null]];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0]), new Set([0]));

  assert.deepEqual(actions, [{ start: 0, count: 1, color: '#FFFF00' }]);
});

test('标黄：连续行合并成一个区间', () => {
  const bg = [
    ['', '', '', '', ''],
    ['', '', '', '', ''],
    ['', '', '', '', ''],
  ];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0, 1, 2]), new Set([0, 1, 2]));

  assert.deepEqual(actions, [{ start: 0, count: 3, color: '#FFFF00' }], '一次 setBackgrounds 覆盖 3 行');
});

test('标黄：真实 getBackgrounds() 形态 —— 小写十六进制与 #ffffff', () => {
  // getBackgrounds() 返回**小写**十六进制；未设置/已清除的单元格返回 "#ffffff" 而非 null/空串。
  // 不做归一的话「目标色 ≠ 当前色」恒为真，每次运行都全量重写 —— 这条用例就是它的回归守卫。
  // 四行构成完整二维：完整/待补全 × 黄/白
  const bg = [
    ['#ffff00', '#ffff00', '#ffff00', '#ffff00', '#ffff00'],   // 完整行却留着黄 → 要清除
    ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],   // 完整行且白 → 已处目标态，不写
    ['#ffff00', '#ffff00', '#ffff00', '#ffff00', '#ffff00'],   // 待补全且已是黄 → 已处目标态，不写
    ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],   // 待补全且白 → 要标黄
  ];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0, 1, 2, 3]), new Set([2, 3]));

  assert.deepEqual(actions, [
    { start: 0, count: 1, color: null },            // 清除人工遗留的黄
    { start: 3, count: 1, color: '#FFFF00' },
  ], '第 1、2 行已处目标态；第 0 行要清除、第 3 行要标黄');
});

test('标黄：仅空格的当前色视为无色', () => {
  const bg = [['   ', '   ', '   ', '   ', '   ']];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0]), new Set([0]));

  assert.deepEqual(actions, [{ start: 0, count: 1, color: '#FFFF00' }]);
});

test('标黄：非 INJ 行夹在两行待补全之间 → 必须切成两个区间', () => {
  // 模块最硬的安全不变量所依赖的形状：区间一旦跨过非 INJ 行，就会把 PK/TF 行一起重涂
  // （表内 INJ 段被 TF/PK 行夹断，例如 1051-1058 那段紧跟在 TF 块之后）
  const bg = [
    ['', '', '', '', ''],
    ['', '', '', '', ''],
    ['', '', '', '', ''],
  ];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0, 2]), new Set([0, 2]));

  assert.deepEqual(actions, [
    { start: 0, count: 1, color: '#FFFF00' },
    { start: 2, count: 1, color: '#FFFF00' },
  ], '不能返回一个跨第 1 行的区间');
});

test('标黄：中间一行已处于目标态 → 不能把三段合并', () => {
  const bg = [
    ['', '', '', '', ''],
    ['#ffff00', '#ffff00', '#ffff00', '#ffff00', '#ffff00'],
    ['', '', '', '', ''],
  ];
  const actions = globalThis._pc_computeColorActions(bg, new Set([0, 1, 2]), new Set([0, 1, 2]));

  assert.deepEqual(actions, [
    { start: 0, count: 1, color: '#FFFF00' },
    { start: 2, count: 1, color: '#FFFF00' },
  ], '第 1 行不需改色，应作为断点而非被合并进去');
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: FAIL —— `TypeError: globalThis._pc_isIncomplete is not a function`

- [ ] **Step 3: 追加实现**

在 `_pc_DELETE_LIMIT` 之后追加常量：

```javascript
const _pc_HIGHLIGHT = "#FFFF00";          // 待补全行的标黄颜色（写入值，大小写皆可）

/**
 * 归一化背景色，供「目标色 ≠ 当前色」比较使用。
 * **必须归一**：`getBackgrounds()` 返回**小写**十六进制，且未设置/已清除的单元格
 * 返回 `"#ffffff"` 而不是 null/空串。不归一的话比较恒不相等 ——
 * 每次运行都会把整段 INJ 行重写一遍，「只写变色区间」的约束会失效。
 * @param {*} value 原始背景色值
 * @returns {string|null} 归一后的小写色值；无色返回 null
 */
function _pc_normalizeBg(value) {
  const s = String(value === undefined || value === null ? "" : value).trim().toLowerCase();
  if (s === "" || s === "#ffffff" || s === "white") return null;
  return s;
}
```

在 `_pc_shouldDelete` 之后追加两个函数：

```javascript
/**
 * 待补全判定：机型(C列,index2) 或 点检人(E列,index4) 为空
 * 追加行天然为空 → 必然待补全，靠标黄+邮件暴露直到人工补齐
 * @param {Array} rowData
 * @returns {boolean}
 */
function _pc_isIncomplete(rowData) {
  const model = String(rowData[2] || "").trim();
  const checker = String(rowData[4] || "").trim();
  return model === "" || checker === "";
}

/**
 * 算出需要改色的区间（纯函数）。只输出「目标色 ≠ 当前色」的连续区间，
 * 颜色没变的行一概不产生动作 —— 非 INJ 行永远不需要改色，因此永远不会被写入
 * @param {Array<Array>} bgData 全表背景色（0-based，含表头行）
 * @param {Set<number>} injRowIdx 需管理的 INJ 行下标（0-based）
 * @param {Set<number>} incompleteIdx 待补全行下标（0-based）
 * @returns {Array<{start: number, count: number, color: string|null}>} start 为 0-based
 */
function _pc_computeColorActions(bgData, injRowIdx, incompleteIdx) {
  const actions = [];
  let cur = null;

  for (let i = 0; i < bgData.length; i++) {
    // 非 INJ 行不受管理：不读它的颜色、不为它产生动作，且**必须打断正在累积的区间**
    // （表内 INJ 段被 PK/TF 行夹断，区间一旦跨过去就会把那些行一起重涂）
    if (!injRowIdx.has(i)) {
      if (cur) { actions.push(cur); cur = null; }
      continue;
    }

    const target = incompleteIdx.has(i) ? _pc_HIGHLIGHT : null;
    // 只读 A 列颜色：本模块写入时整行 A~E 涂同一色，故 A 列即代表整行。
    // 若将来出现「只涂部分列」的写入，这里会静默失效 —— 届时须改为逐列比较
    const current = _pc_normalizeBg((bgData[i] || [])[0]);

    if (_pc_normalizeBg(target) === current) {
      if (cur) { actions.push(cur); cur = null; }
      continue;
    }

    if (cur && cur.color === target) {
      cur.count++;
    } else {
      if (cur) actions.push(cur);
      cur = { start: i, count: 1, color: target };
    }
  }
  if (cur) actions.push(cur);
  return actions;
}
```

- [ ] **Step 4: 跑测试确认通过**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: PASS，23 条全绿

- [ ] **Step 5: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-sync.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.11_点检同步_待补全判定与标黄区间计算

标黄语义由「差异类型1待确认」改为「待补全」；
只输出变色区间，非 INJ 行不产生任何动作。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: 编排与邮件改造

把前面 5 个纯函数接上真实 IO，替换 v1 的差异计算与邮件。

**Files:**
- Modify: `13 - 点检机台核对.js`（改写 `checkPointCheckMachines` 主体、`_pc_updateHighlights` 删除、`_pc_buildEmailHtml` 改写；新增 `dryRunPointCheckSync`、`_pc_executePlan`、`_pc_mapByRowIndex`）
- Test: `tests/pointcheck-sync.test.mjs`（追加邮件断言）

**Interfaces:**
- Consumes: Task 1-5 全部函数
- Produces:
  - `checkPointCheckMachines(e)` —— 入口名不变；`e === undefined` 为手动（绕过安全阀）
  - `dryRunPointCheckSync()` —— 只报告不写入
  - `_pc_buildEmailHtml(result, today) → string`
  - `result` 结构：`{ append, deleted, plasmaKept, incomplete, deleteSkipped, valveReason, badFlags, duplicates, dryRun }`
  - 常量 `_pc_ST_*` 六个状态字符串

- [ ] **Step 1: 写失败测试（追加到 `tests/pointcheck-sync.test.mjs`）**

```javascript
test('邮件包含两张表的链接', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');

  assert.match(html, /1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY\/edit#gid=436306312/, '点检表 MachineList 页签');
  assert.match(html, /12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM\/edit#gid=0/, '表11 Workcenter 页签');
});

function emptyResult(over) {
  return Object.assign({
    append: [], deleted: [], plasmaKept: [], incomplete: [],
    deleteSkipped: false, valveReason: '', badFlags: [], duplicates: [], dryRun: false,
  }, over || {});
}

test('邮件状态列：已添加 / 已删除两种原因 / 待补全 / 仅报告', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult({
    append: [{ machineNo: 'E0EN0001', workshop: 'TB1' }],
    deleted: [
      { machineNo: 'H2FTA001', reason: 'FLAG', snapshot: [mlRow('INJ', 'TB2', '6AX', 'H2FTA001', 'OPC')] },
      { machineNo: 'H2FTA002', reason: 'ABSENT', snapshot: [mlRow('INJ', 'TB2', '3AX', 'H2FTA002', 'OPC')] },
    ],
    plasmaKept: [{ machineNo: 'PLASMA' }],
    incomplete: [{ machineNo: 'E0EN0001', rowData: mlRow('INJ', 'TB1', '', 'E0EN0001', '') }],
  }), '2026-09-30');

  assert.match(html, /已添加·需人工维护/);
  assert.match(html, /已删除·因表11标Y/, '真删了才说已删除');
  assert.match(html, /已删除·表11查不到/);
  assert.match(html, /待补全/);
  assert.match(html, /仅报告·Plasma豁免/);
  assert.doesNotMatch(html, /未执行/, '没被拦下就不能出现未执行');
});

test('邮件：安全阀拦下时必须列出被拦的机台，状态为未执行', () => {
  // 阀拦下的目的就是要人复核这批机台后手动放行 ——
  // 若邮件里不列出它们，人就无从复核，拦下等于白拦（且是一笔"安静的欠账"）
  const html = globalThis._pc_buildEmailHtml(emptyResult({
    deleted: [
      { machineNo: 'H2FTA001', reason: 'FLAG', snapshot: [mlRow('INJ', 'TB2', '6AX', 'H2FTA001', 'OPC')] },
      { machineNo: 'H2FTA002', reason: 'ABSENT', snapshot: [mlRow('INJ', 'TB2', '3AX', 'H2FTA002', 'OPC')] },
    ],
    deleteSkipped: true,
    valveReason: '拟删除 11 台，超过阈值 10',
  }), '2026-09-30');

  assert.match(html, /H2FTA001/, '被拦的机台必须出现在邮件里');
  assert.match(html, /H2FTA002/);
  assert.match(html, /未执行·超阈值/);
  assert.match(html, /阈值/, '须说明为什么没执行');
  assert.doesNotMatch(html, /已删除/, '一台都没删，说「已删除」就是谎报');
});

test('删除快照含整行字段（可据此恢复）', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult({
    deleted: [{ machineNo: 'H2FTA001', reason: 'FLAG', snapshot: [mlRow('INJ', 'TB2', '6AX', 'H2FTA001', 'OPC')] }],
  }), '2026-09-30');

  assert.match(html, /6AX/, '机型字段');
  assert.match(html, /OPC/, '点检人字段');
});

test('安全阀未触发时不出「未执行」字样', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');

  assert.doesNotMatch(html, /未执行/, '没拦下就不能出现未执行，否则是谎报');
});

test('数据质量告警：badFlags 非空才出该区块', () => {
  const withBad = globalThis._pc_buildEmailHtml(emptyResult({ badFlags: ['y', '是'] }), '2026-09-30');
  assert.match(withBad, /是/);
  assert.match(withBad, /数据质量|标了却|异常取值/);

  const noBad = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');
  assert.doesNotMatch(noBad, /异常取值/);
});

test('预演模式：状态列不说谎，追加/删除都标「预演未写入」', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult({
    dryRun: true,
    append: [{ machineNo: 'E0EN0001', workshop: 'TB1' }],
    deleted: [{ machineNo: 'H2FTA001', reason: 'FLAG', snapshot: [mlRow('INJ', 'TB2', '6AX', 'H2FTA001', 'OPC')] }],
  }), '2026-09-30');

  assert.match(html, /仅报告/);
  assert.match(html, /计划追加·预演未写入/);
  assert.match(html, /计划删除·预演未写入/);
  assert.doesNotMatch(html, /已删除/, '预演什么都没删，说「已删除」就是谎报');
  assert.doesNotMatch(html, /已添加/, '预演什么都没写，说「已添加」就是谎报');
});

test('两侧一致且无待补全 → 报告无差异', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');
  assert.match(html, /无差异|完全一致/);
});
```

- [ ] **Step 2: 跑测试确认失败**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: FAIL —— 邮件断言失败（现有 `_pc_buildEmailHtml` 的签名是 `(type1, type2, today)`，且无链接/状态列）

- [ ] **Step 3: 改写 `13 - 点检机台核对.js`——替换文件头注释**

把第 1-6 行的注释块替换为：

```javascript
// V20260930.12 — 点检机台核对 + 主数据双向同步
// 入口：checkPointCheckMachines（每日 08:25 定时 or 手动，手动绕过安全阀）
// 预演：dryRunPointCheckSync（只报告不写入）
// 逻辑：纳入集 = Workcenter 全部行 − (「无需检查Y/N」= Y)
//       纳入集有/MachineList-INJ 无 → 追加到表尾（工序=INJ，车间按第2位推，机型/点检人留空）
//       MachineList-INJ 有/纳入集无 → 删除该机台号全部 INJ 行（Plasma 机型豁免）
//       安全阀：拟删 > 10 台 → 只报告不删；手动运行绕过
// 安全不变量：工序 ≠ INJ 的行，值和背景色一律不写入
```

- [ ] **Step 4: 改写 `13 - 点检机台核对.js`——追加状态常量与邮件函数**

在 `_pc_HIGHLIGHT` 之后追加：

```javascript
// ========== 邮件状态字面量 ==========
const _pc_ST_ADDED = "已添加·需人工维护";
const _pc_ST_DEL_FLAG = "已删除·因表11标Y";
const _pc_ST_DEL_ABSENT = "已删除·表11查不到";
const _pc_ST_DEL_SKIPPED = "未执行·超阈值";
const _pc_ST_PLASMA = "仅报告·Plasma豁免";
const _pc_ST_INCOMPLETE = "待补全";
// 预演：什么都没写，状态列必须说实话，否则首跑核对时会被误导
const _pc_ST_ADD_PLANNED = "计划追加·预演未写入";
const _pc_ST_DEL_PLANNED = "计划删除·预演未写入";

// ========== 邮件中的两张表链接 ==========
const _pc_URL_POINTCHECK = "https://docs.google.com/spreadsheets/d/" + _pc_ID_POINTCHECK + "/edit#gid=436306312";
const _pc_URL_PLAN = "https://docs.google.com/spreadsheets/d/" + _pc_ID_PLAN + "/edit#gid=0";
```

删除整个 `_pc_buildEmailHtml`（原第 242-293 行）与 `_pc_card`（原第 295-299 行），替换为：

```javascript
/**
 * 构造同步报告邮件
 * @param {Object} result 见 Task 6 Interfaces
 * @param {string} today 已格式化的日期串
 * @returns {string} HTML
 */
function _pc_buildEmailHtml(result, today) {
  const statusOf = function (reason) {
    if (result.dryRun) return _pc_ST_DEL_PLANNED;
    // 阀拦下时这些机台**没被删**，状态列必须说实话 —— 这正是状态列存在的理由
    if (result.deleteSkipped) return _pc_ST_DEL_SKIPPED;
    return reason === "FLAG" ? _pc_ST_DEL_FLAG : _pc_ST_DEL_ABSENT;
  };

  var html = '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>';
  html += '<div style="font-family:Arial,\'Microsoft YaHei\',\'Helvetica Neue\',sans-serif;max-width:960px;margin:0 auto">';

  html += '<div style="background:#E60012;color:white;padding:16px 24px">';
  html += '<h2 style="margin:0">注塑机台同步报告</h2>';
  html += '<p style="margin:8px 0 0;opacity:0.95;font-size:14px">比对范围：MachineList(工序=INJ) ↔ Workcenter(排除「无需检查Y/N」=Y)</p>';
  html += '<p style="margin:4px 0 0;opacity:0.95;font-size:14px">';
  html += '<a href="' + _pc_URL_POINTCHECK + '" style="color:#fff">点检机台主数据</a>';
  html += ' &nbsp;|&nbsp; ';
  html += '<a href="' + _pc_URL_PLAN + '" style="color:#fff">注塑计划机台</a>';
  html += '</p>';
  html += '<p style="margin:4px 0 0;opacity:0.7;font-size:12px">发送时间：' + today + '</p>';
  html += '</div>';

  if (result.dryRun) {
    html += '<div style="background:#fff3cd;border-left:4px solid #e67e22;padding:12px 24px;color:#7a5b00">';
    html += '<strong>预演模式：仅报告，未写入任何数据。</strong></div>';
  }

  html += '<div style="padding:24px">';

  html += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px"><tr>';
  html += _pc_card(result.dryRun ? "计划追加" : "本次追加", result.append.length, "#27ae60");
  // 预演与「超阈值未执行」两种情况下都没有真删，标签与数值都说计划
  html += _pc_card(result.dryRun || result.deleteSkipped ? "计划删除" : "本次删除", result.deleted.length, "#e74c3c");
  html += _pc_card("待补全", result.incomplete.length, "#e67e22");
  html += '</tr></table>';

  const nothing = result.append.length === 0 && result.deleted.length === 0
    && result.plasmaKept.length === 0 && result.incomplete.length === 0;

  if (result.deleteSkipped) {
    html += '<p style="color:#e74c3c;font-weight:bold">⚠ ' + result.valveReason + '</p>';
  }
  if (nothing && !result.deleteSkipped) {
    html += '<p style="color:#27ae60;font-weight:bold;font-size:16px">★ 点检机台与计划账完全一致，无差异。</p>';
  }

  // ===== 本次追加 =====
  if (result.append.length > 0) {
    html += _pc_section("本次追加：" + result.append.length + " 台");
    html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">⚠ 新机台的「机型」「点检人」为空白，请人工补齐（补齐前该行在点检表中标黄）</p>';
    html += buildHtmlTable(
      ["机台号", "车间", "状态"],
      result.append.map(function (r) {
        return [r.machineNo, r.workshop || "(待补全)", result.dryRun ? _pc_ST_ADD_PLANNED : _pc_ST_ADDED];
      }),
      "#27ae60");
  }

  // ===== 本次删除 =====
  if (result.deleted.length > 0 || result.deleteSkipped) {
    const n = result.deleted.length;
    const title = result.dryRun ? "计划删除：" + n + " 台（预演未写入）"
      : result.deleteSkipped ? "计划删除：" + n + " 台（超阈值，未执行）"
      : "本次删除：" + n + " 台";
    html += _pc_section(title);
    if (result.dryRun) {
      html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">以下机台<b>尚未删除</b>。确认无误后手动运行 checkPointCheckMachines 执行</p>';
    } else if (result.deleteSkipped) {
      html += '<p style="color:#e74c3c;font-weight:bold;margin-bottom:8px">以下机台本次<b>未删除</b>，待人工确认后手动运行放行</p>';
    } else {
      html += '<p style="color:#e74c3c;font-weight:bold;margin-bottom:8px">已从点检表移除；如需恢复，请照下表整行字段补回</p>';
    }
    // 快照展开：一台机可能占多行，逐行输出便于恢复
    const rows = [];
    result.deleted.forEach(function (d) {
      const st = statusOf(d.reason);
      d.snapshot.forEach(function (s) {
        rows.push([s[0] || "-", s[1] || "-", s[2] || "-", s[3] || "-", s[4] || "-", st]);
      });
    });
    html += buildHtmlTable(["工序", "车间", "机型", "机台号", "点检人", "状态"], rows, "#E60012");
  }

  // ===== 待补全 =====
  if (result.incomplete.length > 0) {
    html += _pc_section("待补全：" + result.incomplete.length + " 行");
    html += '<p style="color:#e67e22;font-weight:bold;margin-bottom:8px">⚠ 「机型」或「点检人」为空，点检任务可能无法正常派发，请补齐</p>';
    html += buildHtmlTable(
      ["机台号", "车间", "状态"],
      result.incomplete.map(function (r) { return [r.machineNo, (r.rowData[1] || "-"), _pc_ST_INCOMPLETE]; }),
      "#e67e22");
  }

  // ===== 仅报告 =====
  if (result.plasmaKept.length > 0) {
    html += _pc_section("仅报告：" + result.plasmaKept.length + " 台");
    html += '<p style="color:#7f8c8d;margin-bottom:8px">这些机台在表11 中查不到，但机型为 Plasma，按规则不删除</p>';
    html += buildHtmlTable(
      ["机台号", "状态"],
      result.plasmaKept.map(function (r) { return [r.machineNo, _pc_ST_PLASMA]; }),
      "#E60012");
  }

  // ===== 数据质量告警 =====
  if (result.badFlags.length > 0 || result.duplicates.length > 0) {
    html += _pc_section("数据质量告警");
    if (result.badFlags.length > 0) {
      html += '<p style="color:#e74c3c">「无需检查Y/N」出现异常取值（<b>不会被当作 Y 排除</b>）：'
        + result.badFlags.join("、") + '。请改为 Y 或清空。</p>';
    }
    if (result.duplicates.length > 0) {
      html += '<p style="color:#e74c3c">Workcenter 出现重复机台号（已取第一条）：'
        + result.duplicates.join("、") + '</p>';
    }
  }

  html += '<p style="color:#bdc3c7;font-size:11px;margin-top:32px">此邮件由 PointCheck Alert 系统自动发送</p>';
  html += '</div></div></body></html>';
  return html;
}

function _pc_section(title) {
  return '<h3 style="color:#E60012;border-left:4px solid #E60012;padding-left:8px;margin-top:32px">' + title + '</h3>';
}

function _pc_card(label, value, color) {
  return '<td style="text-align:center;padding:16px;border:1px solid #ecf0f1;width:33%">' +
    '<div style="font-size:32px;font-weight:bold;color:' + color + '">' + value + '</div>' +
    '<div style="color:#7f8c8d;font-size:13px;margin-top:6px;line-height:1.5">' + label + '</div></td>';
}
```

- [ ] **Step 5: 改写 `13 - 点检机台核对.js`——替换主编排**

删除原 `checkPointCheckMachines`（第 63-179 行）与 `_pc_updateHighlights`（第 181-202 行），替换为：

```javascript
/**
 * 主入口：核对 + 同步
 * @param {Object} e 触发器事件对象。undefined（手动）= 绕过安全阀
 */
function checkPointCheckMachines(e) {
  _pc_run(e !== undefined && e !== null, null);
}

/** 预演入口：只报告不写入。用于上线首次核对将要追加/删除的清单 */
function dryRunPointCheckSync() {
  _pc_run(false, { dryRun: true });
}

/**
 * 编排：读表 → 算计划 → 过阀 → 写入 → 标黄 → 发信 → 记日志
 * @param {boolean} isScheduled true=定时（启用安全阀）
 * @param {Object|null} opts {dryRun: true} 时不执行任何写入
 */
function _pc_run(isScheduled, opts) {
  const dryRun = !!(opts && opts.dryRun);
  const trigger = dryRun ? "预演" : (isScheduled ? "定时" : "手动");

  try {
    console.log("开始执行点检机台核对...");

    // 1. 读 MachineList，过滤工序=INJ
    const ssPC = SpreadsheetApp.openById(_pc_ID_POINTCHECK);
    const wsML = ssPC.getSheetByName(_pc_SHEET_MACHINELIST);
    const dataML = wsML.getDataRange().getValues();

    if (dataML.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 为空或只有表头", trigger, "");
      return;
    }

    const injRows = [];
    for (let i = 1; i < dataML.length; i++) {
      if (String(dataML[i][0] || "").trim() === _pc_FILTER_PROCESS) {
        injRows.push({ rowIndex: i + 1, machineNo: String(dataML[i][3] || "").trim(), rowData: dataML[i] });
      }
    }
    console.log("MachineList INJ 工序行数: " + injRows.length);
    if (injRows.length === 0) {
      writeLog("checkPointCheckMachines", "跳过", "MachineList 无 INJ 工序数据", trigger, "");
      return;
    }

    // 2. 读 Workcenter，建纳入集
    const ssPlan = SpreadsheetApp.openById(_pc_ID_PLAN);
    const wsWC = ssPlan.getSheetByName(_pc_SHEET_WORKCENTER);
    const dataWC = wsWC.getDataRange().getValues();
    if (dataWC.length <= 1) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 为空", trigger, "");
      return;
    }

    const built = _pc_buildIncludedSet(dataWC);
    if (built.missing.length > 0) {
      writeLog("checkPointCheckMachines", "跳过", "Workcenter 表头缺少字段: " + built.missing.join(", "), trigger, "");
      return;
    }
    console.log("Workcenter 纳入集机台数: " + Object.keys(built.included).length);

    // 3. 算计划
    const plan = _pc_computeSyncPlan(injRows, built);
    // 4. 过安全阀（只作用于删除）
    const valve = _pc_shouldDelete(plan.toDelete.length, !isScheduled);

    console.log("追加 " + plan.append.length + " 台 / 删除 " + plan.toDelete.length + " 台 / 仅报告 " + plan.plasmaKept.length + " 台");
    if (!valve.ok) console.warn("安全阀: " + valve.reason);

    // 5-8. 写入与标黄
    // 待补全必须取**写入后**的状态：当天新追加的行机型/点检人为空，
    // 若沿用写入前的 injRows，这些行要到次日才进「待补全」区块，
    // 而那正是"安静的欠账"——这个机制存在的全部意义就是让它们当天可见
    let incomplete = [];
    if (dryRun) {
      incomplete = _pc_collectIncomplete(injRows).concat(plan.append.map(function (a) {
        return { machineNo: a.machineNo, rowData: [_pc_FILTER_PROCESS, a.workshop, "", a.machineNo, ""] };
      }));
    } else {
      incomplete = _pc_executePlan(wsML, dataML, plan, valve, injRows).incomplete;
    }

    // 9. 邮件
    const result = {
      append: plan.append,
      // 「删除」区块列的是**计划删除**的机台，不是实际删除的机台。
      // 安全阀拦下时若传空数组，被拦的这批机台就整个从邮件里消失了 ——
      // 而阀拦下恰恰是要人复核这批机台后手动放行，看不见就拦得毫无意义。
      // 「已删除」还是「未执行」由状态列区分（见 _pc_buildEmailHtml 的 statusOf）
      deleted: plan.toDelete,
      plasmaKept: plan.plasmaKept,
      incomplete: incomplete,
      deleteSkipped: !valve.ok,
      valveReason: valve.reason,
      badFlags: built.badFlags,
      duplicates: built.duplicates,
      dryRun: dryRun,
    };

    const recipients = _pc_getRecipients();
    if (recipients.length === 0) {
      writeLog("checkPointCheckMachines", "跳过", "无匹配收件人(O=INJ,P=S&C)", trigger, "");
      console.warn("未找到匹配收件人");
      return;
    }

    const today = formatVariableAsDate(new Date());
    const subject = "【点检同步】 注塑机台差异报告 " + today;
    const html = _pc_buildEmailHtml(result, today);

    // 删除快照写进 Log，邮件会过期、快照是唯一的恢复依据
    const snapshotJson = JSON.stringify(result.deleted.map(function (d) {
      return { machineNo: d.machineNo, reason: d.reason, rows: d.snapshot };
    }));

    try {
      _pc_sendMail(recipients.join(","), subject, html);
      const summary = "追加=" + plan.append.length + "台, 删除="
        + (valve.ok ? plan.toDelete.length + "台" : "0台(计划" + plan.toDelete.length + "台被安全阀拦下)")
        + ", 待补全=" + incomplete.length + "行, TO=" + recipients.length + "人";
      writeLog("checkPointCheckMachines", "成功", summary, trigger, snapshotJson);
      console.log("邮件发送成功: " + summary);
    } catch (err) {
      writeLog("checkPointCheckMachines", "失败", err.message, trigger, snapshotJson);
      console.error("发送失败: " + err.message);
    }

    console.log("点检机台核对执行完毕");

  } catch (err) {
    console.error(err.stack || err.message);
    try { writeLog("checkPointCheckMachines", "失败", err.message, trigger, err.stack || ""); } catch (e2) { }
  }
}

/**
 * 执行写入：删除 → 追加 → 重新读背景 → 标黄
 * 安全不变量：工序 ≠ INJ 的行，值和背景色一律不写入
 * @returns {{incomplete: Array}} 写入后仍待补全的行（含刚刚追加的，它们天然机型/点检人为空）
 */
function _pc_executePlan(wsML, dataML, plan, valve, injRows) {
  const deleted = valve.ok ? plan.toDelete : [];

  // ---- 删除：按连续行号区间批量 deleteRows（只含 INJ 行）----
  if (deleted.length > 0) {
    const targets = [];
    deleted.forEach(function (d) { d.rowIndexes.forEach(function (r) { targets.push(r); }); });
    targets.sort(function (a, b) { return b - a; });   // 从大到小删，避免行号移位

    let runStart = targets[0];
    let runCount = 1;
    for (let i = 1; i <= targets.length; i++) {
      if (i < targets.length && targets[i] === runStart - runCount) {
        runCount++;
      } else {
        wsML.deleteRows(runStart - runCount + 1, runCount);
        if (i < targets.length) { runStart = targets[i]; runCount = 1; }
      }
    }
    console.log("已删除 " + deleted.length + " 台机台，共 " + targets.length + " 行");
  }

  // ---- 追加：一律到表尾，不插任何现有行中间 ----
  if (plan.append.length > 0) {
    const lastRow = wsML.getLastRow();
    const rows = plan.append.map(function (a) {
      return [_pc_FILTER_PROCESS, a.workshop, "", a.machineNo, ""];
    });
    wsML.insertRowsAfter(lastRow, rows.length);
    wsML.getRange(lastRow + 1, 1, rows.length, 5).setValues(rows);
    console.log("已在表尾追加 " + rows.length + " 行");
  }

  // ---- 标黄：行号已因删除/追加而变，必须重新读背景 ----
  const lastRow = wsML.getLastRow();
  if (lastRow <= 1) return { incomplete: [] };

  const bgData = wsML.getRange(1, 1, lastRow, 5).getBackgrounds();
  const freshML = wsML.getRange(1, 1, lastRow, 5).getValues();

  const injIdx = new Set();
  const incompleteIdx = new Set();
  const incompleteRows = [];
  for (let i = 1; i < freshML.length; i++) {
    if (String(freshML[i][0] || "").trim() !== _pc_FILTER_PROCESS) continue;
    injIdx.add(i);
    if (_pc_isIncomplete(freshML[i])) {
      incompleteIdx.add(i);
      incompleteRows.push({ machineNo: String(freshML[i][3] || "").trim(), rowData: freshML[i] });
    }
  }

  _pc_computeColorActions(bgData, injIdx, incompleteIdx).forEach(function (a) {
    wsML.getRange(a.start + 1, 1, a.count, 5)
      .setBackgrounds(Array(a.count).fill(Array(5).fill(a.color)));
  });
  console.log("标黄完成：待补全 " + incompleteIdx.size + " 行");

  return { incomplete: incompleteRows };
}

/** 收集当前 INJ 中的待补全行，供邮件列出 */
function _pc_collectIncomplete(injRows) {
  return injRows.filter(function (r) { return _pc_isIncomplete(r.rowData); })
    .map(function (r) { return { machineNo: r.machineNo, rowData: r.rowData }; });
}
```

注意：`_pc_executePlan` 里删除完成后 `dataML` 已失效，所以标黄一律重新读表——这正是 spec 第 6 节要求的顺序。

- [ ] **Step 6: 跑测试确认通过**

Run: `node --test tests/pointcheck-sync.test.mjs`
Expected: PASS，31 条全绿

- [ ] **Step 7: 跑全量测试确认无回归**

Run: `node --test tests/*.test.mjs`
Expected: 全部测试文件全绿（`weekly-inspection`、`machine-cycle-workcenter`、`workcenter-sync` 不受影响）

注意用 glob，**不要写 `node --test tests/`** —— 本机 Node v22.12.0 下后者报 `MODULE_NOT_FOUND`。

- [ ] **Step 8: 语法自检（GAS 无本地运行环境，先确认能解析）**

Run: `node --check "13 - 点检机台核对.js"`
Expected: 无输出（解析通过）

- [ ] **Step 9: 提交**

```bash
git add "13 - 点检机台核对.js" tests/pointcheck-sync.test.mjs
git commit -m "$(cat <<'EOF'
V20260930.12_点检同步_编排与邮件改造

接上五个纯函数；删除按区间批量、追加到表尾、标黄前重读背景；
邮件加两张表链接、四区块与状态列、删除整行快照、数据质量告警；
新增 dryRunPointCheckSync 预演入口。

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 7: 配置注册与项目记录

**Files:**
- Modify: `docs/项目记录.md`（顶部追加一节）
- 配置（通过 MCP 改 Google Sheet，不改代码）：`菜单设置` 新增一行；`定时设置` 无需改动

**Interfaces:**
- Consumes: `dryRunPointCheckSync`（Task 6）
- Produces: 无代码接口

**背景：** 配置表 ID `1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM`，`菜单设置` 从第 3 行起、A=菜单名 B=函数名 C=是否启用(是)；`定时设置` 从第 3 行起、7 列。`checkPointCheckMachines` 的定时与菜单配置**沿用不变**。

- [ ] **Step 1: 确认配置表当前内容**

用 MCP 读 `菜单设置!A1:C{lastRow}` 与 `定时设置!A1:G{lastRow}`，确认：
- `菜单设置` 中已有 `点检机台核对` → `checkPointCheckMachines`（不要重复添加）
- `菜单设置` 中**没有** `点检同步预演`
- `定时设置` 中 `checkPointCheckMachines` 的 7 列值（不要改动）

- [ ] **Step 2: 新增菜单项**

在 `菜单设置` 末尾追加一行：

| A 菜单名 | B 函数名 | C 是否启用 |
|---|---|---|
| `点检同步预演` | `dryRunPointCheckSync` | `是` |

- [ ] **Step 3: 验证菜单项写入正确**

用 MCP 重读 `菜单设置` 最后 3 行，确认新增行三列值与 Step 2 一致，且 `checkPointCheckMachines` 那一行**未被改动**。

- [ ] **Step 4: 追加项目记录**

在 `docs/项目记录.md` 的标题行 `# SmartMeeting_Alert 项目记录` 之后、当前最新版本节之前，插入：

```markdown
## V20260930.13 — 点检机台核对升级为主数据双向同步（13号模块）

**改动内容**：
- 起因：模块 13 原为"比对 + 邮件 + 差异类型1标黄"，邮件里「需要更新点检机台主数据」这一步一直是人工的，无机制保证跟进
- **判定依据收敛为单一开关**：纳入集 = `Workcenter` 全部行 − 「无需检查Y/N」= Y。废弃 v1 的 `机器性能` 关键字匹配「闲置」（该匹配漏掉 23 台：`报废` 15 台、`机台号报废,原XXX机器闲置` 6 台、`机器闲置，机台号报废` 1 台、`新品设备改造,暂时闲置` 1 台）
- **注意**：「无需检查Y/N」**不是人工列**，是模块 12 每日 08:20 的派生输出（规则：`(Final Machine Type=6AX 且 是否主设备=N) 或 (Final Machine Type=NA)` → Y，其余写空）。因此人工标 Y 会被次日擦除；本次变更把该列采纳为唯一判据，模块 12 的算法即排除规则的唯一来源
- **新增双向同步**：纳入集有/MachineList-INJ 无 → 追加到表尾（`工序=INJ`，`车间`按机台号第2位推：0/1→TB1、2→TB2，`机型`/`点检人`留空待人工）；MachineList-INJ 有/纳入集无 → 删除该机台号全部 INJ 行
- **保留唯一豁免**：`机型 = Plasma` 的 INJ 行不删（点检侧独有、表11 完全没有，无法用 Y 表达）
- **安全不变量**：只有 `工序 = INJ` 的行可被写入。删除按 (工序=INJ, 机台号) 匹配（表内 `T1GR0301` 在 TF 下有两行，只按机台号会跨工序误删）；追加一律到表尾；标黄只写变色区间
- **安全阀**：拟删机台数 > 10 → 只报告不删（追加照常）；手动运行绕过阀。不会死锁，积压可由人工放行
- **待补全机制**：追加行的 `机型`/`点检人` 为空 → 表内标黄 + 邮件「待补全」区块，直到人工补齐（标黄语义由 v1 的"差异类型1待确认"改为"待补全"）
- **邮件改版**：加【点检】【表11】两个超链接；四区块（追加/删除/待补全/仅报告）；状态列（`已添加·需人工维护`／`已删除·因表11标Y`／`已删除·表11查不到`／`未执行·超阈值`／`仅报告·Plasma豁免`／`待补全`）；删除行输出 A~E 整行快照（并在 Log 存 JSON，邮件过期后是唯一恢复依据）
- **新增入口**：`dryRunPointCheckSync`（预演，只报告不写入），已注册到 `菜单设置`
- **入口函数名不变**（`checkPointCheckMachines`），`定时设置`/`菜单设置` 原有配置无需改动
- 新增 `tests/pointcheck-sync.test.mjs` 共 31 条测试；改写 `tests/pointcheck-workcenter.test.mjs`（8 条，原「机器性能=闲置 排除」断言随语义变更）
- 需求文档：`docs/需求文档_13.md`
- 提交区间：`V20260930.07` ~ `V20260930.13`（本节描述的是这一串提交合起来的整体改动，非单一提交）
```

- [ ] **Step 5: 提交**

```bash
git add docs/项目记录.md
git commit -m "$(cat <<'EOF'
V20260930.12_点检同步_补记项目记录与菜单配置

Co-Authored-By: Claude Code <noreply@anthropic.com>
EOF
)"
```

---

## Task 8: 真实表验证

**这是本计划的交付门槛。单测通过不等于在真表上正确。** 本轮验证必须拿到跑前跑后的对比数据才能回报完成。

**Files:**
- 不改代码。验证产物贴进对话回报。

**Interfaces:**
- Consumes: Task 1-7 全部产出
- Produces: 验证结论（含前后对比数据）

- [ ] **Step 1: 记录跑前基线**

用 MCP 读并记下：

1. `MachineList!A1:A{lastRow}` 的工序分布 → 统计各工序行数（尤其 **PK 行数、TF 行数、总行数**）
2. `MachineList!A1:E{lastRow}` 中 `工序=INJ` 的行数
3. `MachineList` 的行数（`getLastRow()` 等价值）

- [ ] **Step 2: 推送代码到 GAS**

Run: `NODE_OPTIONS="--require $HOME/.claude/clasp-proxy.cjs" npx @google/clasp push -f`
Expected: 推送成功

（本机 DNS 是 Clash fake-IP，clasp 自带的 Node 22.12 不读 `HTTPS_PROXY`，必须挂这个 bootstrap，否则报 `Client network socket disconnected before secure TLS connection was established`。bootstrap 用 clasp 自带的 `https-proxy-agent` 挂到 `https.globalAgent`，代理为 `127.0.0.1:7897`。）

- [ ] **Step 3: 跑预演，核对清单**

在 GAS 里手动执行 `dryRunPointCheckSync()`（或从专项菜单点「点检同步预演」），核对邮件与 Log：

- 「本次追加」清单里的机台号，是否确实应该做注塑点检
- 「本次删除」清单里的机台号，分 FLAG / ABSENT 两类的条数是否合理
- 「仅报告·Plasma豁免」清单
- 「数据质量告警」是否出现非空非 Y 的取值（**这是发现"标了却没生效"的地方**）
- 确认**表里没有任何变化**（预演不写入）

**若追加清单里出现不该点检的机台** → 停下来回报（spec 第 16 节「排除能力受限」：这类机台无法用手工标 Y 排除，需要先决定怎么处理再继续）。

- [ ] **Step 4: 记录预演后的基线（确认预演未写入）**

重复 Step 1 的三项统计，与基线**逐项相同**。不同即为预演写入了数据 —— 停止并回报。

- [ ] **Step 5: 确认无误后执行真实同步**

手动运行 `checkPointCheckMachines()`（手动运行绕过安全阀）。

- [ ] **Step 6: 核对写入结果**

用 MCP 重读并逐项核对：

1. **PK / TF 行数是否与基线完全一致**（这是安全不变量的核心断言，必须逐工序对比并给出数字）
2. 总行数变化 = 追加行数 − 删除行数
3. 追加的每一行，5 个字段是否为 `INJ` / `TB1|TB2` / 空 / 机台号 / 空
4. 被删的机台号是否确实全部消失，且**只删了 INJ 行**（抽查 `T1GR0301` 是否还在 TF 段）
5. 待补全行的 A~E 背景色是否为 `#FFFF00`
6. **非 INJ 行的背景色是否与基线一致**（抽查 3 行 PK、3 行 TF）
7. 邮件：两个链接可点、四个区块、状态列取值正确、删除区能看到整行快照
8. Log：`checkPointCheckMachines` 一行，状态/摘要/删除快照 JSON 齐全

- [ ] **Step 7: 回报验证结果**

把以下内容贴出来（不要只说"验证通过"）：

- 跑前 / 跑后的各工序行数对比表
- 实际追加的机台号清单 + 实际删除的机台号清单（含 reason）
- 邮件截图或关键区块的文字摘要
- Log 那一行的内容
- 任何不符预期的地方

- [ ] **Step 8: 放开定时后确认**

确认 08:25 定时配置仍指向 `checkPointCheckMachines`（Task 7 Step 1 已核对过，这里再确认一次未被误改），然后观察次日 08:25 自动运行的 Log 一行。

---

## 已知不在本计划范围

以下来自 spec 第 16 节，**本次不做**，实施时不要顺手改：

- 点检后台（Web 端）是否会回写 MachineList —— 写入独占性未验证，若确认会冲突需另开任务
- `MachineList（New）` 的迁移
- 为"无法用手工 Y 排除的特定机台"新增人工开关列
- 收件人范围的调整（维持 INJ + S&C）
