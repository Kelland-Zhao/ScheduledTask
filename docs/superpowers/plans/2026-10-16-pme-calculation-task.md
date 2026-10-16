# PME Calculation Task Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 `ScheduledTask` 中新增基于 PM Data 和 AEM SourceData 的 PME 计算任务，更新现有 `Master Data`，并汇总写入 `KPI Summary`。

**Architecture:** 新增独立的 Apps Script 模块 `35 - PME计算.js`。纯计算函数负责计划匹配、班次聚合、IM Before 轮空、指标计算和 KPI 汇总；Spreadsheet I/O 函数负责读取来源及幂等更新目标表。现有 `07 - PM_KPI写入.js` 保持不变，仅继续消费已计算的 KPI 结果。

**Tech Stack:** Google Apps Script V8、`SpreadsheetApp`、现有 `ScheduledTask` 调度器、Node.js 内置 `node:test` 测试运行器。

**Spec:** `docs/superpowers/specs/2026-10-16-pme-calculation-migration-design.md`

## Global Constraints

- 正式保养必须来自 PM Data：`类别` 包含 `PM` 且 `状态` 不等于 `取消`。
- PM 计划中的 `INJ` 映射为 `IM`；`TF` 和 `PK` 保留原工序。
- AEM 班次固定为 `1=夜班、2=早班、3=中班`，排序使用日期加班次顺序。
- `Prod_Hours` 映射为 Production Available Time；UPDT 使用 `1 - NetRuntime / Prod_Hours`。
- IM 仅在 BEFORE PM 应用最后生产班次及多次停机边界的整班轮空规则。
- TF 和 PK 不考虑 CO；PKS/PKF 及 S1DP0001/2/3/4 不进入最终 KPI。
- AFTER PM 从完整保养结束后的第一个完整班次开始，累计到首次 `Prod_Hours >= 144`。
- 保养归属月份按 PM 计划开始日期的 20 日切点计算。
- 计算使用 AEM `SourceData`，不使用 `Master Data_Excel`。
- 输出明细写入现有 `Master Data`，汇总写入现有 `KPI Summary`。
- 不修改旧 `PM_KPI` 项目和 `ScheduledTask/07 - PM_KPI写入.js`。
- 不在代码中保存密码、Token 或服务账号凭据。

## Review Focus

- 同一班次有多条 AEM 分段时，IM 的停机边界判断必须基于聚合后的整班，而不是某一条分段；测试归入 Task 2。
- PM 计划开始时间为空、保养跨越班次 3 到次日班次 1 时，保养结束班次不能偏移；测试归入 Task 2。
- AEM SourceData 季度来源有重叠或缺失时，任务必须在写入前中止；测试归入 Task 4。
- `Master Data` 有历史记录、重复序列号或空指标时，更新必须幂等且不能把空值变成 0；测试归入 Task 4。
- KPI Summary 的 TF 子工序、PK 行和 YTD 区域必须与既有布局一致；测试归入 Task 5。

---

### Task 1: PME 纯计算基础与 PM 计划标准化

**Files:**
- Create: `35 - PME计算.js`
- Create: `tests/pme-calculation.test.mjs`

**Interfaces:**
- Produces `_pme_normalizeWorkcenter(value) -> string`。
- Produces `_pme_headerIndex(headers) -> Object<string, number>`。
- Produces `_pme_mapProcess(value) -> string`。
- Produces `_pme_selectPmPlans(rows, headerIndex) -> Array<PmePlan>`，只返回符合 PM/非取消条件的计划。
- Produces `_pme_compareShift(a, b) -> number`，按日期和 1/2/3 班次排序。

- [ ] **Step 1: Write the failing tests**

  在 `tests/pme-calculation.test.mjs` 中用 `fs.readFileSync` + `eval` 加载模块，先写以下测试：

  - `_pme_normalizeWorkcenter('T2GB0105【SW】')` 返回 `T2GB0105`；
  - `_pme_mapProcess('INJ')` 返回 `IM`，`TF` 和 `PK` 保持原值；
  - `_pme_selectPmPlans` 只保留 `类别` 包含 `PM` 且 `状态` 不为 `取消` 的行；
  - 表头列顺序改变时，计划字段仍按表头名读取；
  - `_pme_compareShift` 将同日班次按 1、2、3 排序，且班次 3 后接次日班次 1。

- [ ] **Step 2: Run tests to verify they fail**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: FAIL，失败原因为新 PME 函数尚未定义。

- [ ] **Step 3: Implement the pure parsing helpers**

  在 `35 - PME计算.js` 中实现上述函数。`_pme_selectPmPlans` 返回统一结构，至少包含 `serialNo`、`keyProcess`、`workcenter`、`pmDate`、`startTime`、`durationHours` 和计划原始字段；机台号先标准化，`serialNo` 使用标准化机台号。

- [ ] **Step 4: Run tests to verify they pass**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: Task 1 tests PASS。

- [ ] **Step 5: Commit**

  ```bash
  git add '35 - PME计算.js' tests/pme-calculation.test.mjs
  git commit -m "feat: add PME plan and shift primitives"
  ```

### Task 2: 班次聚合、PM 窗口与 IM Before 轮空

**Files:**
- Modify: `35 - PME计算.js`
- Modify: `tests/pme-calculation.test.mjs`

**Interfaces:**
- Consumes: Task 1 的 `PmePlan` 和 `_pme_compareShift`。
- Produces `_pme_aggregateAemShifts(rows, headerIndex) -> Array<PmeShift>`。
- Produces `_pme_getMaintenanceWindow(plan) -> {startShiftKey, endShiftKey, afterStartShiftKey}`。
- Produces `_pme_selectBeforeShifts(process, shifts, pmStartShiftKey) -> Array<PmeShift>`。
- Produces `_pme_selectAfterWindow(process, shifts, afterStartShiftKey) -> {rows, status, totalProdHours}`。

- [ ] **Step 1: Write the failing tests**

  增加以下测试：

  - 同一机台、日期、班次的多条 AEM 分段合并为一条班次记录，并分别累加 `Prod_Hours`、`NetRuntime` 和 `PlantAvailableTime`；
  - 计划开始时间为空时从当天早班开始；
  - 4 个班次的保养结束后，After 从第 5 个班次开始；
  - IM 连续生产时，保养前最后一个 `Prod_Hours > 0` 班次不进入 Before；
  - IM 出现“生产班次 → Prod_Hours=0 空班”时，生产班次被轮空，空班自身不增加累计值；
  - IM 多次开停时，每个停机边界前的生产班次都被轮空；
  - TF 和 PK 的 Before 不应用 IM 轮空；
  - After 累计到首次大于等于 144 小时，且最后一班允许超过 144；
  - 班次 3 后的下一个班次正确为次日班次 1。

- [ ] **Step 2: Run tests to verify they fail**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: 新增窗口测试 FAIL，提示窗口和 IM 轮空函数尚未定义或结果不符合预期。

- [ ] **Step 3: Implement shift aggregation and windows**

  先聚合完整班次，再执行 IM 的轮空判断。`_pme_selectBeforeShifts` 只返回实际参与 Before 指标计算的班次；`_pme_selectAfterWindow` 只返回从完整保养结束后开始、直到累计达到 144 小时的班次。TF/PK 直接使用正常窗口，不读取 CO 做筛选。

- [ ] **Step 4: Run tests to verify they pass**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: Task 1 和 Task 2 tests PASS。

- [ ] **Step 5: Commit**

  ```bash
  git add '35 - PME计算.js' tests/pme-calculation.test.mjs
  git commit -m "feat: implement PME maintenance windows"
  ```

### Task 3: AU、UPDT、状态、排除和 Master Data 行模型

**Files:**
- Modify: `35 - PME计算.js`
- Modify: `tests/pme-calculation.test.mjs`

**Interfaces:**
- Consumes: Task 2 的 Before/After 班次窗口。
- Produces `_pme_calculateMetrics(beforeShifts, afterShifts) -> PmeMetrics`。
- Produces `_pme_getYearMonth(pmDate) -> string`。
- Produces `_pme_isKpiExcluded(plan, exclusionRules) -> boolean`。
- Produces `_pme_buildMasterRow(plan, metrics, subprocess, exclusionInfo) -> Array`，输出既有 `Master Data` 的 15 字段顺序。
- Produces `_pme_calculateKpiEligibility(row) -> boolean`。

- [ ] **Step 1: Write the failing tests**

  增加以下测试：

  - AU 使用 `NetRuntime / PlantAvailableTime`；
  - UPDT 使用 `1 - NetRuntime / Prod_Hours`，不使用 `1 - AU`；
  - Gap 等于 After UPDT 减 Before UPDT；
  - After 达到 144 且 After UPDT 不高于 Before UPDT 时为 `好/Good`；
  - After 未达 144 或 Before UPDT 缺失时，绩效状态为空；
  - 日期 20 日归属当月，21 日归属下月；
  - `PKS`、`PKF`、四台 `S1DP` 机台和启用的 `KPI排除条件` 被排除；
  - 被排除记录仍可生成 Master Data 行，但不能进入 KPI 汇总资格；
  - 分母为 0 时不生成无效比例。

- [ ] **Step 2: Run tests to verify they fail**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: 新增指标和排除测试 FAIL。

- [ ] **Step 3: Implement metrics and row construction**

  使用原始累计值计算比率，只有在分母大于 0 且窗口有效时写入 AU/UPDT。`_pme_buildMasterRow` 统一生成：序列号、工序、保养日期、机台号、Before/After PA、Before/After AU、Before/After UPDT、Gap、YearMonth、SubProcess、Production Status、Performance Status。

- [ ] **Step 4: Run tests to verify they pass**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: Task 1–3 tests PASS。

- [ ] **Step 5: Commit**

  ```bash
  git add '35 - PME计算.js' tests/pme-calculation.test.mjs
  git commit -m "feat: calculate PME metrics and eligibility"
  ```

### Task 4: AEM/PM Data 读取与 Master Data 幂等写入

**Files:**
- Modify: `35 - PME计算.js`
- Modify: `tests/pme-calculation.test.mjs`

**Interfaces:**
- Consumes: Task 1–3 的计划、窗口、指标和行模型函数。
- Produces `_pme_loadSourceData() -> {pmPlans, aemRows, exclusionRules, tfTypeMap}`。
- Produces `_pme_calculateMasterRows(source) -> Array<Array>`。
- Produces `_pme_upsertMasterData(existingRows, calculatedRows) -> Array<Array>`。
- Produces `executePMECalculation(event) -> ExecutionSummary`。

- [ ] **Step 1: Write the failing tests**

  增加 I/O 边界和纯函数测试：

  - 来源表缺少必需表头时，读取层抛出明确错误；
  - AEM 季度来源日期重叠时，写入前中止；
  - PM 计划存在但无对应 AEM PM 时产生告警，不把其他 PM 记录强行匹配；
  - 已有 Master Data 行按序列号更新，新行追加；
  - 重复运行不会增加重复序列号；
  - 空 Before/After 指标不会被写成 0；
  - 无法计算的历史行不会被删除；
  - 15 列输出表头缺失时不执行部分写入。

- [ ] **Step 2: Run tests to verify they fail**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: 新增数据读取和幂等写入测试 FAIL。

- [ ] **Step 3: Implement source loading and safe upsert**

  在配置区集中声明 PME 主表、PM Data、AEM 文件及目标 Sheet。按 `Period`/`ProdDate` 选择 AEM 数据，先完整读取、验证和计算，再一次性更新 `Master Data`。写入过程使用表头映射，不使用固定列位置；失败时不写入部分结果。

- [ ] **Step 4: Run tests to verify they pass**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: Task 1–4 tests PASS。

- [ ] **Step 5: Commit**

  ```bash
  git add '35 - PME计算.js' tests/pme-calculation.test.mjs
  git commit -m "feat: load AEM data and upsert Master Data"
  ```

### Task 5: KPI Summary 汇总、入口和调度接入

**Files:**
- Modify: `35 - PME计算.js`
- Modify: `tests/pme-calculation.test.mjs`
- Modify: `docs/项目记录.md`

**Interfaces:**
- Consumes: Task 4 的 `executePMECalculation` 和 `Master Data` 行模型。
- Produces `_pme_calculateKpiSummary(masterRows, months) -> Object`，统计 IM、TF、PK 及 TF 子工序。
- Produces `_pme_writeKpiSummary(summary, months, sheet) -> void`，保持现有行列和 Goal 规则。
- Produces `executePMECalculation` 的完整流程：读取 → 计算 → 写 Master Data → 写 KPI Summary。

- [ ] **Step 1: Write the failing tests**

  增加以下测试：

  - 只统计 `Production Status=生产` 且 `Performance Status` 为 Good/Bad 的行；
  - IM、TF、PK 主工序分别统计；
  - TF 的 `AFT & USW`、`HSL & Bou`、`Zar` 子工序同步统计；
  - 排除记录不计入 Good、Bad 和 PME；
  - 总数等于 Good + Bad；
  - 2026 年月份和 YTD 写入既有行列；
  - 无数据月份写入 `-`，不保留旧的过期 KPI 数值；
  - 目标值从 KPI Summary 现有 Goal 列读取，达标背景行为与既有布局一致；
  - 主入口完成后返回读取数、计算数、写入数、告警数摘要。

  KPI Summary 的既有布局固定为：2025 行 IM/TF/AFT & USW/HSL & Bou/Zar/PK = 2/3/4/5/6/7，2026 行 = 10/11/12/13/14/15；月份列 D:O，YTD 列 P。实现必须从表头或配置读取并验证这些位置，不得把其他表的行号当作目标。

- [ ] **Step 2: Run tests to verify they fail**

  Run: `node --test tests/pme-calculation.test.mjs`

  Expected: KPI Summary 和完整入口测试 FAIL。

- [ ] **Step 3: Implement KPI aggregation and entry**

  复用旧项目中已确认的 KPI 行列结构，但将数据源固定为新任务刚写入的 `Master Data`。入口函数支持手动调用和 `02 - 定时执行.js` 的调度调用；不修改 `07 - PM_KPI写入.js`。

- [ ] **Step 4: Run the full test suite**

  Run: `node --test tests/*.test.mjs`

  Expected: PME tests and all existing tests PASS，0 failures。

- [ ] **Step 5: Commit**

  ```bash
  git add '35 - PME计算.js' tests/pme-calculation.test.mjs docs/项目记录.md
  git commit -m "feat: add scheduled PME calculation and KPI summary"
  ```

### Task 6: 最终验证与交付检查

**Files:**
- Modify: `docs/项目记录.md` only if the final version/date entry needs correction.

**Interfaces:**
- Consumes: Task 5 的完整任务入口和测试结果。
- Produces: 可部署的 `ScheduledTask` 代码、测试证据和变更摘要。

- [ ] **Step 1: Run the full suite again from a clean working tree**

  Run: `node --test tests/*.test.mjs`

  Expected: 全部测试 PASS。

- [ ] **Step 2: Review the Apps Script file statically**

  检查函数名冲突、常量重复、所有外部 Sheet 名称、表头字段、日志入口和 `07 - PM_KPI写入.js` 是否保持未修改。

- [ ] **Step 3: Review the diff against the approved spec**

  确认没有修改旧 `PM_KPI` 项目、没有重新使用 `Master Data_Excel`、没有给 TF/PK 增加 CO 逻辑，也没有增加 V0–V4 正式输出。

- [ ] **Step 4: Commit any documentation-only correction if needed**

  只有在最终检查发现项目记录需要同步时才提交；不做无关代码重构。
