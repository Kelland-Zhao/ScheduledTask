# PME 计算任务迁移设计

## 1. 目标与范围

在 `ScheduledTask` 中新增独立的 PME（Post-PM Effectiveness）计算任务，替代旧 `PM_KPI` 项目的 PME 计算部分。

新任务负责：

1. 从正式 PM 计划读取保养事件；
2. 从 AEM `SourceData` 读取生产班次数据；
3. 计算 IM、TF、PK 的保养前后 PME 明细；
4. 将明细结果写入现有 PME 主表的 `Master Data`；
5. 根据 `Master Data` 汇总并写入 `KPI Summary`。

本任务不负责将月度 PME 写入奖金表。现有 `07 - PM_KPI写入.js` 是独立的下游任务，继续负责读取计算完成的 PME 并写入指定的月度奖金表。

本任务不修改旧的 `PM_KPI` 项目，也不使用 `Master Data_Excel` 作为 PME 计算源。

## 2. 已确认的业务规则

### 2.1 PM 计划匹配

正式保养必须同时满足：

- `类别` 包含 `PM`；
- `状态` 不等于 `取消`；
- 使用计划中的 `开始日期`、`开始时间`、`AEM#`、`保养时间`和`工序`；
- AEM 中单独出现、但 PM 计划中不存在的 PM 数据不作为正式保养。

工序映射：

- PM 计划 `INJ` 映射为 `IM`；
- `TF` 保留为 `TF`；
- `PK` 保留为 `PK`。

机台号匹配前必须标准化，至少支持：

```text
T2GB0105【SW】 → T2GB0105
```

计划开始时间为空时，默认按计划开始日期当天 08:00，即早班开始。

如果计划存在但 AEM 中没有对应 PM 记录，任务记录数据质量告警，并不把其他 AEM PM 记录强行匹配到该计划。

### 2.2 班次与保养窗口

AEM 班次固定为：

- `1`：夜班；
- `2`：早班；
- `3`：中班。

实际排序按日期加班次顺序建立，不能只按班次数字排序。

计划保养按连续时间计算，每个班次为 8 小时。保养期间的班次不计入 AFTER PM。

AFTER PM 从计划保养结束后的第一个完整班次开始，并累加 `Prod_Hours`，直到第一次达到或超过 144 小时。累计值可以超过 144 小时。

### 2.3 IM 规则

IM 使用更新后的保养前规则：

- BEFORE PM 截止于计划开始班次；
- 先将 AEM 明细按机台、日期、班次聚合；
- 连续生产至保养时，轮空保养前最后一个 `Prod_Hours > 0` 的完整生产班次；
- 如果保养前存在多次开停机，则对每个“生产班次后紧接 `Prod_Hours = 0` 空班”的边界，轮空该空班之前的生产班次；
- `Prod_Hours = 0` 的空班不增加 144 小时累计，也不单独作为生产班次轮空；
- 轮空以完整班次为单位，不能只删除 AEM 的一条分段记录；
- 未被轮空的有效生产班次用于计算 BEFORE PM；
- AFTER PM 不使用 CO 轮空逻辑。

### 2.4 TF 规则

TF 保持原有不考虑 CO 的逻辑：

- 不因为 `CO > 0` 删除 AEM 分段；
- 不因为 `CO > 0` 删除整班；
- 不删除 CO 后的下一个或多个班次；
- 按正常保养窗口累计 `Prod_Hours`。

### 2.5 PK 规则

PK 纳入 PME 计算，保持原有逻辑：

- 不使用 IM 的保养前最后生产班次轮空规则；
- 不使用 TF 的 CO 场景逻辑；
- 按正常保养窗口累计和计算。

以下设备不纳入最终 KPI：

- `PKS` 开头的机台号；
- `PKF` 开头的机台号；
- `S1DP0001`、`S1DP0002`、`S1DP0003`、`S1DP0004`；
- `KPI排除条件` 中状态为 `启用` 的具体排除规则。

被排除的记录仍可以保留在 `Master Data` 用于追溯，但不得计入 `KPI Summary` 的 Good、Bad 和 PME。

### 2.6 KPI 公式与状态

使用以下字段映射：

```text
Prod_Hours          → Production_AvailableTime
NetRuntime          → NetRuntime
PlantAvailableTime  → PlantAvailableTime
```

公式：

```text
AU   = NetRuntime / PlantAvailableTime
UPDT = 1 - NetRuntime / Prod_Hours
Gap  = UPDT_AFTER - UPDT_BEFORE
```

不得使用 `1 - AU` 替代 UPDT。

生产状态：

```text
AFTER 累计 Prod_Hours >= 144 → 生产
AFTER 累计 Prod_Hours < 144  → 未达144
```

只有生产状态为 `生产`，且 Before/After UPDT 均为有效数值时，才判断绩效状态：

```text
UPDT_AFTER <= UPDT_BEFORE → 好/Good
UPDT_AFTER > UPDT_BEFORE  → 差/Bad
```

未达到 144 小时或缺少有效 Before 窗口时，绩效状态保持空白，不将空值转换成 0 后强行判断。

### 2.7 保养归属月份

按 PM 计划开始日期的 20 日切点生成 `YearMonth`：

- 日期 1–20 日：归属当月；
- 日期 21 日至月末：归属下月。

例如 `2026-09-26` 归属 `202610`。

### 2.8 PME 汇总

```text
PME = Good / (Good + Bad)
```

`KPI Summary` 保留现有布局和工序映射：

- IM；
- TF；
- AFT & USW；
- HSL & Bou；
- Zar；
- PK。

TF 子工序继续通过 `SubProcess` 汇总。目标值读取现有 `KPI Summary` 的 Goal 列，不在代码中新增或编造目标值。现有月份列和 YTD 结构保持不变。

## 3. 数据源与数据路由

任务运行时读取：

1. PME 主表中的 `PM Data`；
2. 已登记的 AEM 归档文件和当前文件的 `SourceData`；
3. 必要时读取 2025 年历史 AEM 数据，补足 2026 年保养的 BEFORE 窗口；
4. PME 主表中的 `KPI排除条件` 和 TF 机型映射数据。

AEM 文件通过 `Period` 和 `ProdDate` 选择，不根据文件名推断年份。当前季度使用当前文件，已关闭季度使用对应归档文件；同一期间不重复合并。

计算前必须检查：

- SourceData 表头；
- 日期覆盖范围；
- 数据块首尾行和有效行数；
- 不同来源之间是否重叠；
- 关键字段是否存在。

关键 AEM 字段按表头名称定位，不依赖固定列号：

- `ProdDate`
- `Workcenter`
- `Shift`
- `PM`
- `Prod_Hours`
- `NetRuntime`
- `PlantAvailableTime`
- `process`
- `WorkcenterGroup`
- `CO`

IM 只使用 AEM 中 `process = IM` 的记录，并排除 `WorkcenterGroup = OPK` 和 `WorkcenterGroup = Dummy`。TF 和 PK 使用各自对应工序的数据。

## 4. 代码结构设计

新增一个独立脚本文件，例如：

```text
35 - PME计算.js
```

所有内部函数使用 `_pme_` 前缀，避免与 `ScheduledTask` 中已有函数冲突。

建议分成以下层次：

### 4.1 配置层

集中维护：

- PME 主表 ID；
- AEM 来源文件配置；
- 目标 Sheet 名称；
- 处理年份或数据截止日期；
- 15 字段 `Master Data` 表头名称；
- `KPI Summary` 工序行和月份列映射。

不在代码中保存密码、Token 或服务账号凭据。

### 4.2 数据读取层

负责：

- 读取 PM Data；
- 读取 AEM SourceData；
- 读取排除条件；
- 读取 TF Machine Type；
- 使用表头建立字段映射；
- 返回结构化记录。

### 4.3 计算层

使用纯函数完成：

- 日期和班次排序；
- 机台号标准化；
- PM 计划筛选和工序映射；
- AEM 班次聚合；
- IM Before 轮空；
- PM 前后窗口识别；
- 144 小时累计；
- AU、UPDT、Gap 和 Good/Bad 计算。

计算层不直接访问 Spreadsheet 服务，便于使用 Node 测试。

### 4.4 写入层

负责：

- 按序列号更新或新增 `Master Data`；
- 保持统一的 15 列字段结构；
- 保留原有非目标内容；
- 从 `Master Data` 汇总 `KPI Summary`；
- 写入现有 KPI Summary 的月份和 YTD 区域。

写入前先完成全部计算和校验，避免源数据异常时只写入部分结果。

### 4.5 入口层

提供：

- 手动入口：用于验证和补算；
- 定时入口：供 `02 - 定时执行.js` 调度；
- 结果日志：记录读取范围、计划数、匹配数、计算数、写入数和数据质量告警数。

入口函数失败时不得吞掉异常；应记录失败日志并中止本次写入。

## 5. 幂等与写入策略

PME 明细以标准化的 `PM 日期 + AEM#` 作为序列号主键。

每次运行：

1. 读取现有 `Master Data`；
2. 按序列号建立索引；
3. 对已存在记录更新计算字段；
4. 对新记录追加；
5. 不删除无法在本次 AEM 窗口中计算的历史记录；
6. 不将空的 Before/After 指标覆盖成 0。

重复运行同一数据范围时，结果应保持一致，不产生重复序列号。

## 6. 失败保护与数据质量

下列情况应中止写入并记录失败：

- 必需 Sheet 不存在；
- 必需表头缺失；
- AEM 来源存在日期重叠且无法判定覆盖关系；
- PM Data 或 AEM SourceData 读取为空且无法判断是否为合法空数据；
- 输出表头与预期字段不一致。

下列情况记录告警，但不应把数据静默当作有效结果：

- PM 计划没有对应 AEM PM；
- 机台号标准化后仍无法匹配；
- Before 不能形成有效窗口；
- After 截止当前数据仍未达到 144 小时；
- UPDT 字段缺失或分母为 0。

## 7. 测试设计

在 `tests/pme-calculation.test.mjs` 中为纯计算层增加测试，至少覆盖：

- `INJ` 映射为 `IM`；
- 机台号后缀标准化；
- PM 计划筛选条件；
- 计划外 AEM PM 不匹配；
- 班次 1/2/3 的实际排序；
- IM 连续生产时轮空最后生产班次；
- IM 多次开停时轮空每个停机边界前的生产班次；
- `Prod_Hours = 0` 不增加累计值；
- After 累计超过 144 小时；
- TF 的 CO 不影响纳入数据；
- PK 的正常计算；
- PKS/PKF 和 S1DP 排除；
- AU、UPDT、Gap 和 Good/Bad；
- 缺失 UPDT 时保持绩效状态空白；
- 20 日保养月份切点；
- 重复运行不产生重复记录。

现有 123 个测试必须继续通过。

## 8. 非目标范围

本次不做以下修改：

- 不修改旧 `/Users/kelland/gas-projects/PM_KPI/` 项目；
- 不修改 `ScheduledTask/07 - PM_KPI写入.js` 的月度奖金表写入逻辑；
- 不新增 V0/V1/V2/V3/V4 作为正式输出目标；
- 不改变 TF 的 CO 处理逻辑；
- 不改变 PK 的原有计算逻辑；
- 不把 `Master Data_Excel` 重新作为计算源；
- 不修改无关的 ScheduledTask 模块。
