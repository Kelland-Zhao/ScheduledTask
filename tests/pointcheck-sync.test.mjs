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

// —— Workcenter 最小表头 fixture（按表头名定位，只需 2 个必需列）——
// 行值取自生产表真实机台号；仅大小写/空格做变体，用于锁 spec §3.3/§12 的「写入形态」
const WC_HEADERS_MIN = ['Workcenter', '点检无需检查Y/N'];
const wcRow = (no, flag) => [no, flag || ''];

test('追加写入表11 原值：仅 trim，保留大小写（spec §3.3/§12）', () => {
  // 生产表 328 行 Workcenter 当前恰好都是大写无空格，这里取真实机台号 V2FTA164
  // （spec §14.1 取样）的「空格 + 小写」变体 —— 锁的是**写入形态**：
  // 比对键归一，但落到 D 列的值必须是表11 原值（仅 trim）。
  // 若回归成写归一值（V2FTA164 这类大写形态），只有本用例会红。
  const built = globalThis._pc_buildIncludedSet([WC_HEADERS_MIN, wcRow(' v2fta164 ')]);
  const plan = globalThis._pc_computeSyncPlan(
    [{ rowIndex: 1058, machineNo: 'V1FTA958', rowData: mlRow('INJ', 'TB1', 'OMNI', 'V1FTA958', '技术员') }],
    built);

  assert.deepEqual(plan.append, [{ machineNo: 'v2fta164', workshop: 'TB2' }]);
});

test('比对仍用归一值：表11 原值含空格/小写也能与 MachineList 对上，不追加不删除', () => {
  // 与上一条互为约束：写入用原值，**比对**仍必须用归一值，
  // 否则同一台机会因为大小写/空格差异被误判成「两侧都没有」→ 被追加一份、删除原行
  const built = globalThis._pc_buildIncludedSet([
    WC_HEADERS_MIN,
    wcRow(' h2fcs506 '),    // MachineList 里是 H2FCS506（两行，真实取样 A1051:E1052）
    wcRow(' H2HTA520'),
    wcRow('v1fta958'),
  ]);
  const plan = globalThis._pc_computeSyncPlan(INJ_ROWS_FIXTURE, built);

  assert.deepEqual(plan.append, [], '归一后两侧都有 → 不追加');
  assert.deepEqual(plan.toDelete, [], '归一后两侧都有 → 不删除');
});

test('Y 优先：同号「非 Y 行在前、Y 行在后」→ 整台排除且计入重复告警', () => {
  // 顺序是这条用例的全部：先非 Y 后 Y 才会走到 `delete included[machineNo]` ——
  // 「任一行标 Y 就整台不点检」的关键一步，此前没有一条用例会在它被删掉时失败。
  // 反过来的顺序（先 Y 后非 Y）走的是 flagged 分支，在 workcenter 测试里已有覆盖。
  const built = globalThis._pc_buildIncludedSet([
    WC_HEADERS_MIN,
    wcRow('H2FTA002', ''),     // 非 Y：先进纳入集
    wcRow('H2FTA002', 'Y'),    // 同号标 Y：必须把已进的条目清掉，整台排除
    wcRow('H2FTA001', ''),
  ]);

  assert.deepEqual(Object.keys(built.included), ['H2FTA001'], 'H2FTA002 被 Y 行整台排除');
  assert.equal(built.flagged['H2FTA002'], true);
  assert.deepEqual(built.duplicates, ['H2FTA002'], '同号多行且判定冲突，必须告警而不是静默清掉');
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

test('安全阀：拟删台数不是有限数字 → 拒绝删除并说明（fail closed）', () => {
  // 阀守的是不可逆的删除：算不出/读不到台数时必须当作「不确定」拦下。
  // 原实现 `undefined > 10 === false` → 放行（fail open），是最危险的失败方向。
  // 手动运行也不豁免：台数都读不到，人不可能核对过这份清单
  const missing = globalThis._pc_shouldDelete(undefined, false);
  assert.equal(missing.ok, false, 'undefined 不能被当成 0 台放行');
  assert.match(missing.reason, /数字|无效|无法/, '原因要说明台数无效');

  assert.equal(globalThis._pc_shouldDelete(NaN, false).ok, false);
  assert.equal(globalThis._pc_shouldDelete('', false).ok, false, '空串同样不是有效台数');
  assert.equal(globalThis._pc_shouldDelete(undefined, true).ok, false, '台数读不到时手动也不放行');
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

// —— _pc_mergeDeleteRuns：把任意顺序的 1-based 行号合并成 deleteRows 用的连续区间 ——
// 这是全模块唯一会「删错行」的一步：段内行号必须连续（非目标行天然把段断开），
// 段必须按 start 从大到小返回 —— 删除会让下方行号上移，先删大行号才不会错位
test('删除区间合并：空输入 → []', () => {
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([]), [], '不能无条件读 targets[0]');
});

test('删除区间合并：单行 → 一段', () => {
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([7]), [{ start: 7, count: 1 }]);
});

test('删除区间合并：连续行合成一段', () => {
  // 真实形态：MachineList!A1051:E1053 连着三行都要删（同机多行或相邻两台）
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([1051, 1052, 1053]), [{ start: 1051, count: 3 }]);
});

test('删除区间合并：两段有间隔 → 两段，start 大的在前', () => {
  // 间隔里是没被删的行（其他机台或非 INJ 行），段不能跨过去
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([1051, 1052, 1053, 1058]), [
    { start: 1058, count: 1 },
    { start: 1051, count: 3 },
  ], '先删 1058 再删 1051-1053，行号不会错位');
});

test('删除区间合并：乱序输入与有序结果一致', () => {
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([1058, 1051, 1053, 1052]), [
    { start: 1058, count: 1 },
    { start: 1051, count: 3 },
  ]);
});

test('删除区间合并：重复行号只算一次（否则 count 会多圈进相邻行）', () => {
  assert.deepEqual(globalThis._pc_mergeDeleteRuns([1051, 1051, 1052]), [{ start: 1051, count: 2 }]);
});

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

test('邮件包含两张表的链接', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');

  assert.match(html, /1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY\/edit#gid=436306312/, '点检表 MachineList 页签');
  assert.match(html, /12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM\/edit#gid=0/, '表11 Workcenter 页签');
});

function emptyResult(over) {
  return Object.assign({
    append: [], deleted: [], plasmaKept: [], incomplete: [],
    deleteSkipped: false, valveReason: '', badFlags: [], duplicates: [],
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

test('两侧一致且无待补全 → 报告无差异', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');
  assert.match(html, /无差异|完全一致/);
});

test('零差异但有 badFlags 告警 → 不得出「完全一致」', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult({ badFlags: ['y'] }), '2026-09-30');

  assert.doesNotMatch(html, /完全一致/, '有数据质量告警还说完全一致，邮件自相矛盾');
  assert.match(html, /数据质量/, '告警本身仍须照常渲染');
});

test('零差异但有 duplicates 告警 → 不得出「完全一致」', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult({ duplicates: ['E0EN0001'] }), '2026-09-30');

  assert.doesNotMatch(html, /完全一致/, '有重复机台号还说完全一致，邮件自相矛盾');
  assert.match(html, /数据质量/, '告警本身仍须照常渲染');
});

test('零差异且无数据质量告警 → 仍须出「完全一致」', () => {
  const html = globalThis._pc_buildEmailHtml(emptyResult(), '2026-09-30');

  assert.match(html, /完全一致/, '真·无差异时不能把全清结论一并抑制掉');
});
