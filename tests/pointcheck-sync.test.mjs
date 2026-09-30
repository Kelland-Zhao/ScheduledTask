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
