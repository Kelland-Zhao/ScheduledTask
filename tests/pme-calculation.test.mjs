import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// 允许在生产文件尚未创建时运行测试，使 RED 阶段失败在“函数未定义”而不是文件读取错误。
const modulePath = new URL('../35 - PME计算.js', import.meta.url);
const moduleCode = fs.existsSync(modulePath) ? fs.readFileSync(modulePath, 'utf8') : '';
(0, eval)(moduleCode);

test('机台号标准化：去除 SW 后缀', () => {
  assert.equal(_pme_normalizeWorkcenter(' T2GB0105【SW】 '), 'T2GB0105');
});

test('工序映射：INJ 映射 IM，TF 和 PK 保持原值', () => {
  assert.equal(_pme_mapProcess('INJ'), 'IM');
  assert.equal(_pme_mapProcess('TF'), 'TF');
  assert.equal(_pme_mapProcess('PK'), 'PK');
});

test('PM 计划筛选：只保留类别含 PM 且状态非取消的记录', () => {
  const headers = ['状态', '工序', '开始时间', 'AEM#', '保养时间', '类别', '开始日期'];
  const headerIndex = _pme_headerIndex(headers);
  const rows = [
    headers,
    ['已确认', 'INJ', '08:00', 'T2GB0105【SW】', 8, 'PM', '2026-01-21'],
    ['取消', 'TF', '08:00', 'AFT205', 8, 'PM', '2026-01-22'],
    ['已确认', 'PK', '08:00', 'PK001', 8, 'CM', '2026-01-23'],
    ['已确认', 'PK', '08:00', 'PK002', 8, 'PM/Other', '2026-01-24'],
  ];

  const result = _pme_selectPmPlans(rows, headerIndex);

  assert.equal(result.length, 2);
  assert.deepEqual(result.map(row => row.workcenter), ['T2GB0105', 'PK002']);
  assert.equal(result[0].keyProcess, 'IM');
  assert.equal(result[0].serialNo, '2026-01-21T2GB0105');
  assert.equal(result[0].durationHours, 8);
});

test('PM 计划读取按表头名称，不依赖列顺序', () => {
  const headers = ['开始日期', '类别', '保养时间', 'AEM#', '状态', '工序', '开始时间'];
  const headerIndex = _pme_headerIndex(headers);
  const rows = [
    headers,
    ['2026-02-03', 'PM', 16, 'AFT205', '已确认', 'TF', '08:00'],
  ];

  const result = _pme_selectPmPlans(rows, headerIndex);

  assert.equal(result.length, 1);
  assert.equal(result[0].pmDate, '2026-02-03');
  assert.equal(result[0].durationHours, 16);
  assert.equal(result[0].keyProcess, 'TF');
});

test('班次排序：同日 1/2/3，班次 3 后接次日班次 1', () => {
  const shifts = [
    { shiftKey: '2026-01-02-1', prodDate: '2026-01-02', shift: 1 },
    { shiftKey: '2026-01-01-3', prodDate: '2026-01-01', shift: 3 },
    { shiftKey: '2026-01-01-2', prodDate: '2026-01-01', shift: 2 },
    { shiftKey: '2026-01-01-1', prodDate: '2026-01-01', shift: 1 },
  ];

  shifts.sort(_pme_compareShift);

  assert.deepEqual(shifts.map(row => row.shiftKey), [
    '2026-01-01-1',
    '2026-01-01-2',
    '2026-01-01-3',
    '2026-01-02-1',
  ]);
});

const AEM_HEADERS = [
  'ProdDate', 'Workcenter', 'Shift', 'Prod_Hours', 'NetRuntime',
  'PlantAvailableTime', 'CO', 'PM',
];

function aemRow(fields) {
  return AEM_HEADERS.map(header => fields[header] ?? '');
}

test('AEM 班次聚合：同机台同日期班次的分段时间分别累加', () => {
  const rows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', Shift: 1, Prod_Hours: 4, NetRuntime: 2, PlantAvailableTime: 8 }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', Shift: 1, Prod_Hours: 3, NetRuntime: 1, PlantAvailableTime: 8 }),
  ];

  const result = _pme_aggregateAemShifts(rows, _pme_headerIndex(AEM_HEADERS));

  assert.equal(result.length, 1);
  assert.equal(result[0].prodHours, 7);
  assert.equal(result[0].netRuntime, 3);
  assert.equal(result[0].plantAvailableTime, 16);
});

test('保养窗口：空开始时间默认早班，4 个班次后从第 5 个班次开始', () => {
  const plan = {
    pmDate: '2026-01-21',
    startTime: '',
    durationHours: 32,
  };

  const result = _pme_getMaintenanceWindow(plan);

  assert.deepEqual(result, {
    startShiftKey: '2026-01-21-2',
    endShiftKey: '2026-01-22-2',
    afterStartShiftKey: '2026-01-22-3',
  });
});

test('IM Before：连续生产时轮空保养前最后一个生产班次', () => {
  const shifts = [
    { shiftKey: '2026-01-20-1', prodDate: '2026-01-20', shift: 1, prodHours: 8 },
    { shiftKey: '2026-01-20-2', prodDate: '2026-01-20', shift: 2, prodHours: 8 },
    { shiftKey: '2026-01-20-3', prodDate: '2026-01-20', shift: 3, prodHours: 8 },
    { shiftKey: '2026-01-21-2', prodDate: '2026-01-21', shift: 2, prodHours: 0 },
  ];

  const result = _pme_selectBeforeShifts('IM', shifts, '2026-01-21-2');

  assert.deepEqual(result.map(row => row.shiftKey), ['2026-01-20-1', '2026-01-20-2']);
});

test('IM Before：每个生产→空班边界的生产班次都轮空，空班不进入有效生产集合', () => {
  const shifts = [
    { shiftKey: '2026-01-19-1', prodDate: '2026-01-19', shift: 1, prodHours: 8 },
    { shiftKey: '2026-01-19-2', prodDate: '2026-01-19', shift: 2, prodHours: 0 },
    { shiftKey: '2026-01-19-3', prodDate: '2026-01-19', shift: 3, prodHours: 6 },
    { shiftKey: '2026-01-20-1', prodDate: '2026-01-20', shift: 1, prodHours: 0 },
    { shiftKey: '2026-01-20-2', prodDate: '2026-01-20', shift: 2, prodHours: 5 },
    { shiftKey: '2026-01-20-3', prodDate: '2026-01-20', shift: 3, prodHours: 0 },
    { shiftKey: '2026-01-21-2', prodDate: '2026-01-21', shift: 2, prodHours: 0 },
  ];

  const result = _pme_selectBeforeShifts('IM', shifts, '2026-01-21-2');

  assert.deepEqual(result, []);
});

test('TF 和 PK Before：不应用 IM 轮空规则', () => {
  const shifts = [
    { shiftKey: '2026-01-20-2', prodDate: '2026-01-20', shift: 2, prodHours: 8 },
    { shiftKey: '2026-01-20-3', prodDate: '2026-01-20', shift: 3, prodHours: 0 },
    { shiftKey: '2026-01-21-2', prodDate: '2026-01-21', shift: 2, prodHours: 0 },
  ];

  assert.deepEqual(
    _pme_selectBeforeShifts('TF', shifts, '2026-01-21-2').map(row => row.shiftKey),
    ['2026-01-20-2'],
  );
  assert.deepEqual(
    _pme_selectBeforeShifts('PK', shifts, '2026-01-21-2').map(row => row.shiftKey),
    ['2026-01-20-2'],
  );
});

test('After：累计到首次达到或超过 144 小时，保留超出 144 的最后一班', () => {
  const shifts = [
    { shiftKey: '2026-01-22-3', prodDate: '2026-01-22', shift: 3, prodHours: 50 },
    { shiftKey: '2026-01-23-1', prodDate: '2026-01-23', shift: 1, prodHours: 0 },
    { shiftKey: '2026-01-23-2', prodDate: '2026-01-23', shift: 2, prodHours: 100 },
    { shiftKey: '2026-01-23-3', prodDate: '2026-01-23', shift: 3, prodHours: 20 },
    { shiftKey: '2026-01-24-1', prodDate: '2026-01-24', shift: 1, prodHours: 8 },
  ];

  const result = _pme_selectAfterWindow('IM', shifts, '2026-01-22-3');

  assert.equal(result.status, '生产');
  assert.equal(result.totalProdHours, 150);
  assert.deepEqual(result.rows.map(row => row.shiftKey), [
    '2026-01-22-3', '2026-01-23-1', '2026-01-23-2',
  ]);
});

function metricShift(fields) {
  return {
    shiftKey: fields.shiftKey || '2026-01-01-1',
    prodDate: fields.prodDate || '2026-01-01',
    shift: fields.shift || 1,
    prodHours: fields.prodHours ?? 0,
    netRuntime: fields.netRuntime ?? 0,
    plantAvailableTime: fields.plantAvailableTime ?? 0,
  };
}

test('指标公式：AU 使用 PlantAvailableTime，UPDT 使用 Prod_Hours，Gap 为 After 减 Before', () => {
  const before = [metricShift({ prodHours: 100, netRuntime: 70, plantAvailableTime: 80 })];
  const after = [metricShift({ prodHours: 150, netRuntime: 90, plantAvailableTime: 120 })];

  const result = _pme_calculateMetrics(before, after);

  assert.equal(result.beforeAU, 0.875);
  assert.equal(result.afterAU, 0.75);
  assert.equal(result.beforeUPDT, 0.3);
  assert.equal(result.afterUPDT, 0.4);
  assert.ok(Math.abs(result.updtGap - 0.1) < 1e-12);
});

test('绩效判定：达到 144 且 After UPDT 不高于 Before 时为 Good', () => {
  const result = _pme_calculateMetrics(
    [metricShift({ prodHours: 100, netRuntime: 70, plantAvailableTime: 80 })],
    [metricShift({ prodHours: 150, netRuntime: 120, plantAvailableTime: 120 })],
  );

  assert.equal(result.productionStatus, '生产');
  assert.equal(result.performanceStatus, '好/Good');
});

test('未达 144 或 Before UPDT 缺失时不强行生成绩效状态', () => {
  const notEnough = _pme_calculateMetrics(
    [metricShift({ prodHours: 100, netRuntime: 70, plantAvailableTime: 80 })],
    [metricShift({ prodHours: 100, netRuntime: 60, plantAvailableTime: 120 })],
  );
  const missingBefore = _pme_calculateMetrics(
    [],
    [metricShift({ prodHours: 150, netRuntime: 60, plantAvailableTime: 120 })],
  );

  assert.equal(notEnough.productionStatus, '未达144');
  assert.equal(notEnough.performanceStatus, '');
  assert.equal(missingBefore.productionStatus, '生产');
  assert.equal(missingBefore.performanceStatus, '');
});

test('日期切点：20 日归当月，21 日归下月', () => {
  assert.equal(_pme_getYearMonth('2026-09-20'), '202609');
  assert.equal(_pme_getYearMonth('2026-09-21'), '202610');
  assert.equal(_pme_getYearMonth('2026-09-26'), '202610');
});

test('KPI 排除：PKS、PKF、S1DP 四台设备和启用规则均被排除', () => {
  const rules = [{ workcenter: 'H2FCS506', duration: 8 }];
  const plan = (keyProcess, workcenter, durationHours) => ({ keyProcess, workcenter, durationHours });

  assert.equal(_pme_isKpiExcluded(plan('PK', 'PKS001', 8), rules), true);
  assert.equal(_pme_isKpiExcluded(plan('PK', 'PKF001', 8), rules), true);
  assert.equal(_pme_isKpiExcluded(plan('IM', 'S1DP0001', 8), rules), true);
  assert.equal(_pme_isKpiExcluded(plan('IM', 'S1DP0004', 8), rules), true);
  assert.equal(_pme_isKpiExcluded(plan('IM', 'H2FCS506', 8), rules), true);
  assert.equal(_pme_isKpiExcluded(plan('IM', 'H2FCS506', 16), rules), false);
});

test('Master Data 行：被排除记录仍可生成 15 列明细，但不具备 KPI 资格', () => {
  const metrics = _pme_calculateMetrics(
    [metricShift({ prodHours: 100, netRuntime: 70, plantAvailableTime: 80 })],
    [metricShift({ prodHours: 150, netRuntime: 120, plantAvailableTime: 120 })],
  );
  const plan = {
    serialNo: '2026-09-26PKS001',
    keyProcess: 'PK',
    pmDate: '2026-09-26',
    workcenter: 'PKS001',
  };

  const row = _pme_buildMasterRow(plan, metrics, '', { excluded: true });

  assert.equal(row.length, 15);
  assert.equal(row[13], '生产');
  assert.equal(row[14], '好/Good');
  assert.equal(_pme_calculateKpiEligibility(row), true);
  assert.equal(_pme_isKpiExcluded(plan, []), true);
});

test('分母为 0 时不生成无效比例', () => {
  const result = _pme_calculateMetrics(
    [metricShift({ prodHours: 0, netRuntime: 1, plantAvailableTime: 0 })],
    [metricShift({ prodHours: 150, netRuntime: 60, plantAvailableTime: 0 })],
  );

  assert.equal(result.beforeAU, null);
  assert.equal(result.beforeUPDT, null);
  assert.equal(result.afterAU, null);
  assert.equal(result.afterUPDT, 0.6);
  assert.equal(result.performanceStatus, '');
});
