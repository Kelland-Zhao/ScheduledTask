import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// 允许在生产文件尚未创建时运行测试，使 RED 阶段失败在“函数未定义”而不是文件读取错误。
const modulePath = new URL('../35 - PME计算.js', import.meta.url);
const moduleCode = fs.existsSync(modulePath) ? fs.readFileSync(modulePath, 'utf8') : '';
(0, eval)(moduleCode);

test('大批量行追加：不因参数展开触发调用栈溢出', () => {
  const target = [];
  const additions = new Array(134036).fill(null).map((_, index) => [index]);

  _pme_appendRows(target, additions);

  assert.equal(target.length, additions.length);
  assert.deepEqual(target[0], [0]);
  assert.deepEqual(target[target.length - 1], [134035]);
});

test('机台号标准化：去除 SW 后缀', () => {
  assert.equal(_pme_normalizeWorkcenter(' T2GB0105【SW】 '), 'T2GB0105');
});

test('工序映射：INJ 映射 IM，TF 和 PK 保持原值', () => {
  assert.equal(_pme_mapProcess('INJ'), 'IM');
  assert.equal(_pme_mapProcess('TF'), 'TF');
  assert.equal(_pme_mapProcess('PK'), 'PK');
});

test('子工序映射：IM/PK 使用自身工序，TF 使用机台类型映射', () => {
  assert.equal(_pme_getSubprocess({ keyProcess: 'IM', workcenter: 'H2FCS903' }, {}), 'IM');
  assert.equal(_pme_getSubprocess({ keyProcess: 'PK', workcenter: 'P1HA6001' }, {}), 'PK');
  assert.equal(_pme_getSubprocess({ keyProcess: 'TF', workcenter: 'T2GR0126' }, { T2GR0126: 'Zar' }), 'Zar');
  assert.equal(_pme_getSubprocess({ keyProcess: 'TF', workcenter: 'T2GR0126' }, {}), '');
});

test('Master Data 清理：按 PM_Date+机台号去除错行和重复行，并保留 P:R 填写内容', () => {
  const headers = _pme_MASTER_HEADERS.concat([
    '根本原因/  Root cause',
    '预防措施/Preventive action',
    '状态/Status',
  ]);
  const existingRow = (serialNo, pmDate, rootCause = '', preventiveAction = '', status = '') => {
    const row = new Array(headers.length).fill('');
    row[0] = serialNo;
    row[1] = 'TF';
    row[2] = pmDate;
    row[3] = 'T1AFT307';
    row[15] = rootCause;
    row[16] = preventiveAction;
    row[17] = status;
    return row;
  };
  const calculatedRow = new Array(_pme_MASTER_HEADERS.length).fill('');
  calculatedRow[0] = '2026-09-21T1AFT307';
  calculatedRow[1] = 'TF';
  calculatedRow[2] = '2026-09-21';
  calculatedRow[3] = 'T1AFT307';

  const result = _pme_upsertMasterData(
    [
      headers,
      existingRow('2026-09-21T1AFT307', '2026-09-21'),
      existingRow('2026-09-21T1AFT307', '2026-09-09', '原因', '措施', '待确认'),
    ],
    [calculatedRow],
    [{ serialNo: '2026-09-21T1AFT307', pmDate: '2026-09-21', workcenter: 'T1AFT307' }],
    2026,
    [],
  );

  assert.equal(result.length, 2);
  assert.deepEqual(result[1].slice(0, 4), ['2026-09-21T1AFT307', 'TF', '2026-09-21', 'T1AFT307']);
  assert.deepEqual(result[1].slice(15, 18), ['原因', '措施', '待确认']);
});

test('PM 计划筛选：只保留当年且不晚于执行日的有效计划，并跳过带 SW 的机台', () => {
  const headers = ['状态', '工序', '开始时间', 'AEM#', '保养时间', '类别', '开始日期'];
  const headerIndex = _pme_headerIndex(headers);
  const rows = [
    headers,
    ['已确认', 'INJ', '08:00', 'T2GB0105【SW】', 8, 'PM', '2026-01-21'],
    ['已确认', 'INJ', '08:00', 'T2GB0106', 8, 'PM', '2025-12-31'],
    ['取消', 'TF', '08:00', 'AFT205', 8, 'PM', '2026-01-22'],
    ['已确认', 'PK', '08:00', 'PK001', 8, 'CM', '2026-01-23'],
    ['已确认', 'PK', '08:00', 'PK002', 8, 'PM/Other', '2026-01-24'],
    ['已确认', 'PK', '08:00', 'PK003', 8, 'PM', '2026-10-19'],
    ['已确认', 'PK', '08:00', 'PK004', 8, 'PM', '2026-10-18'],
  ];

  const result = _pme_selectPmPlans(rows, headerIndex, 2026, '2026-10-18');

  assert.equal(result.length, 2);
  assert.deepEqual(result.map(row => row.workcenter), ['PK002', 'PK004']);
  assert.equal(result[0].keyProcess, 'PK');
  assert.equal(result[0].serialNo, '2026-01-24PK002');
  assert.equal(result[0].durationHours, 8);
});

test('PM 计划筛选：未来日期不产生计划，也不产生后续 AEM 未匹配告警', () => {
  const headers = ['状态', '工序', '开始时间', 'AEM#', '保养时间', '类别', '开始日期'];
  const headerIndex = _pme_headerIndex(headers);
  const rows = [
    headers,
    ['已确认', 'TF', '08:00', 'AFT205', 8, 'PM', '2026-10-19'],
  ];

  const result = _pme_selectPmPlans(rows, headerIndex, 2026, '2026-10-18');

  assert.deepEqual(result, []);
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
  'PlantAvailableTime', 'CO', 'PM', 'process', 'WorkcenterGroup',
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

test('AEM 计划索引：按唯一机台号分组，并排除 IM 的 OPK/Dummy 行', () => {
  const rows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'IM', WorkcenterGroup: 'Main' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'IM', WorkcenterGroup: 'OPK' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'TF', WorkcenterGroup: 'Main' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M2', process: 'IM', WorkcenterGroup: 'Dummy' }),
  ];

  const index = _pme_buildAemPlanIndex(rows, _pme_headerIndex(AEM_HEADERS), new Set(['M1']));

  assert.equal(index.M1.length, 2);
  assert.equal(index.M2, undefined);
});

test('AEM 数据预筛选：只保留有效 PM 计划涉及的机台', () => {
  const rows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'IM' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1【SW】', process: 'IM' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M2', process: 'TF' }),
  ];

  const result = _pme_filterAemRowsByWorkcenters(
    rows,
    _pme_headerIndex(AEM_HEADERS),
    new Set(['M1']),
  );

  assert.equal(result.length, 2);
  assert.equal(result[1][1], 'M1');
});

test('按机台索引取数时仍按计划工序过滤，避免跨工序混入计算', () => {
  const rows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'IM' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', process: 'TF' }),
  ];
  const headerIndex = _pme_headerIndex(AEM_HEADERS);
  const index = _pme_buildAemPlanIndex(rows, headerIndex, new Set(['M1']));

  const result = _pme_getAemRowsForPlan(
    { workcenter: 'M1', keyProcess: 'IM' },
    index,
    AEM_HEADERS,
    headerIndex,
  );

  assert.equal(result.length, 2);
  assert.equal(result[1][8], 'IM');
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

test('Before：从 PM 前最近班次向前累计，首次达到或超过 144 小时即停止', () => {
  const shifts = [
    { shiftKey: '2026-01-20-1', prodDate: '2026-01-20', shift: 1, prodHours: 50 },
    { shiftKey: '2026-01-20-2', prodDate: '2026-01-20', shift: 2, prodHours: 80 },
    { shiftKey: '2026-01-20-3', prodDate: '2026-01-20', shift: 3, prodHours: 70 },
    { shiftKey: '2026-01-21-2', prodDate: '2026-01-21', shift: 2, prodHours: 0 },
  ];

  const result = _pme_selectBeforeShifts('TF', shifts, '2026-01-21-2');

  assert.deepEqual(result.map(row => row.shiftKey), ['2026-01-20-2', '2026-01-20-3']);
  assert.equal(result.reduce((total, row) => total + row.prodHours, 0), 150);
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

test('Master Data 行：Before/After Production Available Time 四舍五入为整数', () => {
  const row = _pme_buildMasterRow(
    { serialNo: '2026-01-01M1', keyProcess: 'IM', pmDate: '2026-01-01', workcenter: 'M1' },
    {
      beforeProductionAvailableTime: 144.6,
      afterProductionAvailableTime: 150.4,
      beforeAU: 0.8,
      afterAU: 0.8,
      beforeUPDT: 0.2,
      afterUPDT: 0.21,
      updtGap: 0.01,
      productionStatus: '生产',
      performanceStatus: '差/Bad',
    },
    '',
  );

  assert.equal(row[4], 145);
  assert.equal(row[5], 150);
  assert.equal(row[10], 0.01);
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

test('已排除计划的明细行构造仍保持 15 列，但不具备 KPI 资格', () => {
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

test('Master Data 计算结果不生成 S1DP0001–0004 记录', () => {
  const result = _pme_calculateMasterRows({
    pmPlans: [{
      serialNo: '2026-01-02S1DP0001',
      keyProcess: 'IM',
      pmDate: '2026-01-02',
      workcenter: 'S1DP0001',
      startTime: '08:00',
      durationHours: 8,
    }, {
      serialNo: '2026-01-02M1',
      keyProcess: 'IM',
      pmDate: '2026-01-02',
      workcenter: 'M1',
      startTime: '08:00',
      durationHours: 8,
    }],
    aemRows: [AEM_HEADERS],
    aemHeaderIndex: _pme_headerIndex(AEM_HEADERS),
    exclusionRules: [],
    tfTypeMap: {},
  });

  assert.equal(result.length, 0);
  assert.deepEqual(result.warnings, ['未找到对应 AEM PM 记录: 2026-01-02M1']);
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

test('来源表头校验：缺少必需字段时抛出明确错误', () => {
  assert.throws(
    () => _pme_requireHeaders(['ProdDate', 'Workcenter'], ['ProdDate', 'PM'], 'AEM SourceData'),
    /AEM SourceData.*PM/,
  );
});

test('来源覆盖校验：日期范围重叠时中止', () => {
  assert.throws(
    () => _pme_validateSourceCoverage([
      { label: '2026Q1', start: '2026-01-01', end: '2026-03-31' },
      { label: 'current', start: '2026-03-31', end: '2026-04-30' },
    ]),
    /重叠/,
  );
});

test('AEM 来源读取：按配置的 ProdDate 范围过滤记录', () => {
  const rows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2025-12-31', Workcenter: 'M1', Shift: 1, process: 'IM' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', Shift: 1, process: 'IM' }),
    aemRow({ ProdDate: '2026-04-01', Workcenter: 'M1', Shift: 1, process: 'IM' }),
  ];

  const result = _pme_filterAemRowsBySourceRange(rows, _pme_headerIndex(AEM_HEADERS), {
    label: '2026Q1',
    start: '2026-01-01',
    end: '2026-03-31',
  });

  assert.deepEqual(result.slice(1).map(row => row[0]), ['2026-01-01']);
});

test('AEM 来源读取：配置日期范围内没有有效数据时中止', () => {
  assert.throws(
    () => _pme_filterAemRowsBySourceRange(
      [AEM_HEADERS, aemRow({ ProdDate: '2025-12-31', Workcenter: 'M1', Shift: 1, process: 'IM' })],
      _pme_headerIndex(AEM_HEADERS),
      { label: '2026Q1', start: '2026-01-01', end: '2026-03-31' },
    ),
    /2026Q1.*日期范围内没有有效数据/,
  );
});

test('PM 计划没有 AEM PM 记录时返回告警，不强行匹配其他记录', () => {
  const plan = {
    serialNo: '2026-01-02M1',
    keyProcess: 'IM',
    workcenter: 'M1',
    pmDate: '2026-01-02',
    startTime: '08:00',
    durationHours: 8,
  };
  const aemRows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-02', Workcenter: 'M1', Shift: 2, PM: 0, process: 'IM' }),
  ];

  const result = _pme_calculateMasterRows({
    pmPlans: [plan],
    aemRows,
    aemHeaderIndex: _pme_headerIndex(AEM_HEADERS),
    exclusionRules: [],
    tfTypeMap: {},
  });

  assert.equal(result.length, 0);
  assert.match(result.warnings[0], /2026-01-02M1/);
  assert.deepEqual(Array.from(result.unmatchedPlanSerials), ['2026-01-02M1']);
});

test('Master Data：2025 保留，2026 清理失效计划并保留缺 AEM 的有效计划', () => {
  const headers = [
    '序列号/ Serial No.', '工序/ KeyProcess', '保养日期/ PM_Date', '机台号/ Workcenter',
    'Production_AvailableTime_BEFORE PM', 'Production_AvailableTime_AFTER PM',
    'AU_BEFORE PM', 'AU_AFTER PM', 'UPDT_BEFORE PM', 'UPDT_AFTER PM', 'UPDT Gap',
    '保养月份/ YearMonth', '子工序/ SubProcess', '生产状态/ Production Status',
    '绩效状态/ Performance Status',
  ];
  const historical2025 = ['2025-12-31M0', 'IM', '2025-12-31', 'M0', 8, 150, 0.5, 0.6, 0.2, 0.3, 0.1, '202512', '', '生产', '差/Bad'];
  const oldRow = ['2026-01-01M1', 'IM', '2026-01-01', 'M1', 8, 150, 0.5, 0.6, 0.2, 0.3, 0.1, '202601', '', '生产', '差/Bad'];
  const stale2026 = ['2026-01-02M2', 'PK', '2026-01-02', 'M2', 8, 150, 0.5, 0.5, 0.2, 0.2, 0, '202601', '', '生产', '好/Good'];
  const missingAem2026 = ['2026-01-04M4', 'TF', '2026-01-04', 'M4', 8, 150, 0.5, 0.5, 0.2, 0.2, 0, '202601', '', '未达144', ''];
  const calculated = [
    ['2026-01-01M1', 'IM', '2026-01-01', 'M1', 16, 150, 0.4, 0.4, 0.1, 0.1, 0, '202601', '', '生产', '好/Good'],
    ['2026-01-03M3', 'TF', '2026-01-03', 'M3', 20, 160, 0.4, 0.4, null, 0.2, null, '202601', 'AFT & USW', '生产', ''],
  ];
  const pmPlans = [
    { serialNo: '2026-01-01M1', workcenter: 'M1' },
    { serialNo: '2026-01-04M4', workcenter: 'M4' },
  ];

  const first = _pme_upsertMasterData(
    [headers, historical2025, oldRow, stale2026, missingAem2026],
    calculated,
    pmPlans,
    2026,
    ['2026-01-04M4'],
  );
  const second = _pme_upsertMasterData(first, calculated, pmPlans, 2026, ['2026-01-04M4']);

  assert.equal(first.length, 5);
  assert.deepEqual(first[1], historical2025);
  assert.equal(first[2][4], 16);
  assert.deepEqual(first[3].slice(0, 4), missingAem2026.slice(0, 4));
  assert.deepEqual(first[3].slice(4, 11), ['', '', '', '', '', '', '']);
  assert.equal(first[3][11], missingAem2026[11]);
  assert.equal(first[3][12], '');
  assert.equal(first[3][13], '未匹配到保养记录');
  assert.equal(first[3][14], '');
  assert.equal(first[4][0], '2026-01-03M3');
  assert.equal(first[4][8], null);
  assert.equal(second.length, 5);
  assert.equal(second.filter(row => row[0] === '2026-01-03M3').length, 1);
  assert.equal(second.filter(row => row[0] === '2026-01-02M2').length, 0);
});

test('Master Data 未匹配到保养记录：清空旧 PME 结果但保留 P:R 技术员字段', () => {
  const headers = _pme_MASTER_HEADERS.concat([
    '根本原因/  Root cause', '预防措施/Preventive action', '状态/Status',
  ]);
  const existing = [headers, [
    '2026-01-04M4', 'TF', '2026-01-04', 'M4', 8, 150, 0.5, 0.5, 0.2, 0.2, 0,
    '202601', 'Zar', '生产', '好/Good', '历史原因', '历史措施', '待确认',
  ]];

  const merged = _pme_upsertMasterData(
    existing,
    [],
    [{ serialNo: '2026-01-04M4', workcenter: 'M4' }],
    2026,
    ['2026-01-04M4'],
  );

  assert.deepEqual(merged[1].slice(0, 4), existing[1].slice(0, 4));
  assert.deepEqual(merged[1].slice(4, 11), ['', '', '', '', '', '', '']);
  assert.equal(merged[1][12], '');
  assert.equal(merged[1][13], '未匹配到保养记录');
  assert.equal(merged[1][14], '');
  assert.deepEqual(merged[1].slice(15), existing[1].slice(15));
});

test('Master Data 输出表头缺失时拒绝写入', () => {
  assert.throws(
    () => _pme_validateMasterDataHeaders(['序列号/ Serial No.']),
    /Master Data/,
  );
});

test('Master Data 允许 P:R 技术员字段并在更新时原样保留', () => {
  const headers = [
    '序列号/ Serial No.', '工序/ KeyProcess', '保养日期/ PM_Date', '机台号/ Workcenter',
    'Production_AvailableTime_BEFORE PM', 'Production_AvailableTime_AFTER PM',
    'AU_BEFORE PM', 'AU_AFTER PM', 'UPDT_BEFORE PM', 'UPDT_AFTER PM', 'UPDT Gap',
    '保养月份/ YearMonth', '子工序/ SubProcess', '生产状态/ Production Status',
    '绩效状态/ Performance Status', '根本原因/  Root cause',
    '预防措施/Preventive action', '状态/Status',
  ];
  const oldRow = [
    '2026-01-01M1', 'IM', '2026-01-01', 'M1', 8, 150, 0.5, 0.6, 0.2, 0.3, 0.1,
    '202601', '', '生产', '差/Bad', 'CO导致UPDT偏高', '增加停机确认', '已完成',
  ];
  const newRow = [
    '2026-01-02M2', 'PK', '2026-01-02', 'M2', 8, 150, 0.5, 0.5, 0.2, 0.2, 0,
    '202601', '', '生产', '好/Good',
  ];
  const calculated = [[
    '2026-01-01M1', 'IM', '2026-01-01', 'M1', 16, 160, 0.4, 0.4, 0.1, 0.1, 0,
    '202601', '', '生产', '好/Good',
  ]];

  assert.equal(_pme_validateMasterDataHeaders(headers), true);
  const merged = _pme_upsertMasterData([headers, oldRow, newRow], calculated);

  assert.equal(merged[0].length, headers.length);
  assert.deepEqual(merged[1].slice(15), oldRow.slice(15));
  assert.deepEqual(merged[2].slice(15), newRow.slice(15));
});

test('Master Data 新增记录的技术员字段为空，写入只覆盖 A:O', () => {
  const headers = _pme_MASTER_HEADERS.concat([
    '根本原因/  Root cause', '预防措施/Preventive action', '状态/Status',
  ]);
  const calculated = [[
    '2026-01-03M3', 'TF', '2026-01-03', 'M3', 20, 160, 0.4, 0.4, 0.1, 0.1, 0,
    '202601', 'AFT & USW', '生产', '好/Good',
  ]];
  const merged = _pme_upsertMasterData([headers], calculated);

  assert.deepEqual(merged[1].slice(15), ['', '', '']);

  const writes = [];
  const formats = [];
  const sheet = {
    getRange(row, column, numRows, numColumns) {
      return {
        setValues(values) { writes.push({ row, column, numRows, numColumns, values }); },
        setNumberFormat(format) { formats.push({ row, column, numRows, numColumns, format }); },
      };
    },
  };
  _pme_writeMasterDataRows(sheet, merged);

  assert.deepEqual(writes[0], {
    row: 1,
    column: 1,
    numRows: 2,
    numColumns: 15,
    values: [headers.slice(0, 15), calculated[0]],
  });
  assert.deepEqual(formats, [
    { row: 2, column: 5, numRows: 1, numColumns: 2, format: '0' },
    { row: 2, column: 11, numRows: 1, numColumns: 1, format: '0.00%' },
  ]);
});

test('Master Data 清理 2026 失效记录后，物理删除末尾多余行', () => {
  const writes = [];
  const deletes = [];
  const sheet = {
    getLastRow() { return 3; },
    deleteRows(row, howMany) { deletes.push({ row, howMany }); },
    getRange(row, column, numRows, numColumns) {
      return {
        setValues(values) { writes.push({ row, column, numRows, numColumns, values }); },
        setNumberFormat() {},
      };
    },
  };

  _pme_writeMasterDataRows(sheet, [_pme_MASTER_HEADERS, ['2026-01-01M1']]);

  assert.deepEqual(deletes, [{ row: 3, howMany: 1 }]);
  assert.equal(writes[0].numRows, 2);
});

function masterRow(fields) {
  const row = new Array(15).fill('');
  row[0] = fields.serialNo || '';
  row[1] = fields.keyProcess || '';
  row[3] = fields.workcenter || '';
  row[11] = fields.yearMonth || '';
  row[12] = fields.subProcess || '';
  row[13] = fields.productionStatus || '';
  row[14] = fields.performanceStatus || '';
  return row;
}

test('KPI 汇总：只统计生产且 Good/Bad 的行，并同步统计 TF 子工序', () => {
  const rows = [
    masterRow({ serialNo: 'IM-G', keyProcess: 'IM', yearMonth: '202601', productionStatus: '生产', performanceStatus: '好/Good' }),
    masterRow({ serialNo: 'IM-B', keyProcess: 'IM', yearMonth: '202601', productionStatus: '生产', performanceStatus: '差/Bad' }),
    masterRow({ serialNo: 'TF-A', keyProcess: 'TF', subProcess: 'AFT & USW', yearMonth: '202601', productionStatus: '生产', performanceStatus: '好/Good' }),
    masterRow({ serialNo: 'TF-H', keyProcess: 'TF', subProcess: 'HSL & Bou', yearMonth: '202601', productionStatus: '生产', performanceStatus: '差/Bad' }),
    masterRow({ serialNo: 'TF-Z', keyProcess: 'TF', subProcess: 'Zar', yearMonth: '202601', productionStatus: '生产', performanceStatus: '好/Good' }),
    masterRow({ serialNo: 'PK-G', keyProcess: 'PK', yearMonth: '202601', productionStatus: '生产', performanceStatus: '好/Good' }),
    masterRow({ serialNo: 'NO-144', keyProcess: 'IM', yearMonth: '202601', productionStatus: '未达144', performanceStatus: '' }),
  ];

  const summary = _pme_calculateKpiSummary(rows, ['202601'], new Set(['TF-Z']));

  assert.deepEqual(summary.IM['202601'], { good: 1, bad: 1, total: 2 });
  assert.deepEqual(summary.TF['202601'], { good: 1, bad: 1, total: 2 });
  assert.deepEqual(summary.PK['202601'], { good: 1, bad: 0, total: 1 });
  assert.deepEqual(summary['AFT & USW']['202601'], { good: 1, bad: 0, total: 1 });
  assert.deepEqual(summary['HSL & Bou']['202601'], { good: 0, bad: 1, total: 1 });
  assert.deepEqual(summary.Zar['202601'], { good: 0, bad: 0, total: 0 });
});

test('KPI Summary 写入：使用既有 2026 行列、Goal 和 YTD，并清理无数据月份', () => {
  const writes = [];
  const backgrounds = [];
  const goalByRow = { 10: '60%', 11: '50%', 12: '60%', 13: '60%', 14: '60%', 15: '60%' };
  const sheet = {
    getRange(row, column) {
      return {
        getValue() { return column === 3 ? goalByRow[row] || '' : ''; },
        setValue(value) { writes.push({ row, column, value }); },
        setBackground(value) { backgrounds.push({ row, column, value }); },
      };
    },
  };
  const summary = _pme_calculateKpiSummary([
    masterRow({ serialNo: 'IM-G', keyProcess: 'IM', yearMonth: '202601', productionStatus: '生产', performanceStatus: '好/Good' }),
    masterRow({ serialNo: 'IM-B', keyProcess: 'IM', yearMonth: '202601', productionStatus: '生产', performanceStatus: '差/Bad' }),
  ], ['202601', '202602']);

  _pme_writeKpiSummary(summary, ['202601', '202602'], sheet);

  assert.deepEqual(writes.find(item => item.row === 10 && item.column === 4), { row: 10, column: 4, value: '50%' });
  assert.deepEqual(writes.find(item => item.row === 10 && item.column === 5), { row: 10, column: 5, value: '-' });
  assert.deepEqual(writes.find(item => item.row === 10 && item.column === 16), { row: 10, column: 16, value: '50%' });
  assert.ok(backgrounds.some(item => item.row === 10 && item.column === 4 && item.value === '#FFB6C1'));
  assert.ok(backgrounds.some(item => item.row === 10 && item.column === 5 && item.value === null));
});

test('主入口：完成 AEM 读取、Master Data 写入和 KPI Summary 汇总并返回执行摘要', () => {
  const masterWrites = [];
  const kpiWrites = [];
  const logs = [];

  function memorySheet(name, rows, writes) {
    return {
      getLastRow() { return rows.length; },
      getLastColumn() { return rows[0] ? rows[0].length : 0; },
      getRange(row, column, numRows = 1, numColumns = 1) {
        return {
          getValues() {
            return rows.slice(row - 1, row - 1 + numRows)
              .map(item => item.slice(column - 1, column - 1 + numColumns));
          },
          getValue() {
            if (name === 'KPI Summary' && column === 3) return '60%';
            return rows[row - 1] && rows[row - 1][column - 1];
          },
          setValues(values) {
            writes.push({ name, row, column, values });
          },
          setValue(value) {
            writes.push({ name, row, column, value });
          },
          setNumberFormat() {},
          setBackground(value) {
            writes.push({ name, row, column, background: value });
          },
        };
      },
    };
  }

  const pmRows = [
    ['类别', '状态', '开始日期', '开始时间', 'AEM#', '保养时间', '工序'],
    ['PM', '已确认', '2026-01-01', '08:00', 'M1', 8, 'INJ'],
  ];
  const aemRows = [
    AEM_HEADERS,
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', Shift: 1, Prod_Hours: 8, NetRuntime: 4, PlantAvailableTime: 8, PM: 0, process: 'IM' }),
    aemRow({ ProdDate: '2026-01-01', Workcenter: 'M1', Shift: 2, Prod_Hours: 0, NetRuntime: 0, PlantAvailableTime: 0, PM: 1, process: 'IM' }),
    aemRow({ ProdDate: '2026-01-02', Workcenter: 'M1', Shift: 1, Prod_Hours: 48, NetRuntime: 24, PlantAvailableTime: 48, PM: 0, process: 'IM' }),
    aemRow({ ProdDate: '2026-01-02', Workcenter: 'M1', Shift: 2, Prod_Hours: 48, NetRuntime: 24, PlantAvailableTime: 48, PM: 0, process: 'IM' }),
    aemRow({ ProdDate: '2026-01-02', Workcenter: 'M1', Shift: 3, Prod_Hours: 48, NetRuntime: 24, PlantAvailableTime: 48, PM: 0, process: 'IM' }),
  ];
  const mainSheets = new Map();
  const mainSpreadsheet = {
    getSheetByName(name) { return mainSheets.get(name) || null; },
  };
  mainSheets.set('PM Data', memorySheet('PM Data', pmRows, []));
  mainSheets.set('Master Data', memorySheet('Master Data', [_pme_MASTER_HEADERS], masterWrites));
  mainSheets.set('KPI Summary', memorySheet('KPI Summary', [new Array(16).fill('')], kpiWrites));
  const aemSpreadsheet = { getSheetByName() { return memorySheet('SourceData', aemRows, []); } };

  const previousSpreadsheetApp = globalThis.SpreadsheetApp;
  const previousWriteLog = globalThis.writeLog;
  const previousSources = _pme_CONFIG.aemSources;
  _pme_CONFIG.aemSources = [previousSources[4]];
  globalThis.SpreadsheetApp = {
    openById(id) { return id === _pme_CONFIG.mainSpreadsheetId ? mainSpreadsheet : aemSpreadsheet; },
  };
  globalThis.writeLog = (...args) => logs.push(args);

  try {
    const result = executePMECalculation({ triggerType: '手动', targetYear: 2026, months: ['202601'] });
    assert.deepEqual(result, {
      trigger: '手动',
      planCount: 1,
      calculatedCount: 1,
      warningCount: 0,
      months: ['202601'],
    });
    assert.ok(masterWrites.some(item => item.name === 'Master Data' && item.values.length === 2));
    assert.ok(kpiWrites.some(item => item.name === 'KPI Summary' && item.row === 10 && item.column === 4));
    const successLog = logs.find(args => args[0] === 'executePMECalculation' && args[1] === '成功' && args[3] === '手动');
    assert.ok(successLog);
    assert.match(successLog[2], /执行时长: \d+ ms \(\d+\.\d{2} s\)/);
  } finally {
    _pme_CONFIG.aemSources = previousSources;
    globalThis.SpreadsheetApp = previousSpreadsheetApp;
    globalThis.writeLog = previousWriteLog;
  }
});
