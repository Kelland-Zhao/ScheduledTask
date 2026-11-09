// PME（35 - PME计算.js）— AEM 来源区间解析测试
// 覆盖：季度归档 + 当前月双来源、月末/跨年滚动、来源区间不重叠
// 运行：node --test tests/pme-sources.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// 将 GAS 脚本加载进全局作用域（顶层仅 var + 函数声明，无 GAS 调用）
const code = fs.readFileSync(new URL('../35 - PME计算.js', import.meta.url), 'utf8');
(0, eval)(code);

const Q4_ARCHIVE_ID = '1oqAUPihd5NzksKMigORCeYkns_7-QNIqBhCp7nuHPB4';
const CURRENT_MONTH_ID = '1pKLCZfh8sKv6EcWuEsxH6jAIsjbWHU5_aAusEODnJ_o';

function resolved(asOf) {
  return _pme_resolveAemSources(asOf);
}

function byId(sources, id) {
  return sources.find(function (source) { return source.id === id; });
}

test('2026-11-09：Q4 归档覆盖 10 月，当前月覆盖 11 月', () => {
  const sources = resolved('2026-11-09');
  const q4 = byId(sources, Q4_ARCHIVE_ID);
  const current = byId(sources, CURRENT_MONTH_ID);

  assert.ok(q4, '应包含 2026Q4 归档');
  assert.equal(q4.start, '2026-10-01');
  assert.equal(q4.end, '2026-10-31');
  assert.ok(current, '应包含当前月来源');
  assert.equal(current.start, '2026-11-01');
  assert.equal(current.end, '');
});

test('2026-10-15：Q4 归档尚无已完结月份，应跳过', () => {
  const sources = resolved('2026-10-15');

  assert.equal(byId(sources, Q4_ARCHIVE_ID), undefined);
  assert.equal(byId(sources, CURRENT_MONTH_ID).start, '2026-10-01');
});

test('2026-12-15：Q4 归档扩展到 11 月，当前月覆盖 12 月', () => {
  const sources = resolved('2026-12-15');
  const q4 = byId(sources, Q4_ARCHIVE_ID);
  const current = byId(sources, CURRENT_MONTH_ID);

  assert.equal(q4.end, '2026-11-30');
  assert.equal(current.start, '2026-12-01');
});

test('2027-01-05：Q4 归档截止 12-31，当前月进入 1 月', () => {
  const sources = resolved('2027-01-05');
  const q4 = byId(sources, Q4_ARCHIVE_ID);
  const current = byId(sources, CURRENT_MONTH_ID);

  assert.equal(q4.end, '2026-12-31');
  assert.equal(current.start, '2027-01-01');
});

test('2027-02-05：Q4 归档结束日期封顶在季度末 12-31', () => {
  const sources = resolved('2027-02-05');

  assert.equal(byId(sources, Q4_ARCHIVE_ID).end, '2026-12-31');
});

test('解析后的来源区间通过覆盖校验（10月/11月/12月/次年1月）', () => {
  ['2026-10-15', '2026-11-09', '2026-12-15', '2027-01-05'].forEach((asOf) => {
    assert.doesNotThrow(() => _pme_validateSourceCoverage(resolved(asOf)), asOf);
  });
});

// ===== _pme_readAemSources 接线测试 =====
const FULL_HEADERS = [
  'Period', 'FocusFactory', 'OutputPcs', 'Week', 'ProdDate', 'SKU_No', 'Workcenter', 'Manhour',
  'SizeType', 'TypeDesc', 'TotalHC', 'StdHC', 'TrainHC', 'CreatedOn', 'Shift', 'GroupType', 'team',
  'description', 'Prod_Hours', 'output_cs', 'headcount', 'NetRuntime', 'PlantAvailableTime',
  'DisplayinPPR', 'process', 'HourlyStdOutput', 'RR_SizeType', 'HourlyStdOutput_RR', 'StdHC_RR',
  'StdMH_RR', 'ShiftStdOutput_MH', 'StdManhour_PR', 'Products', 'MH_Outsourcing', 'CO', 'MDT', 'PM',
  'No_Material', 'ActHC', 'Deduct_MH_Outsourcing', 'StdType', 'StdManhour', 'ProcessSeq',
  'Calendartime', 'WorkcenterGroup', 'MaterialType', '916_FULL_Outsourcing', 'Total_Scrap_Pcs',
  'Total_Scrap_Kg', 'Block', 'UnitWeight',
];

const DATE_BY_SOURCE = {
  '1QFvpIVWS3yv11Ujoh56cYsrBffM-zaH6Bc3h8sPCdJA': '2025-02-03',
  '1E-Rp9v_083LXJ5DbCW1jh9Wzuh4p5gPgQP2uOn4BRvg': '2025-05-05',
  '1TSthgE2c4RiZYsw6MiMkPy535dXWNsJgj6zicg4-OZ0': '2025-08-04',
  '1yCtPUp84XcaWxVfwIusgGg_LIB9RgtjZ_kmpQ1AhSXs': '2025-11-03',
  '1_vWImIpjTvUfDVY-s9Ij5Ajg5h-ayQ8XUbGcN5sEO68': '2026-02-02',
  '1yTkY1yz4zwiXCzSI8DiHSxWYu4CmBkzsQcKH1_2_Fv8': '2026-05-04',
  '10o2yG9Fzvlo54e0l71KeAM-QobNvUVC3xPcBZHIinC4': '2026-08-03',
  [Q4_ARCHIVE_ID]: '2026-10-15',
  [CURRENT_MONTH_ID]: '2026-11-05',
};

function makeAemRow(prodDate) {
  const row = new Array(FULL_HEADERS.length).fill('');
  row[FULL_HEADERS.indexOf('ProdDate')] = prodDate;
  row[FULL_HEADERS.indexOf('Workcenter')] = 'T1AFT308';
  row[FULL_HEADERS.indexOf('Shift')] = '1';
  row[FULL_HEADERS.indexOf('Prod_Hours')] = '8';
  row[FULL_HEADERS.indexOf('NetRuntime')] = '7';
  row[FULL_HEADERS.indexOf('PlantAvailableTime')] = '8';
  row[FULL_HEADERS.indexOf('process')] = 'TF';
  row[FULL_HEADERS.indexOf('CO')] = '1';
  row[FULL_HEADERS.indexOf('PM')] = '4';
  row[FULL_HEADERS.indexOf('WorkcenterGroup')] = 'AFT';
  return row;
}

function makeSheet(rows) {
  return {
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0].length,
    getRange: (startRow, startColumn, numRows, numColumns) => ({
      getValues: () => {
        const values = [];
        for (let i = 0; i < numRows; i++) {
          const source = rows[startRow - 1 + i] || [];
          const line = [];
          for (let j = 0; j < numColumns; j++) line.push(source[startColumn - 1 + j] ?? '');
          values.push(line);
        }
        return values;
      },
    }),
  };
}

test('2026-11-09：按 Q4 归档 + 当前月两个来源读取，并返回覆盖区间', () => {
  const opened = [];
  globalThis.SpreadsheetApp = {
    openById: (id) => {
      opened.push(id);
      const rows = [FULL_HEADERS.slice(), makeAemRow(DATE_BY_SOURCE[id])];
      return { getSheetByName: (name) => (name === 'SourceData' ? makeSheet(rows) : null) };
    },
  };

  const result = _pme_readAemSources(new Set(['T1AFT308']), '2026-11-09');

  assert.ok(opened.includes(Q4_ARCHIVE_ID), '应读取 2026Q4 归档');
  assert.ok(opened.includes(CURRENT_MONTH_ID), '应读取当前月快照');
  assert.ok(
    result.coverage.some((range) => range.start === '2026-10-01' && range.end === '2026-10-31'),
    '覆盖区间应包含 10 月',
  );
  assert.ok(
    result.coverage.some((range) => range.start === '2026-11-01' && range.end === ''),
    '覆盖区间应包含 11 月（开放结束）',
  );
});

// ===== 覆盖度防御：来源没覆盖的月份不写「未匹配」 =====
const PLAN = {
  serialNo: '2026-10-19T1AFT308',
  workcenter: 'T1AFT308',
  pmDate: '2026-10-19',
  keyProcess: 'TF',
  startTime: '08:00',
  durationHours: 2.5,
};

function calculateWithCoverage(coverage) {
  const source = {
    pmPlans: [Object.assign({}, PLAN)],
    aemRows: [['ProdDate', 'Workcenter', 'process', 'PM']],
    aemHeaderIndex: { ProdDate: 0, Workcenter: 1, process: 2, PM: 3 },
    exclusionRules: [],
    tfTypeMap: {},
  };
  if (coverage !== undefined) source.aemCoverage = coverage;
  return _pme_calculateMasterRows(source);
}

test('计划日期在来源覆盖范围内：记录为未匹配', () => {
  const result = calculateWithCoverage([{ start: '2026-10-01', end: '2026-10-31' }]);
  assert.deepEqual(result.unmatchedPlanSerials, [PLAN.serialNo]);
});

test('计划日期不在来源覆盖范围内：不记录为未匹配，避免清空已有指标', () => {
  const result = calculateWithCoverage([{ start: '2026-11-01', end: '' }]);
  assert.deepEqual(result.unmatchedPlanSerials, []);
});

test('旧调用方未提供覆盖区间：保持原有行为', () => {
  const result = calculateWithCoverage(undefined);
  assert.deepEqual(result.unmatchedPlanSerials, [PLAN.serialNo]);
});
