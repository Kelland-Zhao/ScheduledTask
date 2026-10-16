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
