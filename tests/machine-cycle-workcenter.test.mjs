// 模块 33 机台周期监控 — 从 Workcenter 筛选机台的列定位测试
// 覆盖：11 列 → 20 列新结构下按表头名取列、是否主设备/机型过滤、表头缺失时的保护
// 运行：node --test tests/machine-cycle-workcenter.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const code = fs.readFileSync(new URL('../33 - 机台周期监控.js', import.meta.url), 'utf8');
(0, eval)(code);

// 重构后的 20 列结构（S=工艺无需检查Y/N 派生列，T=点检无需检查Y/N 人工列；本模块两列都不读）
const WC_HEADERS = [
  'Workcenter', 'Machine Type', '机器性能', 'New Formed Cell',
  'HIM/Auto', 'VIM-1', 'VIM-2', 'VIM-3', 'VIM-4',
  'Final Machine Type', '是否主设备', '设备编号',
  '机型', '设备类型1', '设备类型2', '自动化类型',
  '责任人', '备份责任人', '工艺无需检查Y/N', '点检无需检查Y/N',
];

function row(fields) {
  const r = new Array(WC_HEADERS.length).fill('');
  Object.entries(fields).forEach(([k, v]) => { r[WC_HEADERS.indexOf(k)] = v; });
  return r;
}

const HEADER = WC_HEADERS;

test('19 列新结构：按表头名取列，筛出 是否主设备=Y 且 Final Machine Type=6AX 的机台', () => {
  const data = [
    HEADER,
    row({ 'Workcenter': 'M1', 'Machine Type': 'HT160', 'Final Machine Type': '6AX', '是否主设备': 'Y' }),
    row({ 'Workcenter': 'M2', 'Machine Type': 'HT160', 'Final Machine Type': '6AX', '是否主设备': 'N' }),
    row({ 'Workcenter': 'M3', 'Machine Type': 'HT160', 'Final Machine Type': '3AX', '是否主设备': 'Y' }),
  ];

  const result = globalThis._mc_selectMachines(data, '6AX');

  assert.deepEqual(Object.keys(result.machines), ['M1'], '只有 M1 同时满足两个条件');
  assert.deepEqual(result.missing, [], '表头齐全');
});

test('Final Machine Type 取的是 J 列，不是旧结构的 D 列', () => {
  const data = [
    HEADER,
    // D 列(New Formed Cell) = '6AX'，J 列(Final Machine Type) = 'DP' —— 旧代码会把 D 当成机型而误判
    row({ 'Workcenter': 'M1', 'New Formed Cell': '6AX', 'Final Machine Type': 'DP', '是否主设备': 'Y' }),
  ];

  const result = globalThis._mc_selectMachines(data, '6AX');

  assert.deepEqual(Object.keys(result.machines), [], '按 D 列误判的话这里会错误地筛出 M1');
});

test('是否主设备 取的是 K 列，不是旧结构的 E 列', () => {
  const data = [
    HEADER,
    // E 列(HIM/Auto) = 'Y'，K 列(是否主设备) = 'N'
    row({ 'Workcenter': 'M1', 'HIM/Auto': 'Y', 'Final Machine Type': '6AX', '是否主设备': 'N' }),
  ];

  const result = globalThis._mc_selectMachines(data, '6AX');

  assert.deepEqual(Object.keys(result.machines), [], '按 E 列误判的话这里会错误地筛出 M1');
});

test('机台号为空的行被跳过', () => {
  const data = [
    HEADER,
    row({ 'Workcenter': '', 'Final Machine Type': '6AX', '是否主设备': 'Y' }),
  ];

  const result = globalThis._mc_selectMachines(data, '6AX');

  assert.deepEqual(Object.keys(result.machines), []);
});

test('表头缺少必需字段 → 不返回任何机台，并报出缺失字段', () => {
  const broken = WC_HEADERS.filter(h => h !== '是否主设备');
  const data = [
    broken,
    row({ 'Workcenter': 'M1', 'Final Machine Type': '6AX' }),
  ];

  const result = globalThis._mc_selectMachines(data, '6AX');

  assert.deepEqual(Object.keys(result.machines), [], '宁可一台都不返回，也不能读错列');
  assert.deepEqual(result.missing, ['是否主设备']);
});
