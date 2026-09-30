// 模块 13 点检机台核对 — 读取 Workcenter 计划账的列定位测试
// 覆盖：11 列 → 19 列新结构下按表头名取列、纳入集判定（无需检查Y/N）、表头缺失时的保护
// 运行：node --test tests/pointcheck-workcenter.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const code = fs.readFileSync(new URL('../13 - 点检机台核对.js', import.meta.url), 'utf8');
(0, eval)(code);

const WC_HEADERS = [
  'Workcenter', 'Machine Type', '机器性能', 'New Formed Cell',
  'HIM/Auto', 'VIM-1', 'VIM-2', 'VIM-3', 'VIM-4',
  'Final Machine Type', '是否主设备', '设备编号',
  '机型', '设备类型1', '设备类型2', '自动化类型',
  '责任人', '备份责任人', '无需检查Y/N',
];

function row(fields) {
  const r = new Array(WC_HEADERS.length).fill('');
  Object.entries(fields).forEach(([k, v]) => { r[WC_HEADERS.indexOf(k)] = v; });
  return r;
}

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
