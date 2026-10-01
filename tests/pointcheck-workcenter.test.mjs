// 模块 13 点检机台核对 — 读取 Workcenter 计划账的列定位测试
// 覆盖：11 列 → 20 列新结构下按表头名取列、纳入集判定（点检无需检查Y/N）、表头缺失时的保护
// 运行：node --test tests/pointcheck-workcenter.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const code = fs.readFileSync(new URL('../13 - 点检机台核对.js', import.meta.url), 'utf8');
(0, eval)(code);

// 20 列结构：S = 模块 12 派生的「工艺无需检查Y/N」，T = 人工维护的「点检无需检查Y/N」。
// 本模块只读后者（点检口径）
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

test('判据只看点检口径列：工艺标Y 不影响纳入，点检标Y 才排除', () => {
  // 两列名字相近、位置相邻，判据一旦串列（读成工艺列），
  // 点检范围的判定就会跟着工艺口径走 —— 本用例锁的就是这层区分
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', '工艺无需检查Y/N': 'Y', '点检无需检查Y/N': '' }),
    row({ 'Workcenter': 'M2', '工艺无需检查Y/N': '', '点检无需检查Y/N': 'Y' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['M1'], 'M1 工艺标Y、点检未标 → 仍纳入点检');
  assert.equal(r.flagged['M2'], true, 'M2 点检标Y → 排除');
});

test('只要求 Workcenter 与判据列：缺少 Machine Type / Final Machine Type 也能跑', () => {
  // 这两列在模块里没有任何读取方，不再作为必需表头；缺了不该报缺失、更不该整表放弃
  const lean = WC_HEADERS.filter(h => h !== 'Machine Type' && h !== 'Final Machine Type');
  const data = [lean, row({ 'Workcenter': 'M1' })];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(r.missing, [], '非判据列缺失不算缺字段');
  assert.deepEqual(Object.keys(r.included), ['M1']);
});

test('点检无需检查Y/N = Y 的行被排除出纳入集', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'H2FTA001', '点检无需检查Y/N': 'Y' }),
    row({ 'Workcenter': 'H2FTA002', '点检无需检查Y/N': '' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['H2FTA002']);
  assert.equal(r.flagged['H2FTA001'], true, '被排除的机台要能分辨是「因Y」');
});

test('点检无需检查Y/N 小写 y / 中文「是」不当作 Y，但计入告警', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'H2FTA001', '点检无需检查Y/N': 'y' }),
    row({ 'Workcenter': 'H2FTA002', '点检无需检查Y/N': '是' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included).sort(), ['H2FTA001', 'H2FTA002'], '不能被当成Y排除掉');
  assert.deepEqual(r.badFlags.sort(), ['y', '是'], '必须告警，不能静默失效');
});

test('表头缺少 点检无需检查Y/N → 返回空并报出缺失字段', () => {
  const broken = WC_HEADERS.filter(h => h !== '点检无需检查Y/N');
  const data = [broken, row({ 'Workcenter': 'H2FTA001' })];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), [], '宁可空手而归，也不能读错列');
  assert.deepEqual(r.missing, ['点检无需检查Y/N']);
});

test('Workcenter 重复机台号：取第一条并告警，只出现一次', () => {
  // 两行大小写不同、归一后同键；rawNo 保留各自原值，用它证明「取第一条」依然成立
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'h2fta001' }),
    row({ 'Workcenter': 'H2FTA001' }),
  ];

  const r = globalThis._pc_buildIncludedSet(data);

  assert.deepEqual(Object.keys(r.included), ['H2FTA001'], 'key 只有一个');
  assert.equal(r.included['H2FTA001'].rawNo, 'h2fta001', '取第一条（rawNo 为表11 原值）');
  assert.deepEqual(r.duplicates, ['H2FTA001']);
});

test('机台号归一：前后空格与大小写差异视为同一台', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': ' h2fta001 ', '点检无需检查Y/N': 'Y' }),
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
