// 模块 13 点检机台核对 — 读取 Workcenter 计划账的列定位测试
// 覆盖：11 列 → 19 列新结构下按表头名取列、闲置过滤、表头缺失时的保护
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

test('19 列新结构：machineModel 取 J 列 Final Machine Type', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', 'Machine Type': 'HT160', 'Final Machine Type': '6AX', 'New Formed Cell': 'NFC1' }),
  ];

  const result = globalThis._pc_buildWorkcenterMap(data, '闲置');

  assert.equal(result.map['M1'].machineType, 'HT160', 'B 列 Machine Type');
  assert.equal(result.map['M1'].machineModel, '6AX', 'J 列 Final Machine Type');
  assert.deepEqual(result.missing, []);
});

test('machineModel 不再是 D 列（New Formed Cell）', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', 'New Formed Cell': '6AX', 'Final Machine Type': 'DP' }),
  ];

  const result = globalThis._pc_buildWorkcenterMap(data, '闲置');

  assert.equal(result.map['M1'].machineModel, 'DP', '按 D 列误读的话这里会是 6AX');
});

test('机器性能为关键词（闲置）的行被排除', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', '机器性能': '闲置' }),
    row({ 'Workcenter': 'M2', '机器性能': '' }),
  ];

  const result = globalThis._pc_buildWorkcenterMap(data, '闲置');

  assert.deepEqual(Object.keys(result.map), ['M2']);
});

test('表头缺少必需字段 → 返回空并报出缺失字段', () => {
  const broken = WC_HEADERS.filter(h => h !== 'Final Machine Type');
  const data = [
    broken,
    row({ 'Workcenter': 'M1' }),
  ];

  const result = globalThis._pc_buildWorkcenterMap(data, '闲置');

  assert.deepEqual(Object.keys(result.map), [], '宁可空手而归，也不能读错列');
  assert.deepEqual(result.missing, ['Final Machine Type']);
});
