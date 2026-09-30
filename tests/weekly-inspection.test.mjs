// 模块 10 注塑工艺周检 — 机台清单读取测试
// 覆盖：数据源指向 11 表、19 列结构下按表头名取字段、免检过滤、表头缺失保护
// 运行：node --test tests/weekly-inspection.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const code = fs.readFileSync(new URL('../10 - 注塑工艺周检.js', import.meta.url), 'utf8');
(0, eval)(code);

// 11 表 Workcenter 的 19 列结构
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

test('机台数据源：实际去开的是 11 表 Workcenter（03 表退役的前提）', () => {
  const opened = [];
  globalThis.SpreadsheetApp = {
    openById: id => ({
      getSheetByName: name => {
        opened.push({ id, name });
        return { getDataRange: () => ({ getValues: () => [WC_HEADERS] }) };
      },
    }),
  };

  globalThis._wiGetMachineData();

  assert.deepEqual(opened, [{
    id: '12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM',
    name: 'Workcenter',
  }]);
});

test('19 列结构下按表头名取到 Workcenter / 责任人 / 备份责任人', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', '责任人': '游臣', '备份责任人': '吴江峰' }),
  ];

  const result = globalThis._wi_filterMachines(data);

  assert.equal(result.length, 1);
  assert.equal(result[0]['Workcenter'], 'M1');
  assert.equal(result[0]['责任人'], '游臣');
  assert.equal(result[0]['备份责任人'], '吴江峰');
});

test('免检过滤：无需检查Y/N = Y 的机台被排除', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1', '无需检查Y/N': 'Y' }),
    row({ 'Workcenter': 'M2', '无需检查Y/N': '' }),
    row({ 'Workcenter': 'M3' }),
  ];

  const result = globalThis._wi_filterMachines(data);

  assert.deepEqual(result.map(r => r['Workcenter']), ['M2', 'M3'], 'M1 免检被排除，空值纳入');
});

test('责任人可以为空（源里没填的机台也要生成检查记录）', () => {
  const data = [
    WC_HEADERS,
    row({ 'Workcenter': 'M1' }),
  ];

  const result = globalThis._wi_filterMachines(data);

  assert.equal(result.length, 1, '空责任人不应导致机台被丢弃');
  assert.equal(result[0]['责任人'], '');
});

test('表头缺少必需字段 → 抛错，不静默读成 undefined', () => {
  const broken = WC_HEADERS.filter(h => h !== '责任人');
  const data = [
    broken,
    row({ 'Workcenter': 'M1' }),
  ];

  assert.throws(() => globalThis._wi_filterMachines(data), /责任人/);
});
