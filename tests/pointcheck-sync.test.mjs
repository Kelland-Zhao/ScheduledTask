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
