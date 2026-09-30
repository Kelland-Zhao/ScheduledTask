// Workcenter 数据同步（12 - Workcenter数据同步.js）— Node 内置 test runner 测试
// 覆盖：按 Workcenter 键的行级同步（就地更新 / 追加 / 删除）、Active Cell → E–I 同步、安全阀、表头名解析
// 运行：node --test tests/workcenter-sync.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// 将 GAS 脚本加载进全局作用域（顶层仅 const + 函数声明，无 GAS 调用）
const code = fs.readFileSync(new URL('../12 - Workcenter数据同步.js', import.meta.url), 'utf8');
// 只 eval 一次：模块顶层的 const 在全局词法作用域里重复声明会报错
(0, eval)(code);

const SS_PLAN = '11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc';
const SS_EQU = '12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM';

// ===== GAS stub =====
let logs = [];
let sheets = {};

globalThis.writeLog = (funcName, status, detail, trigger, remark) => {
  logs.push({ funcName, status, detail, trigger, remark });
};

let consoleLines = [];
globalThis.console = {
  log: (...a) => consoleLines.push(a.join(' ')),
  warn: (...a) => consoleLines.push(a.join(' ')),
  error: (...a) => consoleLines.push(a.join(' ')),
};

// 真 Sheets 的 getLastRow 会忽略全空行，fake 保持一致
function lastNonEmptyRow(rows) {
  for (let i = rows.length - 1; i >= 0; i--) {
    if ((rows[i] || []).some(c => c !== '' && c !== undefined && c !== null)) return i + 1;
  }
  return 0;
}

function fakeSheet(rows) {
  return {
    getLastRow: () => lastNonEmptyRow(rows),
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    getRange: (r, c, nr, nc) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = rows[r - 1 + i] || [];
          const line = [];
          for (let j = 0; j < nc; j++) line.push(row[c - 1 + j] ?? '');
          out.push(line);
        }
        return out;
      },
      setValues: vals => {
        for (let i = 0; i < vals.length; i++) {
          const ri = r - 1 + i;
          while (rows.length <= ri) rows.push([]);
          for (let j = 0; j < vals[i].length; j++) {
            const ci = c - 1 + j;
            while (rows[ri].length <= ci) rows[ri].push('');
            rows[ri][ci] = vals[i][j];
          }
        }
      },
      clearContent: () => {
        for (let i = 0; i < nr; i++) {
          const row = rows[r - 1 + i];
          if (!row) continue;
          for (let j = 0; j < nc; j++) row[c - 1 + j] = '';
        }
      },
    }),
  };
}

globalThis.SpreadsheetApp = {
  openById: id => ({
    getSheetByName: name => sheets[id]?.[name] ?? null,
  }),
};

// ===== 夹具 =====
const WC_HEADERS = [
  'Workcenter', 'Machine Type', '机器性能', 'New Formed Cell',
  'HIM/Auto', 'VIM-1', 'VIM-2', 'VIM-3', 'VIM-4',
  'Final Machine Type', '是否主设备', '设备编号',
  '机型', '设备类型1', '设备类型2', '自动化类型',
  '责任人', '备份责任人', '无需检查Y/N',
];

const LINE_HEADERS = ['Count', 'Individual Machine', 'Machine Type', '机器性能', 'New Formed Cell', 'Remark'];

const ACTIVE_HEADERS = [
  'Count', 'WS', 'Process', 'New Formed Cell', 'Machine Group', 'Machine Group Description',
  'Len<=30', 'Cell Type', 'Cell Hot/Cold', 'Cell Ratio', 'Sort String', 'Dedicate Bundle',
  'HIM/Auto', 'VIM-1', 'VIM-2', 'VIM-3', 'VIM-4',
  'Sum of VIM Hot/Cold', 'VIM-1 Hot/Cold', 'VIM-2 Hot/Cold', 'VIM-3 Hot/Cold', 'VIM-4 Hot/Cold', 'Remark',
];
const ACTIVE_GROUP_ROW = [
  'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell', 'Cell',
  'Line', 'Line', 'Line', 'Line', 'Line', 'Line', 'Line', 'Line', 'Line', 'Line', '',
];

const ACTIVE_INDEX = Object.fromEntries(ACTIVE_HEADERS.map((h, i) => [h, i]));
const LINE_INDEX = Object.fromEntries(LINE_HEADERS.map((h, i) => [h, i]));

// 一行 Workcenter 数据（19 列），未指定的列留空
function wcRow(fields) {
  const row = new Array(19).fill('');
  Object.entries(fields).forEach(([name, val]) => {
    row[WC_HEADERS.indexOf(name)] = val;
  });
  return row;
}

// 一行 Line Database 数据
function lineRow(machine, machineType, performance, nfc) {
  return [1, machine, machineType, performance, nfc, ''];
}

// 一行 Active Cell 数据，cells 形如 { 'HIM/Auto': 'X', 'VIM-1': 'Y' }
function activeRow(nfc, cells) {
  const row = new Array(23).fill('');
  row[ACTIVE_INDEX['New Formed Cell']] = nfc;
  Object.entries(cells).forEach(([name, val]) => { row[ACTIVE_INDEX[name]] = val; });
  return row;
}

// EAM 表（9 列）：A 列「设备」= 设备编号，I 列「机台号 - Tag」= 匹配键
const EAM_HEADERS = ['设备', '描述', '成本中心', 'ABC 标识', '主工作中心', '系统状态', '创建日期', '更改人', '机台号 - Tag'];
function eamRow(equipmentNo, tag) {
  const row = new Array(9).fill('');
  row[0] = equipmentNo;
  row[8] = tag;
  return row;
}

function setup({ workcenter = [], line = [], activeCell = null, eam = [], workcenterHeaders = WC_HEADERS } = {}) {
  logs = [];
  consoleLines = [];
  sheets = {};
  sheets[SS_PLAN] = {
    '1. Line Database': fakeSheet([LINE_HEADERS, ...line]),
  };
  if (activeCell !== null) {
    sheets[SS_PLAN]['2. Active Cell'] = fakeSheet([ACTIVE_GROUP_ROW, ACTIVE_HEADERS, ...activeCell]);
  }
  sheets[SS_EQU] = {
    'Workcenter': fakeSheet([workcenterHeaders, ...workcenter]),
    'Equipment_Number_EAM': fakeSheet([EAM_HEADERS, ...eam]),
  };
}

function workcenterSheet() {
  return sheets[SS_EQU]['Workcenter'];
}

function rowsOf(sheet) {
  return sheet.getDataRange().getValues();
}

// 数据区（去掉表头行与尾部全空行），等价于真表里"看得见的那些机台"
function dataRows(sheet) {
  return rowsOf(sheet).slice(1).filter(r => r.some(c => c !== '' && c !== undefined && c !== null));
}

function keysOf(sheet) {
  return dataRows(sheet).map(r => r[0]);
}

// ===== Cycle 1：已有机台就地更新 =====
test('已有机台：更新 12 个程序列，人工列 M–S 原样保留', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'M1')],
    eam: [eamRow('EQ-001', 'M1')],
    activeCell: [cellRow('M1', 'H1', 'V1', 'V2', 'V3', 'V4')],
    workcenter: [
      wcRow({
        'Workcenter': 'M1', 'Machine Type': '旧机型', '机器性能': '旧性能', 'New Formed Cell': '旧NFC',
        'HIM/Auto': '旧HIM', 'VIM-1': '旧V1',
        'Final Machine Type': '旧Final', '是否主设备': 'N', '设备编号': '旧编号',
        '机型': '机型X', '设备类型1': '类型1X', '设备类型2': '类型2X', '自动化类型': '自动X',
        '责任人': '甲', '备份责任人': '乙', '无需检查Y/N': 'N',
      }),
    ],
  });

  globalThis.syncWorkcenterData();

  const [, row] = rowsOf(workcenterSheet());
  assert.equal(row[0], 'M1', 'A Workcenter');
  assert.equal(row[1], 'HT160', 'B Machine Type');
  assert.equal(row[2], '', 'C 机器性能');
  assert.equal(row[3], 'M1', 'D New Formed Cell 取自源');
  assert.equal(row[9], 'HT160', 'J Final Machine Type（机器性能为空时取 Machine Type）');
  assert.equal(row[11], 'EQ-001', 'L 设备编号取 EAM 映射');
});

test('已有机台：人工列 M–S 一个都不动', () => {
  setup({
    line: [lineRow('M1', 'HT160', '正常', 'C1')],
    workcenter: [
      wcRow({
        'Workcenter': 'M1', 'Machine Type': '旧机型', '机器性能': '旧性能', 'New Formed Cell': '旧NFC',
        '机型': '机型X', '设备类型1': '类型1X', '设备类型2': '类型2X', '自动化类型': '自动X',
        '责任人': '甲', '备份责任人': '乙', '无需检查Y/N': 'N',
      }),
    ],
  });

  globalThis.syncWorkcenterData();

  const [, row] = rowsOf(workcenterSheet());
  assert.equal(row[12], '机型X', 'M 机型');
  assert.equal(row[13], '类型1X', 'N 设备类型1');
  assert.equal(row[14], '类型2X', 'O 设备类型2');
  assert.equal(row[15], '自动X', 'P 自动化类型');
  assert.equal(row[16], '甲', 'Q 责任人');
  assert.equal(row[17], '乙', 'R 备份责任人');
  assert.equal(row[18], 'N', 'S 无需检查Y/N');
});

// ===== Cycle 2：追加与删除 =====
test('新机台：追加到末尾，只填程序列，人工列留空', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M2', 'FT400', '', 'C2')],
    activeCell: [cellRow('M1', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '机型': '机型X', '责任人': '甲' })],
  });

  globalThis.syncWorkcenterData();

  assert.deepEqual(keysOf(workcenterSheet()), ['M1', 'M2'], 'M2 追加在末尾');

  const m2 = dataRows(workcenterSheet())[1];
  assert.equal(m2[1], 'FT400', 'B Machine Type');
  assert.equal(m2[3], 'C2', 'D New Formed Cell');
  assert.equal(m2[9], 'FT400', 'J Final Machine Type');
  assert.equal(m2[10], 'N', 'K M2 不在 Active Cell 里 → N');

  for (let i = 12; i <= 18; i++) {
    assert.equal(m2[i], '', '人工列 ' + WC_HEADERS[i] + ' 应为空');
  }
});

test('源里消失的机台：整行删除（含它的人工列）', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M3', 'HS', '', 'C3')],
    workcenter: [
      wcRow({ 'Workcenter': 'M1', '责任人': '甲' }),
      wcRow({ 'Workcenter': 'M2', '机型': '机型Y', '责任人': '乙' }),
      wcRow({ 'Workcenter': 'M3', '责任人': '丙' }),
    ],
  });

  globalThis.syncWorkcenterData();

  assert.deepEqual(keysOf(workcenterSheet()), ['M1', 'M3'], 'M2 整行删除，M3 顶上');
  assert.equal(dataRows(workcenterSheet()).length, 2, '表尾不留空行');
  assert.equal(dataRows(workcenterSheet())[0][16], '甲', 'M1 人工列不受影响');
  assert.equal(dataRows(workcenterSheet())[1][16], '丙', 'M3 的人工列跟着行一起上移');
});

// ===== Cycle 3：2. Active Cell → E–I =====
function cellRow(nfc, him, v1, v2, v3, v4) {
  return activeRow(nfc, { 'HIM/Auto': him, 'VIM-1': v1, 'VIM-2': v2, 'VIM-3': v3, 'VIM-4': v4 });
}

test('E–I：按 New Formed Cell 匹配，从 Active Cell 的 M–Q 填入（表头有两行）', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    activeCell: [cellRow('C1', 'H1', 'V1', 'V2', 'V3', 'V4')],
    workcenter: [wcRow({ 'Workcenter': 'M1', 'New Formed Cell': 'C1', 'HIM/Auto': '旧', 'VIM-1': '旧' })],
  });

  globalThis.syncWorkcenterData();

  const row = dataRows(workcenterSheet())[0];
  assert.deepEqual(row.slice(4, 9), ['H1', 'V1', 'V2', 'V3', 'V4'], 'E–I');
});

test('E–I：Active Cell 里匹配不上 → 清空', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    activeCell: [cellRow('OTHER', 'H9', 'V9', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', 'New Formed Cell': 'C1', 'HIM/Auto': '旧H', 'VIM-1': '旧V' })],
  });

  globalThis.syncWorkcenterData();

  const row = dataRows(workcenterSheet())[0];
  assert.deepEqual(row.slice(4, 9), ['', '', '', '', ''], 'E–I 被清空');
});

test('E–I：Active Cell 里 New Formed Cell 重复 → 取第一条并告警', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    activeCell: [cellRow('C1', '第一条', '', '', '', ''), cellRow('C1', '第二条', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', 'New Formed Cell': 'C1' })],
  });

  globalThis.syncWorkcenterData();

  assert.equal(dataRows(workcenterSheet())[0][4], '第一条', '取第一条');
  assert.ok(consoleLines.some(l => l.includes('重复')), '有重复告警');
});

test('安全阀：Active Cell 缺失 → 跳过 E–I 同步，保留原值', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1', 'New Formed Cell': 'C1', 'HIM/Auto': '旧H', 'VIM-1': '旧V' })],
  });

  globalThis.syncWorkcenterData();

  const row = dataRows(workcenterSheet())[0];
  assert.equal(row[4], '旧H', 'HIM/Auto 保留');
  assert.equal(row[5], '旧V', 'VIM-1 保留');
});

test('E–I：机台 New Formed Cell 为空（闲置）→ 不参与匹配，保留原值', () => {
  setup({
    line: [lineRow('M1', 'HT160', '闲置', '')],
    activeCell: [cellRow('C1', 'H1', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', 'HIM/Auto': '旧H' })],
  });

  globalThis.syncWorkcenterData();

  assert.equal(dataRows(workcenterSheet())[0][4], '旧H', 'HIM/Auto 保留');
});

// ===== Cycle 4：安全阀与表头校验 =====
test('安全阀：源表读不到任何机台 → 中止，不改动 Workcenter', () => {
  setup({
    line: [],
    workcenter: [wcRow({ 'Workcenter': 'M1', '责任人': '甲' }), wcRow({ 'Workcenter': 'M2', '责任人': '乙' })],
  });

  globalThis.syncWorkcenterData();

  assert.deepEqual(keysOf(workcenterSheet()), ['M1', 'M2'], '一行都没动');
  assert.ok(logs.some(l => l.status === '跳过'), '记了跳过日志');
});

test('安全阀：源表机台数不足表内一半 → 中止，不改动 Workcenter', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    workcenter: [
      wcRow({ 'Workcenter': 'M1' }), wcRow({ 'Workcenter': 'M2' }),
      wcRow({ 'Workcenter': 'M3' }), wcRow({ 'Workcenter': 'M4' }),
    ],
  });

  globalThis.syncWorkcenterData();

  assert.deepEqual(keysOf(workcenterSheet()), ['M1', 'M2', 'M3', 'M4'], '一行都没删');
  assert.ok(logs.some(l => l.status === '跳过'), '记了跳过日志');
});

test('安全阀边界：源表机台数恰好是表内一半 → 放行', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M2', 'DP', '', 'C2')],
    workcenter: [
      wcRow({ 'Workcenter': 'M1' }), wcRow({ 'Workcenter': 'M2' }),
      wcRow({ 'Workcenter': 'M3' }), wcRow({ 'Workcenter': 'M4' }),
    ],
  });

  globalThis.syncWorkcenterData();

  assert.deepEqual(keysOf(workcenterSheet()), ['M1', 'M2'], '正常执行删除');
  assert.ok(!logs.some(l => l.status === '跳过'), '没有跳过');
});

// ===== Cycle 5：【是否主设备】改为按 Active Cell 成员资格判定 =====
test('是否主设备：机台号出现在 Active Cell 的 D 列 → Y', () => {
  setup({
    line: [lineRow('M1', 'FT400', '', 'C1')],
    activeCell: [cellRow('M1', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '是否主设备': 'N' })],
  });

  globalThis.syncWorkcenterData();

  assert.equal(dataRows(workcenterSheet())[0][10], 'Y', 'K');
});

test('是否主设备：机台号不在 Active Cell 里 → N（旧值 Y 也改写）', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '是否主设备': 'Y' })],
  });

  globalThis.syncWorkcenterData();

  assert.equal(dataRows(workcenterSheet())[0][10], 'N', 'K');
});

test('是否主设备：按 A 列 Workcenter 匹配，不是按 D 列 New Formed Cell', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],           // 机台号 M1、机组号 C1，两者不同
    activeCell: [cellRow('C1', 'H1', '', '', '', '')],  // Active Cell 里只有机组号 C1
    workcenter: [wcRow({ 'Workcenter': 'M1', 'New Formed Cell': 'C1', '是否主设备': 'Y' })],
  });

  globalThis.syncWorkcenterData();

  const row = dataRows(workcenterSheet())[0];
  assert.equal(row[10], 'N', 'K 按 A 列匹配 → M1 不在 → N');
  assert.equal(row[4], 'H1', 'E–I 仍按 D 列匹配 → 命中 C1');
});

test('是否主设备：Active Cell 不可用 → 保留原值；此时新增机台 K 留空', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M2', 'DP', '', 'C2')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '是否主设备': 'Y' })],
  });

  globalThis.syncWorkcenterData();

  const rows = dataRows(workcenterSheet());
  assert.equal(rows[0][10], 'Y', '已有行保留原值');
  assert.equal(rows[1][10], '', '新增行无原值可保留');
});

test('表头缺少必需字段 → 中止并记失败日志，不改动数据', () => {
  const broken = WC_HEADERS.filter(h => h !== '设备编号');
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '责任人': '甲' })],
    workcenterHeaders: broken,
  });

  globalThis.syncWorkcenterData();

  assert.equal(dataRows(workcenterSheet())[0][16], '甲', '数据未被改动');
  assert.ok(logs.some(l => l.status === '失败' && l.detail.includes('设备编号')), '失败日志点名字段');
});
