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
const SS_USER = '1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM';

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

// ===== 邮件与日期桩（Cycle 9 起同步末尾会发「机台主数据维护提醒」）=====
// 模块依赖 01 - Common.js 的 buildHtmlTable/escapeHtml/formatVariableAsDate，
// 但 Common.js 顶层会 SpreadsheetApp.openById()，在 Node 里跑不了，按需打桩。
let mails = [];
globalThis.GmailApp = {
  sendEmail: (to, subject, body, options) => { mails.push({ to, subject, body, options }); },
};
globalThis.escapeHtml = str => String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// 与 Common.js 同构的简化桩：内容经 escapeHtml；断言只做子串匹配，不做像素级 UI 校验
globalThis.buildHtmlTable = function (headers, rows, headerBg) {
  return '<table data-bg="' + headerBg + '"><tr>'
    + headers.map(h => '<th>' + globalThis.escapeHtml(h) + '</th>').join('')
    + '</tr>'
    + rows.map(r => '<tr>' + r.map(c => '<td>' + globalThis.escapeHtml(String(c === undefined || c === null || c === '' ? '-' : c)) + '</td>').join('') + '</tr>').join('')
    + '</table>';
};
globalThis.formatVariableAsDate = () => '2026-10-02';   // 固定日期，便于断言

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
// 20 列：S = 程序派生的「工艺无需检查Y/N」，T = 人工维护的「点检无需检查Y/N」（模块 12 不碰）
const WC_HEADERS = [
  'Workcenter', 'Machine Type', '机器性能', 'New Formed Cell',
  'HIM/Auto', 'VIM-1', 'VIM-2', 'VIM-3', 'VIM-4',
  'Final Machine Type', '是否主设备', '设备编号',
  '机型', '设备类型1', '设备类型2', '自动化类型',
  '责任人', '备份责任人', '工艺无需检查Y/N', '点检无需检查Y/N',
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

// 一行 Workcenter 数据（20 列），未指定的列留空
function wcRow(fields) {
  const row = new Array(WC_HEADERS.length).fill('');
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

function setup({ workcenter = [], line = [], activeCell = null, eam = [], workcenterHeaders = WC_HEADERS, userID = null } = {}) {
  logs = [];
  consoleLines = [];
  mails = [];
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
  if (userID !== null) {
    sheets[SS_USER] = { 'userID': fakeSheet(userID) };
  }
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
test('已有机台：更新程序列，人工列 M–R 原样保留', () => {
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
        '责任人': '甲', '备份责任人': '乙', '工艺无需检查Y/N': 'N',
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

test('已有机台：人工列 M–R 一个都不动', () => {
  setup({
    line: [lineRow('M1', 'HT160', '正常', 'C1')],
    workcenter: [
      wcRow({
        'Workcenter': 'M1', 'Machine Type': '旧机型', '机器性能': '旧性能', 'New Formed Cell': '旧NFC',
        '机型': '机型X', '设备类型1': '类型1X', '设备类型2': '类型2X', '自动化类型': '自动X',
        '责任人': '甲', '备份责任人': '乙',
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

  for (let i = 12; i <= 17; i++) {
    assert.equal(m2[i], '', '人工列 ' + WC_HEADERS[i] + ' 应为空');
  }
  assert.equal(m2[18], '', 'S 非免检 → 留空');
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

// ===== Cycle 7：【工艺无需检查Y/N】自动判定 =====
function sOf(setupOpts) {
  setup(setupOpts);
  globalThis.syncWorkcenterData();
  return dataRows(workcenterSheet())[0][18];
}

test('工艺无需检查：J=6AX 且 K=N → Y', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'FT400', '6AX', 'C1')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],  // M1 不在 Active Cell → K=N
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'Y');
});

test('工艺无需检查：J=6AX 但 K=Y（是主设备）→ 留空', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'FT400', '6AX', 'M1')],
    activeCell: [cellRow('M1', '', '', '', '', '')],    // M1 在 Active Cell → K=Y
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), '');
});

test('工艺无需检查：J=NA（报废/闲置）→ Y', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'FT400', '报废', 'C1')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'Y');
});

test('工艺无需检查：其他机型 → 留空', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), '');
});

// 第三条判据：D=H2HTB363 且 K=N —— 绑在 D 列单元格上的一次性豁免
test('工艺无需检查：D=H2HTB363 且 K=N → Y（一次性单元格豁免）', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', 'H2HTB363')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],  // M1 不在 Active Cell → K=N
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'Y');
});

test('工艺无需检查：D=H2HTB363 但 K=Y（是主设备）→ 留空（豁免只给非主设备）', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', 'H2HTB363')],
    activeCell: [cellRow('M1', '', '', '', '', '')],    // M1 在 Active Cell → K=Y
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), '');
});

test('工艺无需检查：D=H2HTB363X（近似值）且 K=N → 留空（只认精确匹配）', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', 'H2HTB363X')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), '');
});

test('工艺无需检查：D 为 H2HTB363 且带首尾空格、K=N → Y（trim 后精确匹配）', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', '  H2HTB363  ')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'Y');
});

test('工艺无需检查：S 列归程序管，人工改的值会被重算覆盖', () => {
  assert.equal(sOf({
    line: [lineRow('M1', 'HT160', '', 'C1')],           // J=HT160 → 非免检
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({ 'Workcenter': 'M1', '工艺无需检查Y/N': 'Y' })],  // 人工写着 Y
  }), '', '重算后应为空');
});

// ===== Cycle 8：【点检无需检查Y/N】人工列不得被同步覆盖 =====
test('人工列保护：点检无需检查Y/N 的值在同步后原样保留（模块 12 绝不写它）', () => {
  // 模块 12 的写回是「整行重写」：基础行取自表内该行，再逐列覆盖受管列。
  // 点检无需检查Y/N 不在 _WS_MANAGED_HEADERS 里，靠基础行整行带过来 —— 这是设计假设，
  // 不是保证：若它被加进受管列（或按第 20 列硬写），人类的标记每天 08:20 会被静默抹掉。本用例守住这一点。
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M2', 'DP', '', 'C2')],
    activeCell: [cellRow('C1', 'H1', '', '', '', '')],
    workcenter: [
      wcRow({ 'Workcenter': 'M1', '点检无需检查Y/N': 'Y', '责任人': '甲' }),
      // M2 在源里、表里没有 → 走追加分支；新增行没有人工值可带，必须留空
    ],
  });

  globalThis.syncWorkcenterData();

  const col = WC_HEADERS.indexOf('点检无需检查Y/N');
  const rows = dataRows(workcenterSheet());

  assert.equal(rowsOf(workcenterSheet())[0][col], '点检无需检查Y/N', '表头行本身也不得被动过');
  assert.deepEqual(rows.map(r => r[0]), ['M1', 'M2'], '同步确实执行了（M2 被追加），断言不是空跑');
  assert.equal(rows[0][1], 'HT160', '受管列确实被写入');
  assert.equal(rows[0][col], 'Y', 'M1 的人工标记必须原样保留');
  assert.equal(rows[1][col], '', 'M2 是新增行，无人工值 → 留空');
});

// ===== Cycle 6：Final Machine Type 写入规则 =====
function jOf(setupOpts) {
  setup(setupOpts);
  globalThis.syncWorkcenterData();
  return dataRows(workcenterSheet())[0][9];
}

test('Final Machine Type：机器性能为「报废」→ NA', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'FT400', '报废', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'NA');
});

test('Final Machine Type：机器性能为「闲置」→ NA', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'FT400', '闲置', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'NA');
});

test('Final Machine Type：机器性能含混合状态文本（机台号报废,原665机器闲置）→ NA', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'HT160 G', '机台号报废,原665机器闲置', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'NA');
});

test('Final Machine Type：F350 → FCS（不再要求后跟括号数字）', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'F350', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'FCS');
});

test('Final Machine Type：F600(3) → FCS（原有行为不变）', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'F600(3)', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'FCS');
});

test('Final Machine Type：FT400 保持原样，不转 FCS', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'FT400', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'FT400');
});

test('Final Machine Type：E110 → ENG，HT160 → 原样', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'E110', '', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), 'ENG');
  assert.equal(jOf({
    line: [lineRow('M2', 'HT160', '', 'C2')],
    workcenter: [wcRow({ 'Workcenter': 'M2' })],
  }), 'HT160');
});

test('Final Machine Type：机器性能非状态时仍优先于 Machine Type', () => {
  assert.equal(jOf({
    line: [lineRow('M1', 'HT160 3X', '3AX', 'C1')],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
  }), '3AX');
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

// ===== Cycle 9：机台主数据维护提醒（M–R 人工列未补齐）=====
// 判据：J≠NA 且以下六个人工列任一为空；S 是程序列、T 空是「需点检」的合法默认，都不参与判据
const MAINT_HEADERS = ['机型', '设备类型1', '设备类型2', '自动化类型', '责任人', '备份责任人'];

// ---- userID 表夹具：前两行为分组/字段表头（取样 userID!A1:R2），数据行取样自 2026-10-01 生产表 ----
const USER_GROUP_ROW = ['SAPID328*****', '', '', '备件系统', '', 'Digital Process', '', '', 'AM Schedule', '', 'PK BOM', '', '', 'EDS', 'EDS', 'EDS', 'Quality', 'F'];
const USER_HEADERS = ['SAPID', 'NAME', 'PWD', 'Approver-Process', 'Access', '车间', '工序', '权限', 'AM Schedule', 'GMail', 'BOM 上传', '技术部审批失效BOM', '审批类型', '车间', '工序', '职位', 'ROLE'];
// 只填被读取的列（A/B/J/N/O/P），其余列与本用例无关
function userRow(sapid, name, gmail, dept, proc, pos) {
  const row = new Array(17).fill('');
  row[0] = sapid; row[1] = name; row[9] = gmail; row[13] = dept; row[14] = proc; row[15] = pos;
  return row;
}
const USER_FIXTURE = [
  USER_GROUP_ROW,
  USER_HEADERS,
  userRow('66185', '卢少辉', 'andy_lu@colpal.com', 'TB2', 'INJ', 'S&C'),          // 取样 userID!A178
  userRow('68680', '季勇', 'yong_ji@colpal.com', 'TB1', 'INJ', 'S&C'),            // 取样 userID!A184
  userRow('69071', '沈毅', 'frank_shen_shen@colpal.com', 'ALL', 'INJ', 'S&C'),    // 取样 userID!A188
  userRow('38416', '张俊', 'jin_zhang@colpal.com', 'TB1', 'INJ', 'IDL'),          // 取样 userID!A3
  userRow('62564', '吴江峰', 'jiangfeng_wu@colpal.com', 'TB1', 'INJ', 'IDL'),     // 取样 userID!A66
  userRow('68697', '刘元冬', 'yuandong_liu@colpal.com', 'TB2', 'TF', 'IDL'),      // 取样 userID!A73（工序 TF → 不进名单）
  userRow('66284', '毛太林', '', 'TB1', 'INJ', 'IDL'),                            // 取样 userID!A39（无 GMail → 跳过）
];

// ---- 判据纯函数 _ws_findMaintenancePending ----
test('维护提醒判据：新增机台（M–R 全空、J≠NA）→ 命中并缺六列', () => {
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [wcRow({ 'Workcenter': 'M9', 'Machine Type': 'FT400', 'Final Machine Type': 'FT400', '是否主设备': 'N' })];
  assert.deepEqual(
    globalThis._ws_findMaintenancePending(cols, matrix, new Set(['M9'])),
    [{ workcenter: 'M9', missing: MAINT_HEADERS, isNew: true }]);
});

test('维护提醒判据：J=NA（报废/闲置）→ 不进名单', () => {
  // 取样 Workcenter：V1FTA553（J=NA、机型组空）；H1HTB658（机器性能=机器闲置，机台号报废，已维护）
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [
    wcRow({ 'Workcenter': 'V1FTA553', 'Final Machine Type': 'NA', '工艺无需检查Y/N': 'Y', '点检无需检查Y/N': 'Y' }),
    wcRow({
      'Workcenter': 'H1HTB658', 'Final Machine Type': 'NA', '机型': '6AX', '设备类型1': 'VIM',
      '设备类型2': 'NA', '自动化类型': 'NA', '工艺无需检查Y/N': 'Y', '点检无需检查Y/N': 'Y',
    }),
  ];
  assert.deepEqual(globalThis._ws_findMaintenancePending(cols, matrix, new Set()), []);
});

test('维护提醒判据：缺责任人/备份责任人（V1FTA701 真实行）→ 命中两列', () => {
  // 取样 Workcenter：V1FTA701（J=6AX、K=N、机型 VIM/VIM/NA/NA、S=T=Y，责任人列为空）
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [wcRow({
    'Workcenter': 'V1FTA701', 'Machine Type': 'FT400', '机器性能': '6AX', 'New Formed Cell': 'H1HTA755',
    'Final Machine Type': '6AX', '是否主设备': 'N', '设备编号': '10139954',
    '机型': 'VIM', '设备类型1': 'VIM', '设备类型2': 'NA', '自动化类型': 'NA',
    '工艺无需检查Y/N': 'Y', '点检无需检查Y/N': 'Y',
  })];
  assert.deepEqual(
    globalThis._ws_findMaintenancePending(cols, matrix, new Set()),
    [{ workcenter: 'V1FTA701', missing: ['责任人', '备份责任人'], isNew: false }]);
});

test('维护提醒判据：缺机型组（H2HTA362 真实行）→ 命中四列', () => {
  // 取样 Workcenter：H2HTA362（J=3AX、K=Y，机型组四列为空、责任人已维护）
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [wcRow({
    'Workcenter': 'H2HTA362', 'Machine Type': 'HT160 3X', '机器性能': '3AX', 'New Formed Cell': 'H2HTA362',
    'Final Machine Type': '3AX', '是否主设备': 'Y', '设备编号': '10132392',
    '责任人': '束伍魁', '备份责任人': '乔巨欢',
  })];
  assert.deepEqual(
    globalThis._ws_findMaintenancePending(cols, matrix, new Set()),
    [{ workcenter: 'H2HTA362', missing: ['机型', '设备类型1', '设备类型2', '自动化类型'], isNew: false }]);
});

test('维护提醒判据：六列齐全 → 不命中；S/T 的值与判据无关', () => {
  // 取样 Workcenter：S1HS0001（六列齐全，S 空、T=Y）；第二条 T=Y 且 S=Y（双免检）同样六列齐全
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [
    wcRow({
      'Workcenter': 'S1HS0001', 'Final Machine Type': 'HS', '机型': 'HS', '设备类型1': 'HS',
      '设备类型2': 'NA', '自动化类型': 'NA', '责任人': '游臣', '备份责任人': '吴江峰', '点检无需检查Y/N': 'Y',
    }),
    wcRow({
      'Workcenter': 'S2HS0003', 'Final Machine Type': 'HS', '机型': 'HS', '设备类型1': 'HS',
      '设备类型2': 'NA', '自动化类型': 'NA', '责任人': '杨杨', '备份责任人': '乔军',
      '工艺无需检查Y/N': 'Y', '点检无需检查Y/N': 'Y',
    }),
  ];
  assert.deepEqual(globalThis._ws_findMaintenancePending(cols, matrix, new Set()), []);
});

test('维护提醒判据：空白串（含空格）视为未维护', () => {
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [wcRow({
    'Workcenter': 'M1', 'Final Machine Type': 'HT160', '机型': '  ',
    '设备类型1': 'HS', '设备类型2': 'NA', '自动化类型': 'NA', '责任人': '甲', '备份责任人': '乙',
  })];
  assert.deepEqual(
    globalThis._ws_findMaintenancePending(cols, matrix, new Set()),
    [{ workcenter: 'M1', missing: ['机型'], isNew: false }]);
});

test('维护提醒判据：重复机台号只报第一条', () => {
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [
    wcRow({ 'Workcenter': 'M1', 'Final Machine Type': 'HT160' }),   // 全空 → 命中
    wcRow({ 'Workcenter': 'M1', 'Final Machine Type': 'HT160', '责任人': '甲' }),
  ];
  const pending = globalThis._ws_findMaintenancePending(cols, matrix, new Set());
  assert.equal(pending.length, 1, '同号只报一次');
  assert.equal(pending[0].workcenter, 'M1');
});

test('维护提醒判据：addedSet 决定 isNew 标记', () => {
  const cols = globalThis._ws_headerIndex(WC_HEADERS);
  const matrix = [wcRow({ 'Workcenter': 'M1', 'Final Machine Type': 'HT160' })];
  assert.equal(globalThis._ws_findMaintenancePending(cols, matrix, new Set(['M1']))[0].isNew, true);
  assert.equal(globalThis._ws_findMaintenancePending(cols, matrix, new Set(['OTHER']))[0].isNew, false);
});

// ---- 收件人纯函数 _ws_parseMaintenanceRecipients ----
test('维护提醒收件人：INJ S&C → TO；INJ IDL → CC；非 INJ 与无邮箱跳过', () => {
  const r = globalThis._ws_parseMaintenanceRecipients(USER_FIXTURE);
  assert.deepEqual(r.to, ['andy_lu@colpal.com', 'yong_ji@colpal.com', 'frank_shen_shen@colpal.com']);
  assert.deepEqual(r.cc, ['jin_zhang@colpal.com', 'jiangfeng_wu@colpal.com']);
});

test('维护提醒收件人：职位大小写与空格归一（S&C/s&c 同效；工序 trim 后精确匹配 INJ）', () => {
  const rows = [
    USER_GROUP_ROW, USER_HEADERS,
    userRow('66185', '卢少辉', 'andy_lu@colpal.com', 'TB2', 'INJ', ' s&c '),   // 职位变体（大小写/空格）
    userRow('38416', '张俊', 'jin_zhang@colpal.com', 'TB1', ' INJ ', 'idl'),   // 工序带空格 + 职位小写
  ];
  const r = globalThis._ws_parseMaintenanceRecipients(rows);
  assert.deepEqual(r.to, ['andy_lu@colpal.com']);
  assert.deepEqual(r.cc, ['jin_zhang@colpal.com']);
});

test('维护提醒收件人：重复邮箱去重、空表返回空名单', () => {
  const rows = [
    USER_GROUP_ROW, USER_HEADERS,
    userRow('66185', '卢少辉', 'andy_lu@colpal.com', 'TB2', 'INJ', 'S&C'),
    userRow('66185', '卢少辉', 'andy_lu@colpal.com', 'TB2', 'INJ', 'S&C'),     // 同人重复行
  ];
  assert.deepEqual(globalThis._ws_parseMaintenanceRecipients(rows).to, ['andy_lu@colpal.com']);
  assert.deepEqual(globalThis._ws_parseMaintenanceRecipients([USER_GROUP_ROW, USER_HEADERS]), { to: [], cc: [] });
});

// ---- 邮件 HTML 纯函数 _ws_buildMaintenanceEmailHtml ----
const PENDING_FIXTURE = [
  { workcenter: 'V1FTA701', missing: ['责任人', '备份责任人'], isNew: true },
  { workcenter: 'H2HTA362', missing: ['机型', '设备类型1', '设备类型2', '自动化类型'], isNew: false },
];

test('维护提醒邮件：含机台号、缺失字段、新增/待补齐标记、车间与两个链接', () => {
  const html = globalThis._ws_buildMaintenanceEmailHtml(PENDING_FIXTURE, '2026-10-02');
  assert.ok(html.includes('V1FTA701'));
  assert.ok(html.includes('H2HTA362'));
  assert.ok(html.includes('责任人、备份责任人'));
  assert.ok(html.includes('机型、设备类型1、设备类型2、自动化类型'));
  assert.ok(html.includes('本次新增'));
  assert.ok(html.includes('待补齐'));
  assert.ok(html.includes('2 台'));
  assert.ok(html.includes('TB1'));
  assert.ok(html.includes('TB2'));
  assert.ok(html.includes('2026-10-02'));
  // 维护入口必须先落 EDS 登录页（登录后由 next 回跳维护页）；不得再出现未登录可直达维护页的链接
  assert.ok(html.includes('https://script.google.com/a/colpal.com/macros/s/AKfycbyaQjG5yFGYxU825DrODhSLl2bdfbYKpqAH4qOIzKoTJ4b-5qU/exec?v=home_new_1.0&next=INJ_MachineMaster'), 'EDS 登录页链接（带 next 回跳）');
  assert.ok(!html.includes('exec?v=INJ_MachineMaster'), '不得再出现直达维护页的链接');
  assert.ok(html.includes('https://docs.google.com/spreadsheets/d/12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM/edit#gid=0'), 'Workcenter 表链接');
});

test('维护提醒邮件：机台号经 HTML 转义', () => {
  const html = globalThis._ws_buildMaintenanceEmailHtml([{ workcenter: 'A<b>', missing: ['机型'], isNew: false }], '2026-10-02');
  assert.ok(html.includes('A&lt;b&gt;'));
  assert.ok(!html.includes('A<b>'));
});

// ---- 同步集成：发送时机与收件人 ----
function pendingSetup(extra = {}) {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1'), lineRow('M9', 'FT400', '', 'C9')],
    activeCell: [cellRow('OTHER', '', '', '', '', '')],
    workcenter: [wcRow({
      'Workcenter': 'M1', '机型': 'HS', '设备类型1': 'HS', '设备类型2': 'NA', '自动化类型': 'NA',
      '责任人': '游臣', '备份责任人': '吴江峰',
    })],
    userID: USER_FIXTURE,
    ...extra,
  });
}

test('维护提醒（定时）：发真实收件人 TO=S&C、CC=IDL，正文只列未补齐机台', () => {
  pendingSetup();

  globalThis.syncWorkcenterData({ triggerType: 'scheduled' });   // 有事件对象 = 定时

  assert.equal(mails.length, 1, '发了一封');
  const mail = mails[0];
  assert.equal(mail.to, 'andy_lu@colpal.com,yong_ji@colpal.com,frank_shen_shen@colpal.com');
  assert.equal(mail.options.cc, 'jin_zhang@colpal.com,jiangfeng_wu@colpal.com');
  assert.ok(mail.subject.includes('【机台主数据维护】'));
  assert.ok(mail.subject.includes('1 台'));
  assert.ok(mail.subject.includes('2026-10-02'));
  assert.ok(mail.options.htmlBody.includes('M9'), '新增机台在名单里');
  assert.ok(!mail.options.htmlBody.includes('M1'), 'M1 六列齐全，不进名单');
  assert.equal(mail.options.name, '机台主数据维护提醒');
});

test('维护提醒（手动）：只发操作者，不带 CC，防止调试时误发全员', () => {
  pendingSetup();

  globalThis.syncWorkcenterData();   // 无事件对象 = 手动

  assert.equal(mails.length, 1);
  assert.equal(mails[0].to, 'kelland_zhao@colpal.com');
  assert.equal(mails[0].options.cc, undefined);
});

test('维护提醒：无待维护机台 → 不发信', () => {
  setup({
    line: [lineRow('M1', 'HT160', '', 'C1')],
    workcenter: [wcRow({
      'Workcenter': 'M1', '机型': 'HS', '设备类型1': 'HS', '设备类型2': 'NA', '自动化类型': 'NA',
      '责任人': '游臣', '备份责任人': '吴江峰',
    })],
    userID: USER_FIXTURE,
  });

  globalThis.syncWorkcenterData({ triggerType: 'scheduled' });

  assert.equal(mails.length, 0);
});

test('维护提醒（定时）：收件人读不到 → 不发信、有告警，同步照常成功', () => {
  setup({
    line: [lineRow('M9', 'FT400', '', 'C9')],
    workcenter: [],
  });   // 不注册 userID 表

  globalThis.syncWorkcenterData({ triggerType: 'scheduled' });

  assert.equal(mails.length, 0);
  assert.ok(consoleLines.some(l => l.includes('收件人')), '有收件人告警');
  assert.ok(logs.some(l => l.status === '成功'), '同步仍成功');
});

test('维护提醒：同步被安全阀中止 → 不发信', () => {
  setup({
    line: [],
    workcenter: [wcRow({ 'Workcenter': 'M1' })],
    userID: USER_FIXTURE,
  });

  globalThis.syncWorkcenterData({ triggerType: 'scheduled' });

  assert.equal(mails.length, 0);
});

test('维护提醒：发信抛错不影响同步结果，摘要里注明失败', () => {
  setup({
    line: [lineRow('M9', 'FT400', '', 'C9')],
    workcenter: [],
    userID: USER_FIXTURE,
  });
  const orig = globalThis.GmailApp.sendEmail;
  globalThis.GmailApp.sendEmail = () => { throw new Error('quota exceeded'); };
  try {
    globalThis.syncWorkcenterData({ triggerType: 'scheduled' });
  } finally {
    globalThis.GmailApp.sendEmail = orig;
  }

  assert.ok(logs.some(l => l.status === '成功' && l.detail.includes('维护提醒失败')), '主日志仍成功并注明提醒失败');
  assert.ok(consoleLines.some(l => l.includes('quota exceeded')));
});

test('维护提醒：手动测试入口只读生产表、只发操作者', () => {
  setup({
    workcenter: [
      wcRow({ 'Workcenter': 'V1FTA701', 'Final Machine Type': '6AX', '机型': 'VIM', '设备类型1': 'VIM', '设备类型2': 'NA', '自动化类型': 'NA' }),
    ],
  });

  globalThis.testWorkcenterMaintenanceReminder();

  assert.equal(mails.length, 1);
  assert.equal(mails[0].to, 'kelland_zhao@colpal.com');
  assert.ok(mails[0].subject.includes('测试'));
  assert.ok(mails[0].options.htmlBody.includes('V1FTA701'));
});
