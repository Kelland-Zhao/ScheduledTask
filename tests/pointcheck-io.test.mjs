// 模块 13「点检机台核对」IO 层测试 —— 假表驱动真实编排
// （_pc_run / _pc_executePlan / _pc_getRecipients / _pc_sendMail）
// 运行：node --test tests/pointcheck-io.test.mjs
//
// 模块本体（13 - 点检机台核对.js）一行不改。本文件在 Node 里搭一套**会真移位**的假
// SpreadsheetApp：deleteRows / insertRowsAfter 同时 splice 值数组与背景色数组，
// 行号像真表一样移动 —— 于是「删行后必须重新读背景色再涂」这类**顺序**约束才可能被证伪
// （涂早一步，黄就落到错误的行上，甚至落到 PK/TF 行）。
//
// 假表建模 / 未建模的语义，见文件末尾注释块。
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// ===== 模块依赖的公共函数打桩 =====
// 01 - Common.js 顶层会 SpreadsheetApp.openById()，Node 里加载不了，按需打桩（同 sync 测试）
globalThis.buildHtmlTable = function (headers, rows) {
  return '<table><tr>' + headers.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr>'
    + rows.map(function (r) {
        return '<tr>' + r.map(function (c) { return '<td>' + c + '</td>'; }).join('') + '</tr>';
      }).join('')
    + '</table>';
};
globalThis.formatVariableAsDate = function (d) { return String(d); };

// ===== 捕获器：writeLog（在 Common.js 里，Node 加载不了）/ GmailApp / console =====
let logs = [];
let mails = [];
let consoleLines = [];

globalThis.writeLog = function (funcName, status, detail, trigger, remark) {
  logs.push({ funcName, status, detail, trigger, remark });
};
globalThis.GmailApp = {
  sendEmail: function (to, subject, body, options) {
    mails.push({ to, subject, body, options });
  },
};
// 模块会大量 console.log；吞掉以免刷屏（node --test 的 reporter 不走 globalThis.console）
globalThis.console = {
  log: function () { consoleLines.push([].join.call(arguments, ' ')); },
  warn: function () { consoleLines.push([].join.call(arguments, ' ')); },
  error: function () { consoleLines.push([].join.call(arguments, ' ')); },
};

// 模块本体只 eval 一次（顶层 const 在全局词法作用域重复声明会报错）
const code = fs.readFileSync(new URL('../13 - 点检机台核对.js', import.meta.url), 'utf8');
(0, eval)(code);

// ===== 假电子表格 =====
// 与模块数据源常量一一对应（模块里的 const 不在 globalThis 上，这里按值复制）
const SS_ID_POINTCHECK = '1RQql-PrcBWiAQNeg7hQKcocpllSUMRhT5XPrDTVWoBY';
const SS_ID_PLAN = '12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM';
const SS_ID_PERMISSION = '1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM';

function blankRow(width) { return new Array(width).fill(''); }

/** 假表返回背景色的归一形态：小写十六进制；未设置/已清除 → '#ffffff'（与真表一致） */
function normalizeStoredBg(v) {
  if (v === undefined || v === null) return '#ffffff';
  const s = String(v).trim().toLowerCase();
  if (s === '' || s === '#ffffff' || s === 'white') return '#ffffff';
  return s;
}

/**
 * 假工作表：值 / 背景色两条平行数组，含表头行（下标 0）。
 * 建模语义（与真表一致）：
 *  - getLastRow() 忽略尾部全空行；
 *  - getRange 返回的窗口是**活的**：每次 getValues()/getBackgrounds() 都读当前数组；
 *  - deleteRows(start,count)：start 为 1-based，count 行从两条数组里一起移走，后续行上移；
 *  - insertRowsAfter(afterRow,n)：在 afterRow 之后插入 n 行空行（5 列、无色）；
 *  - setValues/setBackgrounds 的矩阵尺寸必须与区间一致（真表会抛错）；
 *  - 写入超出已存行数时视为写进了网格空白区：自动补空行。
 * 所有变更调用都记录进 sheet.calls（含参数与**调用前**的快照），供断言"当时覆盖了哪些行"。
 */
function makeSheet(name, rows, backgrounds) {
  const values = rows.map(function (r) { return r.slice(); });
  const bgs = backgrounds
    ? backgrounds.map(function (r) { return r.slice(); })
    : rows.map(function () { return []; });
  while (bgs.length < values.length) bgs.push([]);
  const calls = [];

  function lastRow() {
    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i] || [];
      if (row.some(function (v) { return v !== '' && v !== undefined && v !== null; })) return i + 1;
    }
    return 0;
  }

  function lastCol() {
    return values.reduce(function (m, r) { return Math.max(m, (r || []).length); }, 0);
  }

  function readValues(row, col, numRows, numCols) {
    const out = [];
    for (let i = 0; i < numRows; i++) {
      const src = values[row - 1 + i] || [];
      const line = [];
      for (let j = 0; j < numCols; j++) {
        const v = src[col - 1 + j];
        line.push(v === undefined || v === null ? '' : v);   // 真空单元格读回 ''（真表如此）
      }
      out.push(line);
    }
    return out;
  }

  function readBackgrounds(row, col, numRows, numCols) {
    const out = [];
    for (let i = 0; i < numRows; i++) {
      const src = bgs[row - 1 + i] || [];
      const line = [];
      for (let j = 0; j < numCols; j++) line.push(normalizeStoredBg(src[col - 1 + j]));
      out.push(line);
    }
    return out;
  }

  /** 1-based 行列；按需把两条数组补到该单元格（真表网格远比数据大） */
  function ensureCell(r, c) {
    while (values.length < r) { values.push(blankRow(5)); bgs.push(blankRow(5)); }
    const vr = values[r - 1];
    const br = bgs[r - 1];
    for (let k = vr.length; k < c; k++) vr.push('');
    for (let k = br.length; k < c; k++) br.push('');
  }

  function record(method, args, extra) {
    const entry = {
      method: method,
      args: args,
      // 调用前快照：断言"这次写入覆盖的行，当时是不是 INJ 行"必须用调用时的状态
      before: {
        values: values.map(function (r) { return r.slice(); }),
        backgrounds: bgs.map(function (r) { return r.slice(); }),
      },
    };
    if (extra) Object.assign(entry, extra);
    calls.push(entry);
    return entry;
  }

  function assertGrid(grid, numRows, numCols, what) {
    if (!Array.isArray(grid) || grid.length !== numRows
      || grid.some(function (r) { return !Array.isArray(r) || r.length !== numCols; })) {
      throw new Error(what + ' 矩阵尺寸与区间不符，期望 ' + numRows + 'x' + numCols);
    }
  }

  function range(row, col, numRows, numCols) {
    if (!(row >= 1 && col >= 1 && numRows >= 1 && numCols >= 1)) {
      throw new Error('getRange 参数非法: ' + [row, col, numRows, numCols].join(','));
    }
    return {
      getValues: function () { return readValues(row, col, numRows, numCols); },
      getBackgrounds: function () { return readBackgrounds(row, col, numRows, numCols); },
      setValues: function (grid) {
        assertGrid(grid, numRows, numCols, 'setValues');
        record('setValues', [row, col, numRows, numCols], { values: grid.map(function (r) { return r.slice(); }) });
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            ensureCell(row + i, col + j);
            values[row - 1 + i][col - 1 + j] = grid[i][j];
          }
        }
      },
      setBackgrounds: function (grid) {
        assertGrid(grid, numRows, numCols, 'setBackgrounds');
        record('setBackgrounds', [row, col, numRows, numCols], { colors: grid.map(function (r) { return r.slice(); }) });
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            ensureCell(row + i, col + j);
            bgs[row - 1 + i][col - 1 + j] = grid[i][j];
          }
        }
      },
      clearContent: function () {
        record('clearContent', [row, col, numRows, numCols]);
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            ensureCell(row + i, col + j);
            values[row - 1 + i][col - 1 + j] = '';
          }
        }
      },
    };
  }

  return {
    name: name,
    calls: calls,
    getLastRow: lastRow,
    getLastColumn: lastCol,
    getDataRange: function () {
      return range(1, 1, Math.max(lastRow(), 1), Math.max(lastCol(), 1));
    },
    getRange: range,
    deleteRows: function (start, count) {
      if (!Number.isInteger(start) || !Number.isInteger(count) || start < 1 || count < 1) {
        throw new Error('deleteRows 参数非法: ' + start + ',' + count);
      }
      record('deleteRows', [start, count]);
      // 真表网格有 1000 行：起点落在数据之外的空白行时，删除不改变任何数据
      if (start > values.length) return;
      values.splice(start - 1, count);
      bgs.splice(start - 1, count);
    },
    insertRowsAfter: function (afterRow, numRows) {
      if (!Number.isInteger(afterRow) || !Number.isInteger(numRows) || afterRow < 1 || numRows < 1) {
        throw new Error('insertRowsAfter 参数非法: ' + afterRow + ',' + numRows);
      }
      record('insertRowsAfter', [afterRow, numRows]);
      const at = Math.min(afterRow, values.length);   // 0-based 插入位 = afterRow
      for (let i = 0; i < numRows; i++) {
        values.splice(at, 0, blankRow(5));
        bgs.splice(at, 0, blankRow(5));
      }
    },
  };
}

// ===== 世界搭建 =====
/**
 * @param {Object} spec
 *   mlRows/mlBgs: MachineList 的行与背景（背景可省，默认全无色）
 *   wcRows: Workcenter 行（表头行必须自带，模块按表头名定位列）
 *   recipients: [{email, process?, role?}]，默认 process=INJ / role=S&C；
 *               _pc_getRecipients 从第 3 行起读，故前两行留白
 *   userIdRows: 直接给 userID 表的行（优先于 recipients）
 */
function setupWorld(spec) {
  const ml = makeSheet('MachineList', spec.mlRows, spec.mlBgs);
  const wc = makeSheet('Workcenter', spec.wcRows);
  const user = makeSheet('userID', spec.userIdRows || recipientRows(spec.recipients || []));

  const books = new Map([
    [SS_ID_POINTCHECK, new Map([['MachineList', ml]])],
    [SS_ID_PLAN, new Map([['Workcenter', wc]])],
    [SS_ID_PERMISSION, new Map([['userID', user]])],
  ]);
  globalThis.SpreadsheetApp = {
    openById: function (id) {
      const book = books.get(id);
      if (!book) throw new Error('假表没有这个 ID: ' + id);
      return { getSheetByName: function (n) { return book.get(n) || null; } };
    },
  };

  logs = [];
  mails = [];
  consoleLines = [];
  const sheets = [ml, wc, user];
  return {
    ml: ml,
    wc: wc,
    user: user,
    sheets: sheets,
    /** 三张表上的全部变更调用（读操作不记录） */
    mutations: function () {
      return sheets.reduce(function (acc, s) { return acc.concat(s.calls); }, []);
    },
    /** 清空调用记录，但不动表内容（幂等测试两轮之间用） */
    resetCalls: function () {
      sheets.forEach(function (s) { s.calls.length = 0; });
    },
  };
}

function recipientRows(entries) {
  // 第 1 行表头、第 2 行副表头 —— 模块从第 3 行开始读，这两行故意留白即可
  const rows = [blankRow(16), blankRow(16)];
  entries.forEach(function (e) {
    const row = blankRow(16);
    row[9] = e.email;                                   // J 列：GMail
    row[14] = e.process === undefined ? 'INJ' : e.process;   // O 列：工序
    row[15] = e.role === undefined ? 'S&C' : e.role;    // P 列：职位
    rows.push(row);
  });
  return rows;
}

const rec = function (email, process, role) { return { email: email, process: process, role: role }; };

// ===== 运行入口（走模块的真实入口，不直接调 _pc_run）=====
/** 手动：e === undefined → 绕过安全阀 */
function runManual() { globalThis.checkPointCheckMachines(); }
/** 定时：e 存在 → 启用安全阀 */
function runScheduled() { globalThis.checkPointCheckMachines({}); }
/** 预演：dryRunPointCheckSync */
function runDryRun() { globalThis.dryRunPointCheckSync(); }

/** 模块把一切异常都吞成「失败」Log —— 不守的话 fixture 写错会伪装成"零写入" */
function assertNoFailure() {
  const failed = logs.filter(function (l) { return l.status === '失败'; });
  assert.equal(failed.length, 0, '模块异常被吞成失败 Log: ' + JSON.stringify(failed));
}

// ===== 读表小工具 =====
function allRows(world) { return world.ml.getDataRange().getValues(); }
function rowsOf(world, process) {
  return allRows(world).filter(function (r) { return String(r[0] || '').trim() === process; });
}
/** 按 D 列机台号找 1-based 行号（全表找，不区分工序） */
function findRow(world, machineNo) {
  const rows = allRows(world);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3] || '').trim() === machineNo) return i + 1;
    // 注意：i 是 0-based，行号 = i + 1；上面 return 的是 1-based 行号
  }
  throw new Error('找不到机台号: ' + machineNo);
}
function bgAt(world, row, col) {
  return world.ml.getRange(row, 1, 1, 5).getBackgrounds()[0][col || 0];
}

/** 快照所有非 INJ 行（工序+机台号+出现序号 作身份键），用于逐格比对值与背景 */
function snapshotNonInj(sheet) {
  const values = sheet.getDataRange().getValues();
  const bgs = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 5).getBackgrounds();
  const map = new Map();
  const counts = new Map();
  values.forEach(function (row, i) {
    if (String(row[0] || '').trim() === 'INJ') return;
    const base = String(row[0] || '') + '|' + String(row[3] || '');
    const n = (counts.get(base) || 0) + 1;
    counts.set(base, n);
    map.set(base + '#' + n, {
      values: row.slice(0, 5),
      backgrounds: (bgs[i] || []).slice(0, 5),
    });
  });
  return map;
}

// ===== 夹具 =====
const ML_HEADERS = ['工序', '车间', '机型', '机台号', '点检人'];
const r5 = function (a, b, c, d, e) { return [a, b, c, d, e]; };
const W = ['', '', '', '', ''];
const YEL = ['#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00', '#FFFF00'];

// 生产形态：PK / TF / INJ 段交错；TF 段有 T1GR0301 两行，INJ 段也有一台同号机
// （生产表里 T1GR0301 就同时出现在 TF 与 INJ 段）
// 背景刻意留两处人工遗留的黄（PK、TF 各一）—— 本模块永远不许碰非 INJ 行
const ML_MAIN_ROWS = [
  ML_HEADERS,                                        // 1
  r5('PK',  'TB1', '6AX',  'H2PK0001', '张三'),      // 2  人工遗留黄，不可碰
  r5('TF',  'TB1', 'OMNI', 'T1GR0301', 'OP1'),       // 3
  r5('TF',  'TB1', 'OMNI', 'T1GR0301', 'OP2'),       // 4  同号两行
  r5('INJ', 'TB2', '6AX',  'D1DEL001', 'OPC'),       // 5  删（表11 标 Y）
  r5('INJ', 'TB2', '6AX',  'D2DEL002', 'OPC'),       // 6  删（表11 查不到）
  r5('INJ', 'TB2', '3AX',  'K1KEEP01', 'OPC'),       // 7  完整却留着黄 → 应清除
  r5('INJ', 'TB1', '',     'K2KEEP02', ''),          // 8  待补全 → 应标黄（删两行后上移到第 6 行）
  r5('INJ', 'TB2', '6AX',  'T1GR0301', 'OPC'),       // 9  与 TF 段同号，出现在 INJ 段
  r5('PK',  'TB2', '6AX',  'H2PK0002', '李四'),      // 10 涂错位时黄会落到这行
  r5('TF',  'TB2', 'OMNI', 'T2TF0001', 'OP5'),       // 11 人工遗留黄，不可碰
  r5('INJ', 'TB1', 'OMNI', 'K3KEEP03', 'OPC'),       // 12
];
const ML_MAIN_BGS = [
  W,     // 表头
  YEL,   // 2  PK 行人工遗留的黄
  W, W,
  W, W,  // 5-6 待删
  YEL,   // 7  K1KEEP01 完整但留黄 → 应清除
  W,     // 8  K2KEEP02 待补全
  W,     // 9
  W,     // 10 PK
  YEL,   // 11 TF 行人工遗留的黄
  W,     // 12
];

const WC_HEADERS = ['Workcenter', '点检无需检查Y/N'];
const wc = function (no, flag) { return [no, flag || '']; };

const WC_MAIN_ROWS = [
  WC_HEADERS,
  wc('K1KEEP01'),
  wc('K2KEEP02'),
  wc('T1GR0301'),
  wc('K3KEEP03'),
  wc('D1DEL001', 'Y'),     // 标 Y → FLAG 删除
  wc('E0EN0001'),          // 表11 有、MachineList 无 → 追加到表尾
];

// 删除区间夹具：A1ADEL01 一台两行（连续段），Z1ZDEL99 一台单行（断开段）
const ML_RUNS_ROWS = [
  ML_HEADERS,                                        // 1
  r5('INJ', 'TB1', '6AX', 'A1ADEL01', 'OPC'),        // 2  ┐ 一段（两行）
  r5('INJ', 'TB1', '6AX', 'A1ADEL01', 'OPC'),        // 3  ┘
  r5('INJ', 'TB1', '6AX', 'K1KEEP01', 'OPC'),        // 4  断开点
  r5('INJ', 'TB1', '6AX', 'Z1ZDEL99', 'OPC'),        // 5  ┐ 另一段（单行）
  r5('INJ', 'TB1', '6AX', 'K2KEEP02', 'OPC'),        // 6
];
const WC_RUNS_ROWS = [WC_HEADERS, wc('K1KEEP01'), wc('K2KEEP02')];

// 安全阀夹具：11 台待删（>10）＋ 1 台保留；表11 另有一台待追加
const ML_VALVE_ROWS = [ML_HEADERS]
  .concat(Array.from({ length: 11 }, function (_, i) {
    return r5('INJ', 'TB1', '6AX', 'X' + String(i + 1).padStart(2, '0') + 'DEL001', 'OPC');
  }))
  .concat([r5('INJ', 'TB1', '3AX', 'K1KEEP01', 'OPC')]);
const WC_VALVE_ROWS = [WC_HEADERS, wc('K1KEEP01'), wc('E0EN0001')];

// ===== 测试 =====

test('IO-1 一次真实同步只动 INJ 行：PK/TF 的值与背景逐格不变，删除只落在 INJ', () => {
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS, wcRows: WC_MAIN_ROWS,
    recipients: [rec('ops@example.com')],
  });
  const before = snapshotNonInj(world.ml);

  runManual();
  assertNoFailure();

  // 非 INJ 行（含表头、PK、TF）必须原样还在：身份键集合一致，逐格值与背景一致
  const after = snapshotNonInj(world.ml);
  assert.equal(after.size, before.size, 'PK/TF/表头 行数不得变化');
  before.forEach(function (snap, key) {
    assert.ok(after.has(key), '非 INJ 行丢失: ' + key);
    assert.deepEqual(after.get(key).values, snap.values, key + ' 的值被改动');
    assert.deepEqual(after.get(key).backgrounds, snap.backgrounds, key + ' 的背景被改动');
  });

  // 每条删除调用覆盖的行，在当时都必须是 INJ 行（错位一格就会删到 PK/TF）
  const delCalls = world.ml.calls.filter(function (c) { return c.method === 'deleteRows'; });
  assert.ok(delCalls.length > 0, '本次运行必须真的删过行，否则本用例没在守东西');
  delCalls.forEach(function (c) {
    const start = c.args[0];
    for (let i = 0; i < c.args[1]; i++) {
      assert.equal(String(c.before.values[start - 1 + i][0] || '').trim(), 'INJ',
        'deleteRows 覆盖了非 INJ 行（第 ' + (start + i) + ' 行）');
    }
  });

  // 每条涂色调用覆盖的行，在当时也必须是 INJ 行
  const paintCalls = world.ml.calls.filter(function (c) { return c.method === 'setBackgrounds'; });
  assert.ok(paintCalls.length > 0, '本次运行必须真的涂过色（含清除），否则本用例没在守东西');
  paintCalls.forEach(function (c) {
    const start = c.args[0];
    for (let i = 0; i < c.args[2]; i++) {
      assert.equal(String(c.before.values[start - 1 + i][0] || '').trim(), 'INJ',
        'setBackgrounds 覆盖了非 INJ 行（第 ' + (start + i) + ' 行）');
    }
  });

  // 两台删除机台消失，四台在册机台＋当天追加机台保留（含 TF/INJ 同号的 T1GR0301）
  assert.deepEqual(rowsOf(world, 'INJ').map(function (r) { return r[3]; }).sort(),
    ['E0EN0001', 'K1KEEP01', 'K2KEEP02', 'K3KEEP03', 'T1GR0301']);

  // 邮件与 Log 把这次真删说清楚
  assert.equal(mails.length, 1);
  assert.equal(mails[0].to, 'ops@example.com');
  assert.match(mails[0].options.htmlBody, /D1DEL001/);
  assert.match(mails[0].options.htmlBody, /已删除·因表11标Y/);
  assert.match(mails[0].options.htmlBody, /已删除·表11查不到/);
  assert.equal(logs[0].status, '成功');
  assert.equal(logs[0].trigger, '手动');
  assert.match(logs[0].detail, /追加=1台, 删除=2台/);
  assert.match(logs[0].remark, /D1DEL001/, 'Log 必须留删除快照（唯一恢复依据）');
});

test('IO-2 deleteRows 的确切区间：断开的两段分别删，且 start 大的先删', () => {
  const world = setupWorld({
    mlRows: ML_RUNS_ROWS, wcRows: WC_RUNS_ROWS, recipients: [rec('ops@example.com')],
  });

  runManual();
  assertNoFailure();

  // 行号若从小到大删，后面的行号会因上移而错位 —— 这里锁的是调用序列本身
  const calls = world.ml.calls
    .filter(function (c) { return c.method === 'deleteRows'; })
    .map(function (c) { return c.args; });
  assert.deepEqual(calls, [[5, 1], [2, 2]],
    'A1ADEL01 两行一段 (2,2)、Z1ZDEL99 单行一段 (5,1)；先删大行号（5），再删小行号（2）');

  // 最终状态：只剩两台保留机，且顺序正确（移位没错行）
  assert.deepEqual(rowsOf(world, 'INJ').map(function (r) { return r[3]; }), ['K1KEEP01', 'K2KEEP02']);
  assert.equal(mails.length, 1);
});

test('IO-3 涂色在删除之后重新读背景：黄落在最终行号上，绝不错位到 PK/TF 行', () => {
  // 本用例专治"把背景色读取提到删除之前"的优化：删除让 K2KEEP02 从第 8 行上移到
  // 第 6 行，旧实现按删除前的第 8 行涂，黄会落到 PK 行（原第 10 行上移后的第 8 行）
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS, wcRows: WC_MAIN_ROWS,
    recipients: [rec('ops@example.com')],
  });

  runManual();
  assertNoFailure();

  assert.equal(bgAt(world, findRow(world, 'K2KEEP02')), '#ffff00',
    '待补全行必须在删除后的最终行号上被标黄');
  assert.equal(bgAt(world, findRow(world, 'E0EN0001')), '#ffff00',
    '当天追加的行必须当天标黄（读完背景才追加的话也会漏）');
  assert.equal(bgAt(world, findRow(world, 'H2PK0002')), '#ffffff',
    'PK 行的背景绝不能被写（错位实现在这里会把黄涂上来）');
  assert.equal(bgAt(world, findRow(world, 'T2TF0001')), '#ffff00',
    'TF 行人工遗留的黄不得被清');

  // 调用序列也锁死：清除第 5 行(K1KEEP01)、标黄第 6 行(K2KEEP02)、标黄第 11 行(追加行)
  const paintRows = world.ml.calls
    .filter(function (c) { return c.method === 'setBackgrounds'; })
    .map(function (c) { return c.args[0]; });
  assert.deepEqual(paintRows, [5, 6, 11]);
});

test('IO-4 安全阀拦下大额删除：一行没删、邮件列出全部被拦机台、追加照常', () => {
  const world = setupWorld({
    mlRows: ML_VALVE_ROWS, wcRows: WC_VALVE_ROWS, recipients: [rec('ops@example.com')],
  });

  runScheduled();
  assertNoFailure();

  assert.equal(world.ml.calls.filter(function (c) { return c.method === 'deleteRows'; }).length, 0,
    '阀拦下时 deleteRows 一次都不能调用');
  assert.deepEqual(
    world.ml.calls.filter(function (c) { return c.method === 'insertRowsAfter'; }).map(function (c) { return c.args; }),
    [[13, 1]], '追加不设阀（可逆），必须照常执行');

  const html = mails[0].options.htmlBody;
  for (let i = 1; i <= 11; i++) {
    assert.match(html, new RegExp('X' + String(i).padStart(2, '0') + 'DEL001'),
      '被拦机台必须出现在邮件里供人复核（拦下却不列出等于白拦）');
  }
  assert.match(html, /计划删除：11 台（超阈值，未执行）/);
  assert.match(html, /拟删除 11 台，超过阈值 10/);
  assert.equal((html.match(/未执行·超阈值/g) || []).length, 11, '11 台的状态都必须是「未执行」');
  assert.doesNotMatch(html, /已删除/, '一台都没删，说「已删除」就是谎报');
  assert.match(html, /已添加·需人工维护/, '追加照常，邮件按已写入报');

  // 11 台待删 + 1 台保留 + 1 台追加，一行都没少
  assert.equal(world.ml.getLastRow(), 14);
  assert.deepEqual(rowsOf(world, 'INJ').map(function (r) { return r[3]; }).sort(),
    ['E0EN0001', 'K1KEEP01'].concat(
      Array.from({ length: 11 }, function (_, i) { return 'X' + String(i + 1).padStart(2, '0') + 'DEL001'; })
    ).sort());

  // Log 摘要必须说清楚：删除=0 台，是"计划 11 台被阀拦下"，不是没有差异
  assert.equal(logs[0].status, '成功');
  assert.equal(logs[0].trigger, '定时');
  assert.match(logs[0].detail, /追加=1台, 删除=0台\(计划11台被安全阀拦下\)/);
  assert.match(logs[0].remark, /X01DEL001/, '被拦机台的快照也要进 Log');
});

test('IO-5 收件人为空：一个字节都不写，Log 记跳过', () => {
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS, wcRows: WC_MAIN_ROWS,
    recipients: [],
  });

  runManual();
  assertNoFailure();

  assert.deepEqual(world.mutations(), [],
    '收件人为空时必须早于一切写入返回（写库不可逆，静默改动不可接受）');
  assert.equal(mails.length, 0, '没有收件人不发信');
  assert.equal(logs.length, 1);
  assert.equal(logs[0].status, '跳过');
  assert.match(logs[0].detail, /无匹配收件人/);
  assert.match(logs[0].detail, /TO=0人/);
  assert.match(logs[0].detail, /计划追加1台\/删除2台/, '摘要要带上被放弃的计划量');
  assert.deepEqual(world.ml.getDataRange().getValues(), ML_MAIN_ROWS, 'MachineList 原封不动');
});

test('IO-6 Workcenter 缺必需表头：跳过且不写（缺守时会把全表当待追加灌进去）', () => {
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS,
    wcRows: [['Workcenter'], ['K1KEEP01']],   // 少了「点检无需检查Y/N」
    recipients: [rec('ops@example.com')],
  });

  runManual();
  assertNoFailure();

  assert.deepEqual(world.mutations(), [], '表头缺失时必须一个字节都不写');
  assert.equal(mails.length, 0);
  assert.equal(logs.length, 1);
  assert.equal(logs[0].status, '跳过');
  assert.match(logs[0].detail, /缺少字段/);
  assert.match(logs[0].detail, /点检无需检查Y\/N/, '必须点名缺了哪个表头');
  assert.deepEqual(world.ml.getDataRange().getValues(), ML_MAIN_ROWS, 'MachineList 不得被灌入');
});

test('IO-7 预演不写任何数据：邮件状态列说「计划…·预演未写入」', () => {
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS, wcRows: WC_MAIN_ROWS,
    recipients: [rec('ops@example.com')],
  });

  runDryRun();
  assertNoFailure();

  assert.deepEqual(world.mutations(), [], '预演不得有任何写操作（含涂色）');
  assert.equal(mails.length, 1, '预演仍要发报告邮件');

  const html = mails[0].options.htmlBody;
  assert.match(html, /计划追加·预演未写入/);
  assert.match(html, /计划删除·预演未写入/);
  assert.doesNotMatch(html, /已添加/, '预演什么都没写，说「已添加」就是谎报');
  assert.doesNotMatch(html, /已删除/, '预演什么都没删，说「已删除」就是谎报');
  assert.match(html, /E0EN0001/, '计划追加的机台要列出来');
  assert.match(html, /D1DEL001/, '计划删除的机台要列出来');

  assert.equal(logs[0].status, '成功');
  assert.equal(logs[0].trigger, '预演');
  assert.match(logs[0].detail, /预演未写入/);
  assert.equal(logs[0].remark, '', '预演没删任何行，不能写删除快照');
  assert.deepEqual(world.ml.getDataRange().getValues(), ML_MAIN_ROWS, 'MachineList 原封不动');
});

test('IO-8 幂等：连跑两遍，第二遍零写入（删除/追加/涂色都没有）', () => {
  const world = setupWorld({
    mlRows: ML_MAIN_ROWS, mlBgs: ML_MAIN_BGS, wcRows: WC_MAIN_ROWS,
    recipients: [rec('ops@example.com')],
  });

  runManual();
  assertNoFailure();
  assert.ok(world.mutations().length > 0, '第一遍必须真的写过，否则第二遍的"零写入"没有意义');
  const mailsAfterFirst = mails.length;
  const stateAfterFirst = world.ml.getDataRange().getValues();

  world.resetCalls();   // 只清调用记录，不动表内容
  runManual();
  assertNoFailure();

  assert.deepEqual(world.mutations(), [],
    '第二遍不得再写：deleteRows / insertRowsAfter / setValues / setBackgrounds 全为 0'
    + '（删除已收敛、追加已存在、颜色已在目标态）');
  assert.deepEqual(world.ml.getDataRange().getValues(), stateAfterFirst, '第二遍表内容不变');
  assert.equal(mails.length, mailsAfterFirst + 1, '第二遍仍要发信 —— 证明它真跑了，不是静默跳过');
  const secondLog = logs[logs.length - 1];
  assert.equal(secondLog.status, '成功');
  assert.match(secondLog.detail, /追加=0台, 删除=0台/);
});

// ===== 假表建模语义备忘（供 reviewer 判断会不会给出虚假的信心）=====
// 建模了（与真表一致 / 故意保留的行为）：
//  * getRange/getDataRange 返回**活窗口**：读的时刻取当前数组；窗口按坐标绑定，
//    不随 deleteRows 自动调整（真表会调整/失效 —— 模块从不跨删除持窗口，
//    而"涂色前用了变更前读到的行号"这一缺陷恰好靠此语义被放大成可观测的错位）；
//  * deleteRows 1-based，值/背景两条数组同步 splice，后续行上移；
//  * insertRowsAfter 在 afterRow 之后插入空行（5 列、无色），值/背景同步插入；
//  * getLastRow() = 最后一行"有任意非空单元格"的行（忽略尾部全空行）；
//  * getBackgrounds() 一律小写十六进制，未设置/已清除 → '#ffffff'；
//    setBackgrounds(null) 视为清除（读回 '#ffffff'）；
//  * setValues/setBackgrounds 的矩阵尺寸不符会抛错（真表如此）；
//  * 越界写入按网格空白区处理：自动补空行（真表网格 1000 行，远大于数据）。
// 未建模（不能靠本文件断言的部分）：
//  * 表格对象模型只有 openById + getSheetByName（无 getSheets/A1 记法/getValue/setValue/
//    merge/公式/保护）；日期与数字不做真表的类型强转；
//  * 无 1000 行网格上限、无配额/超时/事务回滚；
//  * GmailApp 只记录参数，不渲染 HTML、不计配额；buildHtmlTable 是简化桩，
//    不是 Common.js 的真实实现（邮件断言只做子串匹配，不做像素级 UI 校验）；
//  * writeLog 是记录器，不写真实 Log 表（真实签名/字段已按 Common.js 对齐）；
//  * getLastRow 的"全空行"定义未覆盖表内夹着整行空值的情形（夹具里没有这种行）。
