const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const modulePath = path.join(__dirname, "..", "31 - 注塑测试日报.js");

function loadModule() {
  const context = vm.createContext({
    Utilities: {
      formatDate(value) {
        const parts = new Intl.DateTimeFormat("en-CA", {
          timeZone: "Asia/Shanghai",
          year: "numeric",
          month: "2-digit",
          day: "2-digit"
        }).formatToParts(value);
        const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
        return `${values.year}-${values.month}-${values.day}`;
      }
    }
  });
  vm.runInContext(fs.readFileSync(modulePath, "utf8"), context, {
    filename: modulePath
  });
  return context;
}

function row(overrides = {}) {
  const values = Array(21).fill("");
  Object.entries(overrides).forEach(([index, value]) => {
    values[Number(index)] = value;
  });
  return { rowNumber: 8, values, links: {} };
}

test("ITR_CONFIG contains the approved data source and runtime settings", () => {
  const gas = loadModule();
  assert.equal(
    gas.ITR_CONFIG.SPREADSHEET_ID,
    "17ys3UDFWjhfaPnk0TErqqeU0FnMP7nsRoRsTmlmm2fg"
  );
  assert.equal(gas.ITR_CONFIG.TEST_SHEET, "注塑测试");
  assert.equal(gas.ITR_CONFIG.NOTIFICATION_SHEET, "通知清单");
  assert.equal(gas.ITR_CONFIG.NAME_SHEET, "Name Database");
  assert.equal(gas.ITR_CONFIG.TIME_ZONE, "Asia/Shanghai");
  assert.equal(gas.ITR_CONFIG.TEST_EMAIL, "kelland_zhao@colpal.com");
  assert.equal(gas.ITR_CONFIG.FROM_EMAIL, "CSX_PlantSystem@colpal.com");
});

test("_itr_normalizeDate supports Date, ISO, dash, slash and empty values", () => {
  const gas = loadModule();
  const realmDate = vm.runInContext(
    'new Date("2026-06-19T00:00:00+08:00")',
    gas
  );
  assert.equal(gas._itr_normalizeDate(realmDate), "2026-06-19");
  assert.equal(gas._itr_normalizeDate("2026-06-19T00:00:00.000Z"), "2026-06-19");
  assert.equal(gas._itr_normalizeDate("2026-6-9"), "2026-06-09");
  assert.equal(gas._itr_normalizeDate("2026/06/19"), "2026-06-19");
  assert.equal(gas._itr_normalizeDate(""), "");
  assert.equal(gas._itr_normalizeDate(null), "");
});

test("_itr_isSkippedStatus recognizes full production, offline mold and cancelled", () => {
  const gas = loadModule();
  assert.equal(gas._itr_isSkippedStatus("满产"), true);
  assert.equal(gas._itr_isSkippedStatus(" 模具不在线 "), true);
  assert.equal(gas._itr_isSkippedStatus("取消"), true);
  assert.equal(gas._itr_isSkippedStatus("已完成"), false);
  assert.equal(gas._itr_isSkippedStatus(""), false);
});

test("_itr_checkAfterProblems applies completion and test-record rules", () => {
  const gas = loadModule();
  assert.deepEqual(
    Array.from(gas._itr_checkAfterProblems(row())),
    ["完成状态未维护"]
  );
  assert.deepEqual(
    Array.from(gas._itr_checkAfterProblems(row({ 15: "已完成" }))),
    ["无测试记录"]
  );
  assert.deepEqual(
    Array.from(gas._itr_checkAfterProblems(row({ 15: "已完成", 17: "记录链接" }))),
    []
  );
  for (const status of ["满产", "模具不在线", "取消"]) {
    assert.deepEqual(
      Array.from(gas._itr_checkAfterProblems(row({ 15: status }))),
      []
    );
  }
});

test("_itr_checkBeforeProblems checks machine, owner and sample", () => {
  const gas = loadModule();
  assert.deepEqual(
    Array.from(gas._itr_checkBeforeProblems(row())),
    ["测试机台未维护", "测试负责人未维护", "测试样单链接未维护"]
  );
  assert.deepEqual(
    Array.from(gas._itr_checkBeforeProblems(row({
      11: "IMM-01",
      14: "张三",
      16: "样单链接"
    }))),
    []
  );
  for (const status of ["满产", "模具不在线", "取消"]) {
    assert.deepEqual(
      Array.from(gas._itr_checkBeforeProblems(row({ 15: status }))),
      []
    );
  }
});

test("_itr_uniqueEmails trims, lowercases, validates and deduplicates", () => {
  const gas = loadModule();
  assert.deepEqual(
    Array.from(gas._itr_uniqueEmails([
      " Kelland_Zhao@colpal.com ",
      "kelland_zhao@colpal.com",
      "valid.user+tag@example.co.uk",
      "invalid",
      "",
      null
    ])),
    ["kelland_zhao@colpal.com", "valid.user+tag@example.co.uk"]
  );
});

test("_itr_escapeHtml escapes ampersand, brackets, quotes and apostrophes", () => {
  const gas = loadModule();
  assert.equal(
    gas._itr_escapeHtml(`A&B <tag a="x">'ok'</tag>`),
    "A&amp;B &lt;tag a=&quot;x&quot;&gt;&#39;ok&#39;&lt;/tag&gt;"
  );
  assert.equal(gas._itr_escapeHtml(null), "");
});
