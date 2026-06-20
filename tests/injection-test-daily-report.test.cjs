const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const modulePath = path.join(__dirname, "..", "31 - 注塑测试日报.js");

function loadModule(overrides = {}) {
  const context = vm.createContext({
    Utilities: {
      formatDate(value, timeZone, pattern) {
        assert.equal(timeZone, "Asia/Shanghai");
        assert.equal(pattern, "yyyy-MM-dd");
        const parts = new Intl.DateTimeFormat("en-CA", {
          timeZone,
          year: "numeric",
          month: "2-digit",
          day: "2-digit"
        }).formatToParts(value);
        const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
        return `${values.year}-${values.month}-${values.day}`;
      }
    },
    ...overrides
  });
  vm.runInContext(fs.readFileSync(modulePath, "utf8"), context, {
    filename: modulePath
  });
  return context;
}

function richText(directLink, runLinks = []) {
  return {
    getLinkUrl() {
      return directLink;
    },
    getRuns() {
      return runLinks.map((link) => ({
        getLinkUrl() {
          return link;
        }
      }));
    }
  };
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
  assert.equal(gas._itr_normalizeDate("2026-06-09"), "2026-06-09");
  assert.equal(gas._itr_normalizeDate("2026/06/19"), "2026-06-19");
  assert.equal(gas._itr_normalizeDate(""), "");
  assert.equal(gas._itr_normalizeDate(null), "");
});

test("_itr_normalizeDate rejects invalid and non-spec date strings", () => {
  const gas = loadModule();
  for (const value of [
    "2026-02-30",
    "2026/13/01",
    "2026-6-9",
    "6/19",
    "June 19, 2026",
    "2026-02-30T12:00:00.000Z",
    "2026-06-19T25:00:00.000Z",
    "2026-06-19 12:00:00"
  ]) {
    assert.equal(gas._itr_normalizeDate(value), "", value);
  }
});

test("_itr_normalizeDate converts ISO timestamps across the UTC Shanghai date boundary", () => {
  const gas = loadModule();
  assert.equal(
    gas._itr_normalizeDate("2026-06-18T16:30:00.000Z"),
    "2026-06-19"
  );
});

test("_itr_addDays adds calendar days without mutating the input", () => {
  const gas = loadModule();
  const source = vm.runInContext('new Date("2026-06-20T04:00:00.000Z")', gas);
  const result = gas._itr_addDays(source, -1);

  assert.equal(result.toISOString(), "2026-06-19T04:00:00.000Z");
  assert.equal(source.toISOString(), "2026-06-20T04:00:00.000Z");
});

test("_itr_dateKey formats a Date in the configured time zone", () => {
  const gas = loadModule();
  const source = vm.runInContext('new Date("2026-06-18T16:30:00.000Z")', gas);

  assert.equal(gas._itr_dateKey(source), "2026-06-19");
});

test("_itr_getRichTextLink prefers a direct link and falls back to the first run link", () => {
  const gas = loadModule();

  assert.equal(
    gas._itr_getRichTextLink(richText("https://direct.example", ["https://run.example"])),
    "https://direct.example"
  );
  assert.equal(
    gas._itr_getRichTextLink(richText(null, [null, "https://first.example", "https://second.example"])),
    "https://first.example"
  );
  assert.equal(gas._itr_getRichTextLink(null), "");
});

test("_itr_getRecordsByDate reads A:U once and keeps matching rows with actual row numbers", () => {
  const gas = loadModule();
  const values = [
    ["A2", "2026-06-19", ...Array(19).fill("")],
    ["A3", "2026-06-20", ...Array(14).fill(""), "sample text", "record text", "", "", ""],
    ["A4", "2026/06/20", ...Array(14).fill(""), "sample 2", "record 2", "", "", ""]
  ];
  const richRows = values.map(() => Array(21).fill(null));
  richRows[1][16] = richText("https://sample.example");
  richRows[1][17] = richText(null, ["https://record.example"]);
  richRows[2][16] = richText(null, []);
  richRows[2][17] = richText(null, []);
  const calls = [];
  const sheet = {
    getLastRow() {
      return 4;
    },
    getRange(rowNumber, column, rowCount, columnCount) {
      calls.push([rowNumber, column, rowCount, columnCount]);
      return {
        getValues() {
          return values;
        },
        getRichTextValues() {
          return richRows;
        }
      };
    }
  };

  const records = Array.from(gas._itr_getRecordsByDate(sheet, "2026-06-20"));

  assert.deepEqual(calls, [[2, 1, 3, 21]]);
  assert.equal(records.length, 2);
  assert.equal(records[0].rowNumber, 3);
  assert.equal(records[0].values[0], "A3");
  assert.deepEqual(
    JSON.parse(JSON.stringify(records[0].links)),
    { sample: "https://sample.example", record: "https://record.example" }
  );
  assert.equal(records[1].rowNumber, 4);
  assert.deepEqual(
    JSON.parse(JSON.stringify(records[1].links)),
    { sample: "", record: "" }
  );
});

test("_itr_fillSmartChipLinks fills missing chip URIs and continues after one cell fails", () => {
  const calls = [];
  const warnings = [];
  const gas = loadModule({
    console: {
      warn(message) {
        warnings.push(message);
      }
    },
    Sheets: {
      Spreadsheets: {
        get(spreadsheetId, options) {
          calls.push([spreadsheetId, options.ranges[0]]);
          if (options.ranges[0].endsWith("!R5")) {
            throw new Error("cell unavailable");
          }
          return {
            sheets: [{
              data: [{
                rowData: [{
                  values: [{
                    chipRuns: [
                      { chip: {} },
                      {
                        chip: {
                          richLinkProperties: {
                            uri: "https://chip.example/sample"
                          }
                        }
                      }
                    ]
                  }]
                }]
              }]
            }]
          };
        }
      }
    }
  });
  const first = row({ 16: "样单", 17: "记录" });
  first.rowNumber = 5;
  const second = row({ 16: "已有样单", 17: "" });
  second.rowNumber = 6;
  second.links.sample = "https://existing.example/sample";

  gas._itr_fillSmartChipLinks([first, second], "Test's Plan");

  assert.deepEqual(calls, [
    [gas.ITR_CONFIG.SPREADSHEET_ID, "'Test''s Plan'!Q5"],
    [gas.ITR_CONFIG.SPREADSHEET_ID, "'Test''s Plan'!R5"]
  ]);
  assert.equal(first.links.sample, "https://chip.example/sample");
  assert.equal(first.links.record, "");
  assert.equal(second.links.sample, "https://existing.example/sample");
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /R5/);
});

test("_itr_getRecipients merges notification-list emails with exact owner matches and deduplicates", () => {
  const gas = loadModule();
  const notificationSheet = {
    getDataRange() {
      return {
        getValues() {
          return [
            ["邮箱", "通知类型"],
            [" Notify@Example.com ", "测试后日报"],
            ["before@example.com", "每日测试前日报提醒"],
            ["skip@example.com", "其他通知"]
          ];
        }
      };
    }
  };
  const nameSheet = {
    getDataRange() {
      return {
        getValues() {
          return [
            ["姓名", "邮箱"],
            [" Alice ", "owner@example.com"],
            ["Bob", "notify@example.com"],
            ["Alice Chen", "wrong@example.com"]
          ];
        }
      };
    }
  };
  const spreadsheet = {
    getSheetByName(name) {
      if (name === gas.ITR_CONFIG.NOTIFICATION_SHEET) return notificationSheet;
      if (name === gas.ITR_CONFIG.NAME_SHEET) return nameSheet;
      return null;
    }
  };
  const alice = row({ 12: "Alice" });
  const bob = row({ 12: " Bob " });

  assert.deepEqual(
    Array.from(gas._itr_getRecipients(spreadsheet, [alice, bob])),
    ["notify@example.com", "before@example.com", "owner@example.com"]
  );
  assert.deepEqual(
    Array.from(gas._itr_getRecipients(spreadsheet, [])),
    ["notify@example.com", "before@example.com"]
  );
});

test("_itr_getRecipients throws clear errors when required sheets are missing", () => {
  const gas = loadModule();
  const existingSheet = {
    getDataRange() {
      return { getValues() { return []; } };
    }
  };

  assert.throws(
    () => gas._itr_getRecipients({ getSheetByName() { return null; } }, []),
    /通知清单/
  );
  assert.throws(
    () => gas._itr_getRecipients({
      getSheetByName(name) {
        return name === gas.ITR_CONFIG.NOTIFICATION_SHEET ? existingSheet : null;
      }
    }, []),
    /Name Database/
  );
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

test("_itr_uniqueEmails rejects unsafe characters and invalid domain labels", () => {
  const gas = loadModule();
  assert.deepEqual(
    Array.from(gas._itr_uniqueEmails([
      "normal.user@colpal.com",
      "comma,user@colpal.com",
      "semi;user@colpal.com",
      '"quoted"@colpal.com',
      "Name <user@colpal.com>",
      "white space@colpal.com",
      "line\nbreak@colpal.com",
      "user@-colpal.com",
      "user@colpal-.com",
      "user@colpal..com"
    ])),
    ["normal.user@colpal.com"]
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
