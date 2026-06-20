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

function reportRecord(overrides = {}) {
  const record = row({
    1: "2026-06-19",
    5: "产品&A",
    6: "<测试说明>",
    11: "IMM-01",
    12: "项目负责人",
    14: "测试负责人",
    15: "已完成",
    16: "测试样单",
    17: "测试记录",
    ...overrides
  });
  record.links = {
    sample: "https://example.com/sample?a=1&b=2",
    record: "https://example.com/record"
  };
  return record;
}

function fakeReportSpreadsheet(gas, yesterdayRecords, tomorrowRecords) {
  const testSheet = {
    getSheetId() {
      return 636324977;
    }
  };
  return {
    getSheetByName(name) {
      if (name === gas.ITR_CONFIG.TEST_SHEET) return testSheet;
      return null;
    },
    __testSheet: testSheet,
    __recordsByDate: {
      "2026-06-19": yesterdayRecords,
      "2026-06-21": tomorrowRecords
    }
  };
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

test("_itr_buildEmailHtml renders only populated sections with red bilingual escaped content and result colors", () => {
  const gas = loadModule();
  const abnormal = reportRecord();
  abnormal.problems = ["缺少<记录>"];
  const normal = reportRecord({ 5: "正常产品", 6: "正常说明" });
  normal.rowNumber = 9;
  normal.problems = [];

  const html = gas._itr_buildEmailHtml({
    reportDate: "2026-06-20",
    sourceUrl: "https://docs.google.com/spreadsheets/d/source/edit",
    yesterday: { date: "2026-06-19", records: [abnormal, normal] },
    tomorrow: { date: "2026-06-21", records: [] }
  });

  assert.match(html, /注塑测试日报/);
  assert.match(html, /Injection Test Daily Report/);
  assert.match(html, /昨日测试复盘/);
  assert.doesNotMatch(html, /明日测试提醒/);
  assert.match(html, /#E60012/i);
  assert.match(html, /产品&amp;A/);
  assert.match(html, /&lt;测试说明&gt;/);
  assert.match(html, /缺少&lt;记录&gt;/);
  assert.match(html, /正常产品/);
  assert.match(html, /正常说明/);
  assert.match(html, /gid=636324977&amp;range=A8/);
  assert.match(html, /#fce8e6/i);
  assert.match(html, /#e6f4ea/i);
  assert.match(html, /昨日 2/);
  assert.match(html, /明日 0/);
  assert.match(html, /异常 1/);
});

test("_itr_linkHtml uses safe http text as fallback link and rejects other schemes", () => {
  const gas = loadModule();

  assert.equal(
    gas._itr_linkHtml("https://example.com/a?x=1&y=2", ""),
    '<a href="https://example.com/a?x=1&amp;y=2" style="color:#E60012;text-decoration:none;">https://example.com/a?x=1&amp;y=2</a>'
  );
  assert.equal(
    gas._itr_linkHtml("http://example.com/<sample>", ""),
    '<a href="http://example.com/&lt;sample&gt;" style="color:#E60012;text-decoration:none;">http://example.com/&lt;sample&gt;</a>'
  );
  assert.equal(gas._itr_linkHtml("javascript:alert(1)", ""), "javascript:alert(1)");
  assert.equal(gas._itr_linkHtml("file:///tmp/test", ""), "file:///tmp/test");
  assert.equal(gas._itr_linkHtml("点击", "javascript:alert(1)"), "点击");
  assert.equal(gas._itr_linkHtml("点击", "data:text/html,<script>"), "点击");
  assert.equal(gas._itr_linkHtml("点击", "ftp://example.com/file"), "点击");
  assert.match(
    gas._itr_linkHtml("点击", "HTTPS://example.com/Path"),
    /^<a href="HTTPS:\/\/example\.com\/Path"/
  );
});

test("_itr_getRecordsByDateKeys reads A:U once and groups multiple target dates", () => {
  const calls = { values: 0, richText: 0 };
  const values = [
    ["A2", "2026-06-19", ...Array(19).fill("")],
    ["A3", "2026-06-20", ...Array(19).fill("")],
    ["A4", "2026-06-21", ...Array(19).fill("")]
  ];
  const sheet = {
    getLastRow() {
      return 4;
    },
    getRange() {
      return {
        getValues() {
          calls.values++;
          return values;
        },
        getRichTextValues() {
          calls.richText++;
          return values.map(() => Array(21).fill(null));
        }
      };
    }
  };

  const gas = loadModule();
  const grouped = gas._itr_getRecordsByDateKeys(
    sheet,
    ["2026-06-19", "2026-06-21"]
  );

  assert.equal(calls.values, 1);
  assert.equal(calls.richText, 1);
  assert.equal(grouped["2026-06-19"].length, 1);
  assert.equal(grouped["2026-06-19"][0].rowNumber, 2);
  assert.equal(grouped["2026-06-21"].length, 1);
  assert.equal(grouped["2026-06-21"][0].rowNumber, 4);
  assert.equal(grouped["2026-06-20"], undefined);
});

test("_itr_run reads sheet values and rich text only once for yesterday and tomorrow", () => {
  const calls = { values: 0, richText: 0 };
  const yesterdayValues = Array(21).fill("");
  yesterdayValues[1] = "2026-06-19";
  yesterdayValues[5] = "昨日产品";
  yesterdayValues[15] = "已完成";
  yesterdayValues[17] = "https://example.com/record";
  const tomorrowValues = Array(21).fill("");
  tomorrowValues[1] = "2026-06-21";
  tomorrowValues[5] = "明日产品";
  tomorrowValues[11] = "IMM-01";
  tomorrowValues[14] = "测试负责人";
  tomorrowValues[16] = "https://example.com/sample";
  const testSheet = {
    getLastRow() {
      return 3;
    },
    getRange() {
      return {
        getValues() {
          calls.values++;
          return [yesterdayValues, tomorrowValues];
        },
        getRichTextValues() {
          calls.richText++;
          return [Array(21).fill(null), Array(21).fill(null)];
        }
      };
    }
  };
  const sent = [];
  const gas = loadModule({
    Sheets: {
      Spreadsheets: {
        get() {
          return { sheets: [] };
        }
      }
    },
    GmailApp: {
      getAliases() {
        return [];
      },
      sendEmail(...args) {
        sent.push(args);
      }
    }
  });
  const spreadsheet = {
    getSheetByName(name) {
      return name === gas.ITR_CONFIG.TEST_SHEET ? testSheet : null;
    },
    getUrl() {
      return "https://docs.google.com/spreadsheets/d/source/edit";
    }
  };
  gas._itr_getRecipients = () => ["notify@example.com"];

  const result = gas._itr_run({
    now: vm.runInContext('new Date("2026-06-20T04:00:00.000Z")', gas),
    spreadsheet
  });

  assert.equal(calls.values, 1);
  assert.equal(calls.richText, 1);
  assert.equal(result.yesterdayCount, 1);
  assert.equal(result.tomorrowCount, 1);
  assert.equal(sent.length, 1);
});

test("_itr_run skips before recipient lookup and Gmail when both date sections are empty", () => {
  let recipientCalls = 0;
  let gmailCalls = 0;
  const gas = loadModule({
    GmailApp: {
      sendEmail() {
        gmailCalls++;
      }
    }
  });
  const spreadsheet = fakeReportSpreadsheet(gas, [], []);
  gas._itr_getRecordsByDateKeys = (sheet, dateKeys) =>
    Object.fromEntries(dateKeys.map((dateKey) => [
      dateKey,
      spreadsheet.__recordsByDate[dateKey] || []
    ]));
  gas._itr_getRecipients = () => {
    recipientCalls++;
    return ["should-not-run@example.com"];
  };

  const result = gas._itr_run({
    now: vm.runInContext('new Date("2026-06-20T04:00:00.000Z")', gas),
    spreadsheet
  });

  assert.equal(result.status, "skipped");
  assert.equal(result.yesterdayCount, 0);
  assert.equal(result.tomorrowCount, 0);
  assert.equal(recipientCalls, 0);
  assert.equal(gmailCalls, 0);
});

test("_itr_run testMode sends only to the test mailbox with test subject", () => {
  const sent = [];
  const gas = loadModule({
    GmailApp: {
      getAliases() {
        return [];
      },
      sendEmail(to, subject, body, options) {
        sent.push({ to, subject, body, options });
      }
    }
  });
  const yesterday = reportRecord();
  const spreadsheet = fakeReportSpreadsheet(gas, [yesterday], []);
  gas._itr_getRecordsByDateKeys = (sheet, dateKeys) =>
    Object.fromEntries(dateKeys.map((dateKey) => [
      dateKey,
      spreadsheet.__recordsByDate[dateKey] || []
    ]));
  gas._itr_fillSmartChipLinks = (records) => records;
  gas._itr_getRecipients = () => {
    throw new Error("testMode must not read production recipients");
  };

  const result = gas._itr_run({
    now: vm.runInContext('new Date("2026-06-20T04:00:00.000Z")', gas),
    spreadsheet,
    testMode: true
  });

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, "kelland_zhao@colpal.com");
  assert.match(sent[0].subject, /^\[测试\]/);
  assert.equal(result.status, "sent");
  assert.deepEqual(Array.from(result.recipients), ["kelland_zhao@colpal.com"]);
});

test("_itr_run production merges recipients, evaluates problems and returns statistics", () => {
  const sent = [];
  const gas = loadModule({
    GmailApp: {
      getAliases() {
        return ["CSX_PlantSystem@colpal.com"];
      },
      sendEmail(to, subject, body, options) {
        sent.push({ to, subject, body, options });
      }
    }
  });
  const yesterday = reportRecord({ 15: "" });
  const tomorrow = reportRecord({ 11: "", 14: "", 16: "" });
  tomorrow.rowNumber = 10;
  const spreadsheet = fakeReportSpreadsheet(gas, [yesterday], [tomorrow]);
  gas._itr_getRecordsByDateKeys = (sheet, dateKeys) =>
    Object.fromEntries(dateKeys.map((dateKey) => [
      dateKey,
      spreadsheet.__recordsByDate[dateKey] || []
    ]));
  gas._itr_fillSmartChipLinks = (records) => records;
  gas._itr_getRecipients = (source, records) => {
    assert.equal(records.length, 2);
    return ["notify@example.com", "owner@example.com"];
  };

  const result = gas._itr_run({
    now: vm.runInContext('new Date("2026-06-20T04:00:00.000Z")', gas),
    spreadsheet
  });

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, "notify@example.com,owner@example.com");
  assert.match(sent[0].subject, /【注塑测试日报】2026-06-20 昨日复盘 & 明日提醒/);
  assert.equal(sent[0].options.from, "CSX_PlantSystem@colpal.com");
  assert.deepEqual(
    JSON.parse(JSON.stringify(result)),
    {
      status: "sent",
      reportDate: "2026-06-20",
      yesterdayCount: 1,
      tomorrowCount: 1,
      abnormalCount: 2,
      recipients: ["notify@example.com", "owner@example.com"]
    }
  );
});

test("_itr_sendEmail uses configured alias only when available", () => {
  const withAlias = [];
  const gasWithAlias = loadModule({
    GmailApp: {
      getAliases() {
        return ["other@example.com", "csx_plantsystem@colpal.com"];
      },
      sendEmail(to, subject, body, options) {
        withAlias.push(options);
      }
    }
  });
  gasWithAlias._itr_sendEmail(["a@example.com"], "subject", "<p>html</p>");
  assert.equal(withAlias[0].from, "CSX_PlantSystem@colpal.com");
  assert.equal(withAlias[0].name, "注塑测试日报");

  const withoutAlias = [];
  const gasWithoutAlias = loadModule({
    GmailApp: {
      getAliases() {
        return ["other@example.com"];
      },
      sendEmail(to, subject, body, options) {
        withoutAlias.push(options);
      }
    }
  });
  gasWithoutAlias._itr_sendEmail(["a@example.com"], "subject", "<p>html</p>");
  assert.equal(Object.hasOwn(withoutAlias[0], "from"), false);
  assert.equal(withoutAlias[0].name, "注塑测试日报");
});

test("sendInjectionTestDailyReport logs scheduled only for explicit scheduled trigger type", () => {
  const logs = [];
  const gas = loadModule({
    writeLog(...args) {
      logs.push(args);
    }
  });
  gas._itr_run = () => ({
    status: "sent",
    yesterdayCount: 1,
    tomorrowCount: 2,
    abnormalCount: 1,
    recipients: ["a@example.com"]
  });
  gas.sendInjectionTestDailyReport({ triggerType: "scheduled" });
  assert.deepEqual(logs[0].slice(0, 4), [
    "sendInjectionTestDailyReport",
    "成功",
    "昨日 1 条，明日 2 条，异常 1 条，收件人 1 人",
    "定时"
  ]);

  gas.sendInjectionTestDailyReport({ triggerUid: "ordinary-event" });
  assert.deepEqual(logs[1].slice(0, 4), [
    "sendInjectionTestDailyReport",
    "成功",
    "昨日 1 条，明日 2 条，异常 1 条，收件人 1 人",
    "手动"
  ]);
});

test("sendInjectionTestDailyReport logs manual skip without an event", () => {
  const logs = [];
  const gas = loadModule({
    writeLog(...args) {
      logs.push(args);
    }
  });
  gas._itr_run = () => ({
    status: "skipped",
    yesterdayCount: 0,
    tomorrowCount: 0,
    abnormalCount: 0,
    recipients: []
  });

  gas.sendInjectionTestDailyReport();

  assert.deepEqual(logs[0].slice(0, 4), [
    "sendInjectionTestDailyReport",
    "跳过",
    "昨日与明日均无测试记录",
    "手动"
  ]);
});

test("sendInjectionTestDailyReport logs failure and rethrows", () => {
  const logs = [];
  const gas = loadModule({
    writeLog(...args) {
      logs.push(args);
    }
  });
  gas._itr_run = () => {
    throw new Error("mail failed");
  };

  assert.throws(
    () => gas.sendInjectionTestDailyReport(),
    /mail failed/
  );
  assert.deepEqual(logs[0].slice(0, 4), [
    "sendInjectionTestDailyReport",
    "失败",
    "mail failed",
    "手动"
  ]);
});

test("_itr_safeWriteLog swallows logging failures and reports them to console", () => {
  const errors = [];
  const gas = loadModule({
    console: {
      error(...args) {
        errors.push(args);
      }
    },
    writeLog() {
      throw new Error("log unavailable");
    }
  });

  assert.equal(
    gas._itr_safeWriteLog("fn", "成功", "detail", "手动", ""),
    false
  );
  assert.equal(errors.length, 1);
  assert.match(errors[0].join(" "), /log unavailable/);
});

test("entry rethrows original run error even when failure logging also fails", () => {
  const gas = loadModule({
    console: { error() {} },
    writeLog() {
      throw new Error("log unavailable");
    }
  });
  const original = new Error("original run failure");
  gas._itr_run = () => {
    throw original;
  };

  assert.throws(
    () => gas.sendInjectionTestDailyReport(),
    (error) => error === original
  );
});

test("entry returns sent result when success logging fails", () => {
  const gas = loadModule({
    console: { error() {} },
    writeLog() {
      throw new Error("log unavailable");
    }
  });
  const sentResult = {
    status: "sent",
    yesterdayCount: 1,
    tomorrowCount: 0,
    abnormalCount: 0,
    recipients: ["a@example.com"]
  };
  gas._itr_run = () => sentResult;

  assert.equal(gas.sendInjectionTestDailyReport(), sentResult);
});

test("testInjectionTestDailyReport forces test mode and test log trigger", () => {
  const logs = [];
  let receivedOptions;
  const gas = loadModule({
    writeLog(...args) {
      logs.push(args);
    }
  });
  gas._itr_run = (options) => {
    receivedOptions = options;
    return {
      status: "sent",
      yesterdayCount: 1,
      tomorrowCount: 0,
      abnormalCount: 0,
      recipients: ["kelland_zhao@colpal.com"]
    };
  };

  gas.testInjectionTestDailyReport();

  assert.equal(receivedOptions.testMode, true);
  assert.deepEqual(logs[0].slice(0, 4), [
    "testInjectionTestDailyReport",
    "成功",
    "昨日 1 条，明日 0 条，异常 0 条，收件人 1 人",
    "测试"
  ]);
});
