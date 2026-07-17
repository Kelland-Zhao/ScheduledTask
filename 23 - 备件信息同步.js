// V20260613.1 — 备件信息同步（从独立备件管理项目迁移并入）
// 入口：archiveMasterData（MasterData月度归档）、updateSparePartsBasicInfo（备件基础信息同步）
// 数据源：备件管理表 https://docs.google.com/spreadsheets/d/1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs

const _sp_SPREADSHEET_ID = "1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs";
const _sp_ARCHIVE_FOLDER_ID = "1lyGcCycQdywQ7Q2jwjAme_-wVq6g1Oq91";
const _sp_DRIVE_ID = "0AKZKsnGzkrDwUk9PVA"; // 00 - 注塑公共盘 / INJ Drive

// ========== MasterData 月度归档 ==========
function archiveMasterData(e) {
  const trigger = (e && e.triggerType) ? e.triggerType : "手动";
  const saas = SpreadsheetApp.openById(_sp_SPREADSHEET_ID);
  const sbnMasterData = saas.getSheetByName("MasterData");

  try {
    // 计算上个月
    const now          = new Date();
    const lastMonth    = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const archiveYear  = lastMonth.getFullYear();
    const archiveMonth = lastMonth.getMonth() + 1;
    const archiveYM    = archiveYear + "-" + String(archiveMonth).padStart(2, "0");

    // 找或创建归档 Spreadsheet（名称=年份，存放于指定文件夹）
    var ssName = String(archiveYear);
    var archiveSS = null;
    var searchQ = "name%3D%27" + ssName + "%27+and+%27" + _sp_ARCHIVE_FOLDER_ID + "%27+in+parents+and+mimeType%3D%27application%2Fvnd.google-apps.spreadsheet%27";
    var existing = _sp_driveApi("GET", "files?q=" + searchQ + "&fields=files(id,name)&pageSize=1&corpora=drive&driveId=" + _sp_DRIVE_ID + "&supportsAllDrives=true&includeItemsFromAllDrives=true");
    if (existing.files && existing.files.length > 0) {
      archiveSS = SpreadsheetApp.openById(existing.files[0].id);
    } else {
      archiveSS = SpreadsheetApp.create(ssName);
      _sp_driveApi("PATCH", "files/" + archiveSS.getId() + "?supportsAllDrives=true&addParents=" + _sp_ARCHIVE_FOLDER_ID + "&removeParents=root");
    }

    // 找或创建归档 Sheet（名称=YYYY-MM）
    let archiveSbn = archiveSS.getSheetByName(archiveYM);
    if (!archiveSbn) {
      archiveSbn = archiveSS.insertSheet(archiveYM);
    }

    // 读取 MasterData 全量数据
    const mdLr = sbnMasterData.getLastRow();
    if (mdLr <= 1) {
      writeLog("archiveMasterData", "成功", "无数据", trigger, archiveYM);
      return;
    }

    const header  = sbnMasterData.getRange(1, 1, 1, 6).getValues()[0];
    const allData = sbnMasterData.getRange(2, 1, mdLr - 1, 6).getValues();

    const archiveRows = [];
    const keepRows    = [];
    allData.forEach(row => {
      const dateVal = row[4];
      if (!dateVal) { keepRows.push(row); return; }
      const d = new Date(dateVal);
      if (d.getFullYear() === archiveYear && d.getMonth() + 1 === archiveMonth) {
        archiveRows.push(row);
      } else {
        keepRows.push(row);
      }
    });

    // 写入归档 Sheet（首次写入时添加表头）
    if (archiveRows.length > 0) {
      if (archiveSbn.getLastRow() === 0) archiveSbn.appendRow(header);
      archiveSbn.getRange(archiveSbn.getLastRow() + 1, 1, archiveRows.length, 6).setValues(archiveRows);
    }

    // 清空 MasterData 数据行并写回保留数据
    sbnMasterData.getRange(2, 1, mdLr - 1, 6).clearContent();
    if (keepRows.length > 0) {
      sbnMasterData.getRange(2, 1, keepRows.length, 6).setValues(keepRows);
    }

    const detail = "归档" + archiveRows.length + "条，保留" + keepRows.length + "条";
    writeLog("archiveMasterData", "成功", detail, trigger, archiveYM);

  } catch (err) {
    writeLog("archiveMasterData", "失败", err.message, trigger, "");
  }
}

// ========== 备件基础信息同步 ==========
function updateSparePartsBasicInfo(e) {
  const trigger = (e && e.triggerType) ? e.triggerType : "手动";
  const saas = SpreadsheetApp.openById(_sp_SPREADSHEET_ID);
  const sbnMasterData = saas.getSheetByName("MasterData");
  const sbnSafeStock  = saas.getSheetByName("安全库存数据");
  const sbnBasicInfo  = saas.getSheetByName("备件基础信息");

  try {
    const mdLr  = sbnMasterData.getLastRow();
    const mdMap = {};
    if (mdLr > 1) {
      sbnMasterData.getRange(2, 1, mdLr - 1, 5).getValues().forEach(row => {
        const mat = String(row[0]).trim();
        if (!mat) return;
        if (!mdMap[mat] || row[4] > mdMap[mat].date) {
          mdMap[mat] = { desc: row[1], stock: row[2], value: row[3], date: row[4] };
        }
      });
    }

    // 1b. 取 MasterData 中的最新日期作为同步基准
    let maxDate = null;
    Object.values(mdMap).forEach(v => { if (!maxDate || v.date > maxDate) maxDate = v.date; });

    // 2. 安全库存数据: A=物料, B=安全库存
    const ssLr    = sbnSafeStock.getLastRow();
    const safeMap = {};
    if (ssLr > 1) {
      sbnSafeStock.getRange(2, 1, ssLr - 1, 2).getValues().forEach(row => {
        const mat = String(row[0]).trim();
        if (mat) safeMap[mat] = row[1];
      });
    }

    // 3. 读取备件基础信息现有数据，按同步规则更新 B/D/E，保留 C/F/G
    const biLr         = sbnBasicInfo.getLastRow();
    const existingMats = new Set();
    let updateCount    = 0;
    let outputData     = [];

    if (biLr > 1) {
      outputData = sbnBasicInfo.getRange(2, 1, biLr - 1, 8).getValues().map(row => {
        const mat       = String(row[0]).trim();
        existingMats.add(mat);
        const safeStock = safeMap[mat] !== undefined ? safeMap[mat] : 0;
        const inLatest = mat && mdMap[mat] && maxDate && mdMap[mat].date >= maxDate;
        if (inLatest) {
          updateCount++;
          // 最新日期有记录：更新 B/D/E，保留 C/F
          return [mat, mdMap[mat].desc, row[2], safeStock, mdMap[mat].stock, row[5], row[6], mdMap[mat].value];
        } else {
          // 最新日期无记录：E 置 0，desc 优先用历史描述
          return [mat, mdMap[mat] ? mdMap[mat].desc : row[1], row[2], safeStock, 0, row[5], row[6], 0];
        }
      });
    }

    // 4. MasterData 最新日期中新增但备件基础信息中没有的物料 → 追加
    let addCount = 0;
    Object.keys(mdMap).forEach(mat => {
      if (!existingMats.has(mat) && maxDate && mdMap[mat].date >= maxDate) {
        const safeStock = safeMap[mat] !== undefined ? safeMap[mat] : 0;
        outputData.push([mat, mdMap[mat].desc, "", safeStock, mdMap[mat].stock, "", "", mdMap[mat].value]);
        addCount++;
      }
    });

    // 5. 清空旧数据并一次性写入
    if (biLr > 1) sbnBasicInfo.getRange(2, 1, biLr - 1, 8).clearContent();
    if (outputData.length > 0) sbnBasicInfo.getRange(2, 1, outputData.length, 8).setValues(outputData);

    // 6. 写入 Log
    writeLog("updateSparePartsBasicInfo", "成功", "更新" + updateCount + "条，新增" + addCount + "条", trigger, "");

  } catch (err) {
    writeLog("updateSparePartsBasicInfo", "失败", err.message, trigger, "");
  }
}

// ========== Drive REST API 辅助（支持 Shared Drive）==========
function _sp_driveApi(method, path, payload) {
  var url = "https://www.googleapis.com/drive/v3/" + path;
  var options = {
    method: method,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    contentType: "application/json",
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var body = JSON.parse(resp.getContentText());
  if (code >= 200 && code < 300) return body;
  return { error: body.error ? body.error.message : ("HTTP " + code) };
}
