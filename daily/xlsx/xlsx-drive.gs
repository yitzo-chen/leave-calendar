/**
 * 日報 xlsx 雲端存檔（Apps Script 端）。
 * 依賴：xlsx-builder.gs（XlsxBuilder）＋ google-apps-script.gs 內的 SHEET_* / CASE_COL_* 常數與 normalizeDate / unsanitizeCell / getCaseRow_。
 *
 * 流程：讀試算表資料（依案場過濾） → 用 Drive 上共用的「日報範本.xlsx」產生某天（及之後已存在日期）的分頁
 *       → 組成一個 xlsx 存回同資料夾、該案場專屬的「施工日報彙整_<案場名稱>.xlsx」。
 * 各案場格式相同、共用同一份範本，只有輸出檔案、基本資料、資料列各自獨立（見 project_daily_report_webify 決議）。
 * 對外入口：xlsxUpdateForDate(date, caseId, caseInfo)、xlsxRebuildAll(caseId?)、xlsxSelfTest()。
 */
var XLSX_TEMPLATE_NAME = "日報範本.xlsx";
var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** 該案場的輸出檔檔名（案場名稱可能之後改名，檔名以目前的名稱為準） */
function xlsxOutputName_(caseId, caseInfo) {
  var label = (caseInfo && caseInfo.name) ? caseInfo.name : caseId;
  return "施工日報彙整_" + label + ".xlsx";
}

// ---------- Drive 與 zip ----------
function xlsxFolder_() {
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/** xlsx Blob → {路徑: 文字}；略過資料夾項目與二進位零件（printerSettings） */
function xlsxUnzip_(blob) {
  var blobs = Utilities.unzip(blob.setContentType("application/zip")), parts = {}; // unzip 要求類型為 application/zip
  for (var i = 0; i < blobs.length; i++) {
    var name = blobs[i].getName();
    if (/\/$/.test(name) || /\.bin$/.test(name)) continue;
    parts[name] = blobs[i].getDataAsString("UTF-8");
  }
  return parts;
}

/** {路徑: 文字} → xlsx Blob（[Content_Types].xml 放第一個） */
function xlsxZip_(parts, fileName) {
  var names = Object.keys(parts).sort(function (a, b) {
    var ra = a === "[Content_Types].xml" ? 0 : a === "_rels/.rels" ? 1 : 2;
    var rb = b === "[Content_Types].xml" ? 0 : b === "_rels/.rels" ? 1 : 2;
    return ra - rb || (a < b ? -1 : a > b ? 1 : 0);
  });
  var blobs = names.map(function (n) { return Utilities.newBlob(parts[n], "text/xml", n); });
  return Utilities.zip(blobs, fileName).setContentType(XLSX_MIME);
}

function xlsxLoadTemplate_() {
  var files = xlsxFolder_().getFilesByName(XLSX_TEMPLATE_NAME);
  if (!files.hasNext()) throw new Error("找不到範本檔「" + XLSX_TEMPLATE_NAME + "」，請放在與試算表相同的 Drive 資料夾");
  var parts = xlsxUnzip_(files.next().getBlob());
  if (!parts["xl/worksheets/sheet1.xml"] || !parts["xl/styles.xml"]) throw new Error("範本檔結構不符（找不到 sheet1 或 styles）");
  return XlsxBuilder.prepareTemplate(parts);
}

/** 取得該案場的輸出檔（先看記錄的檔案 ID，再依名稱找），沒有回傳 null */
function xlsxFindOutput_(caseId, caseInfo) {
  var props = PropertiesService.getScriptProperties();
  var propKey = "XLSX_FILE_ID_" + caseId;
  var id = props.getProperty(propKey);
  if (id) {
    try {
      var f = DriveApp.getFileById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* 檔案已被刪除，往下重新找 */ }
  }
  var outputName = xlsxOutputName_(caseId, caseInfo);
  var it = xlsxFolder_().getFilesByName(outputName);
  if (it.hasNext()) {
    var found = it.next();
    props.setProperty(propKey, found.getId());
    return found;
  }
  return null;
}

/** 覆蓋既有 xlsx 檔內容（保持同一個檔案 ID 與分享連結）；用 Drive REST 直接更新二進位內容 */
function xlsxOverwrite_(file, blob) {
  var res = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files/" + file.getId() + "?uploadType=media", {
    method: "patch",
    contentType: XLSX_MIME,
    payload: blob.getBytes(),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error("更新 xlsx 失敗（HTTP " + res.getResponseCode() + "）：" + res.getContentText().slice(0, 200));
}

function xlsxSave_(blob, caseId, caseInfo) {
  var file = xlsxFindOutput_(caseId, caseInfo);
  if (file) {
    xlsxOverwrite_(file, blob);
    return file;
  }
  var created = xlsxFolder_().createFile(blob.setName(xlsxOutputName_(caseId, caseInfo)));
  PropertiesService.getScriptProperties().setProperty("XLSX_FILE_ID_" + caseId, created.getId());
  return created;
}

// ---------- 讀取試算表資料（依案場過濾） ----------
/** 一次讀進某案場需要的資料（日報頭 / 日報記錄 / 本工出勤 / 清單，皆依「案場」欄過濾＋案場設定的基本資料），並算好逐日累計 */
function xlsxLoadDb_(caseId, caseInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function rows(name) { // 資料分頁：第 1 列是表頭，略過
    var v = ss.getSheetByName(name).getDataRange().getValues();
    v.shift();
    return v;
  }
  // 以日期/項目名稱當 key 的物件用無原型物件，避免 constructor / __proto__ 等名稱汙染
  var db = { headers: {}, recordsByDate: {}, attendanceByDate: {}, catalog: [], categoryByName: Object.create(null), basic: {}, records: [], warnings: [] };

  db.basic = {
    "業主": (caseInfo && caseInfo.owner) || "",
    "工程名稱": (caseInfo && caseInfo.name) || "",
    "合約金額（元）": (caseInfo && caseInfo.contract) || "",
    "開工日期（YYYY/MM/DD）": (caseInfo && caseInfo.startDate) || "",
    "公司名稱": (caseInfo && caseInfo.company) || "",
  };

  rows(SHEET_LIST).forEach(function (r) {
    if (String(r[CASE_COL_LIST - 1] || "").trim() !== caseId) return;
    var name = unsanitizeCell(r[1]).trim();
    if (!name) return;
    db.catalog.push({ category: r[0], name: name, active: String(r[3]).trim().toUpperCase() === "TRUE" });
    db.categoryByName[name] = r[0];
  });

  rows(SHEET_HEADER).forEach(function (r) {
    if (String(r[CASE_COL_HEADER - 1] || "").trim() !== caseId) return;
    var d = normalizeDate(r[0]);
    if (!d) return;
    db.headers[d] = {
      weather: unsanitizeCell(r[1]), status: unsanitizeCell(r[2]), todayWork: unsanitizeCell(r[3]), tomorrowPlan: unsanitizeCell(r[4]), remark: unsanitizeCell(r[5]), reporter: unsanitizeCell(r[6]),
      directorAm: unsanitizeCell(r[9]), directorPm: unsanitizeCell(r[10]), safetyAm: unsanitizeCell(r[11]), safetyPm: unsanitizeCell(r[12]),
    };
  });

  rows(SHEET_RECORD).forEach(function (r) {
    if (String(r[CASE_COL_RECORD - 1] || "").trim() !== caseId) return;
    var d = normalizeDate(r[0]);
    if (!d) return;
    var rec = { date: d, name: unsanitizeCell(r[1]).trim(), am: Number(r[2]) || 0, pm: Number(r[3]) || 0 };
    db.records.push(rec);
    (db.recordsByDate[d] = db.recordsByDate[d] || Object.create(null))[rec.name] = rec;
  });

  rows(SHEET_ATTENDANCE).forEach(function (r) {
    if (String(r[CASE_COL_ATTENDANCE - 1] || "").trim() !== caseId) return;
    var d = normalizeDate(r[0]);
    if (!d) return;
    (db.attendanceByDate[d] = db.attendanceByDate[d] || []).push({
      name: unsanitizeCell(r[1]).trim(), am: r[2] === "V", pm: r[3] === "V",
      amHours: Number(r[4]) || 0, pmHours: Number(r[5]) || 0, reason: unsanitizeCell(r[6]),
    });
  });

  db.cumByDate = XlsxBuilder.computeCumulatives(db.records, db.categoryByName, db.warnings);
  return db;
}

/** 組出某一天要交給 XlsxBuilder.buildDaySheet 的資料；該日沒有日報頭則回傳 null */
function xlsxDayData_(db, date) {
  var h = db.headers[date];
  if (!h) return null;
  var cum = XlsxBuilder.cumulativeAsOf(db.cumByDate, date);
  var today = db.recordsByDate[date] || {};
  var items = [];
  db.catalog.forEach(function (c) {
    var t = today[c.name];
    if (!c.active && !t) return; // 已停用的項目，只有當天有紀錄才列出
    items.push({ category: c.category, name: c.name, am: t ? t.am : 0, pm: t ? t.pm : 0, cumulative: cum[c.name] || 0 });
  });
  return { date: date, basic: db.basic, header: h, items: items, attendance: db.attendanceByDate[date] || [] };
}

// ---------- 對外入口 ----------
/**
 * 更新某案場的雲端 xlsx：重寫指定日期，以及該檔案中「不早於該日」的所有分頁（後面日期的累計會跟著變）。
 * 指定日期已沒有日報頭（資料被刪除）時，會移除該日分頁；分頁全被移除時把檔案丟垃圾桶（回傳 removed:true）。
 * date 為 null 時，依該案場日報頭的全部日期整個重建。
 * 失敗時拋出例外，由呼叫端決定是否影響日報送出。
 * @return {ok, fileId, url, sheets, rebuilt, warnings}
 */
function xlsxRun_(date, caseId, caseInfo) {
  var db = xlsxLoadDb_(caseId, caseInfo);
  var template = xlsxLoadTemplate_();
  var tsheet = template["xl/worksheets/sheet1.xml"];
  var file = date ? xlsxFindOutput_(caseId, caseInfo) : null;
  var existing = [], full = !date;
  if (file) {
    var oldParts = xlsxUnzip_(file.getBlob());
    // 範本改版（樣式表不同）時，舊分頁的樣式編號會對不上，必須全部重建
    if (oldParts["xl/styles.xml"] !== template["xl/styles.xml"]) full = true;
    else existing = XlsxBuilder.readSheets(oldParts);
  }
  var byName = {};
  existing.forEach(function (s) { byName[s.name] = s; });

  var targets = {};
  if (full) Object.keys(db.headers).forEach(function (d) { targets[d] = true; });
  else {
    targets[date] = true;
    existing.forEach(function (s) {
      var d = XlsxBuilder.dateForSheetName(s.name);
      if (d && d >= date) targets[d] = true;
    });
  }

  var warnings = db.warnings.slice(), rebuilt = 0;
  Object.keys(targets).sort().forEach(function (d) {
    var data = xlsxDayData_(db, d);
    if (!data) { delete byName[XlsxBuilder.sheetNameForDate(d)]; return; } // 該日已沒有日報頭（被刪除）：移除舊分頁
    var r = XlsxBuilder.buildDaySheet(tsheet, data);
    byName[XlsxBuilder.sheetNameForDate(d)] = { name: XlsxBuilder.sheetNameForDate(d), xml: r.xml };
    r.warnings.forEach(function (w) { warnings.push(d + "：" + w); });
    rebuilt++;
  });

  var sheets = Object.keys(byName).map(function (k) { return byName[k]; });
  if (!sheets.length) {
    // 最後一個分頁也被移除、整份 xlsx 已沒有內容：把檔案丟垃圾桶（可還原），下次送出日報時會用範本重建
    if (date && file) {
      file.setTrashed(true);
      PropertiesService.getScriptProperties().deleteProperty("XLSX_FILE_ID_" + caseId);
      return { ok: true, removed: true, sheets: 0, rebuilt: 0, full: full, warnings: warnings };
    }
    throw new Error("沒有任何日期可寫入 xlsx");
  }
  var active = date ? XlsxBuilder.sheetNameForDate(date) : null;
  if (active && !byName[active]) active = null; // 該日分頁已被移除，改用預設的作用中分頁
  var parts = XlsxBuilder.assembleWorkbook(template, sheets, active);
  var saved = xlsxSave_(xlsxZip_(parts, xlsxOutputName_(caseId, caseInfo)), caseId, caseInfo);
  return { ok: true, fileId: saved.getId(), url: saved.getUrl(), sheets: sheets.length, rebuilt: rebuilt, full: full, warnings: warnings };
}

function xlsxUpdateForDate(date, caseId, caseInfo) { return xlsxRun_(date, caseId, caseInfo); }

/** 該案場目前的 xlsx 內有沒有該日期的分頁（後台管理「刪除某天資料」用來判斷 xlsx 是否要處理） */
function xlsxHasSheetForDate(date, caseId, caseInfo) {
  var file = xlsxFindOutput_(caseId, caseInfo);
  if (!file) return false;
  var name = XlsxBuilder.sheetNameForDate(date);
  return XlsxBuilder.readSheets(xlsxUnzip_(file.getBlob())).some(function (s) { return s.name === name; });
}

/**
 * 管理用：依試算表資料整個重建 xlsx（範本改版或檔案被刪除後使用），在編輯器直接執行。
 * 傳 caseId 只重建該案場；不傳則依「案場設定」把所有啟用中的案場各自重建一次。
 */
function xlsxRebuildAll(caseId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!caseId) {
    var caseSheet = ss.getSheetByName(SHEET_CASES);
    var rows = caseSheet.getDataRange().getValues();
    rows.shift();
    var results = [];
    rows.filter(function (r) { return String(r[2]).trim().toUpperCase() === "TRUE"; })
      .forEach(function (r) {
        var id = String(r[0]).trim();
        var info = getCaseRow_(ss, id);
        var res = xlsxRun_(null, id, info);
        results.push({ caseId: id, result: res });
        Logger.log(id + "：" + JSON.stringify(res));
      });
    return results;
  }
  var caseInfo = getCaseRow_(ss, caseId);
  if (!caseInfo) throw new Error("案場不存在或已停用：" + caseId);
  var r = xlsxRun_(null, caseId, caseInfo);
  Logger.log(JSON.stringify(r));
  return r;
}

/**
 * 自我測試（不動任何案場的正式檔）：用範本與假資料產生「施工日報彙整-自我測試.xlsx」，再讀回確認。
 * 用來確認 Utilities.zip/unzip 與 Drive 權限在真實環境可用。在編輯器直接執行。
 */
function xlsxSelfTest() {
  var template = xlsxLoadTemplate_();
  var tsheet = template["xl/worksheets/sheet1.xml"];
  var sheets = ["2026-09-18", "2026-09-19"].map(function (d, i) {
    var r = XlsxBuilder.buildDaySheet(tsheet, {
      date: d, basic: { "公司名稱": "", "業主": "自我測試", "工程名稱": "自我測試" },
      header: { weather: "晴", status: "施工", todayWork: "測試第" + (i + 1) + "天", remark: "備註", reporter: "測試員", directorAm: "甲", directorPm: "甲", safetyAm: "乙", safetyPm: "乙" },
      items: [{ category: "工種", name: "公司工", am: i + 1, pm: i + 1, cumulative: (i + 1) * (i + 2) / 2 }],
      attendance: [{ name: "丙", am: true, pm: true, amHours: 0, pmHours: i, reason: i ? "趕工" : "" }],
    });
    return { name: XlsxBuilder.sheetNameForDate(d), xml: r.xml };
  });
  var parts = XlsxBuilder.assembleWorkbook(template, sheets);
  var blob = xlsxZip_(parts, "施工日報彙整-自我測試.xlsx");
  var it = xlsxFolder_().getFilesByName("施工日報彙整-自我測試.xlsx");
  while (it.hasNext()) it.next().setTrashed(true);
  var file = xlsxFolder_().createFile(blob);
  var back = xlsxUnzip_(file.getBlob());
  var names = XlsxBuilder.readSheets(back).map(function (s) { return s.name; });
  var out = { ok: names.length === 2, sheets: names, parts: Object.keys(back).length, bytes: file.getSize(), url: file.getUrl() };
  Logger.log(JSON.stringify(out));
  return out;
}
