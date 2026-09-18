/**
 * 貼到「日報網頁填寫資料庫」試算表：
 * 選單「擴充功能」→「Apps Script」→ 把整份檔案內容貼進去覆蓋預設的 Code.gs → 存檔
 * → 先跑一次 setupSheets() 建立分頁結構 → 再依 README「部署」步驟部署成 Web App。
 */

var SHEET_LIST = "工種機具材料清單";
var SHEET_RECORD = "日報記錄";
var SHEET_HEADER = "日報頭";
var SHEET_BASIC = "基本資料";

// ---------- 一次性設定：建立所有分頁結構（重複執行不會清空既有資料，只補齊缺的分頁/表頭） ----------
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var basic = ss.getSheetByName(SHEET_BASIC) || ss.insertSheet(SHEET_BASIC);
  if (basic.getRange(1, 1).getValue() === "") {
    basic.getRange(1, 1, 5, 2).setValues([
      ["業主", ""],
      ["工程名稱", "洲際液化天然氣接收站"],
      ["合約金額（元）", ""],
      ["開工日期（YYYY/MM/DD）", ""],
      ["公司名稱", ""],
    ]);
  }

  var list = ss.getSheetByName(SHEET_LIST) || ss.insertSheet(SHEET_LIST);
  if (list.getRange(1, 1).getValue() === "") {
    list.getRange(1, 1, 1, 4).setValues([["類別", "項目名稱", "工項編號", "啟用中"]]);
    list.getRange(2, 1, SEED_ITEMS.length, 4).setValues(SEED_ITEMS);
    var activeRule = SpreadsheetApp.newDataValidation().requireValueInList(["TRUE", "FALSE"], true).build();
    list.getRange(2, 4, 500, 1).setDataValidation(activeRule);
    var catRule = SpreadsheetApp.newDataValidation().requireValueInList(["工種", "機具", "材料"], true).build();
    list.getRange(2, 1, 500, 1).setDataValidation(catRule);
  }

  var record = ss.getSheetByName(SHEET_RECORD) || ss.insertSheet(SHEET_RECORD);
  if (record.getRange(1, 1).getValue() === "") {
    record.getRange(1, 1, 1, 6).setValues([["日期", "項目名稱", "上午", "下午", "clientId", "更新時間"]]);
  }

  var header = ss.getSheetByName(SHEET_HEADER) || ss.insertSheet(SHEET_HEADER);
  if (header.getRange(1, 1).getValue() === "") {
    header.getRange(1, 1, 1, 11).setValues([[
      "日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目",
      "加班人員(上午)", "加班人員(下午)", "備註", "填表人", "clientId", "更新時間",
    ]]);
  }
}

// 工種/機具/材料清單的起始種子資料，抄自現行 Excel 日報範本的預設項目。
// 啟用中預設 TRUE；工項編號留空，之後要接成本管理系統時再填。
var SEED_ITEMS = [
  ["工種", "公司工", "", "TRUE"],
  ["工種", "鋼筋工(公司)", "", "TRUE"],
  ["工種", "鋼筋工(工地)", "", "TRUE"],
  ["工種", "模板工", "", "TRUE"],
  ["工種", "搭架工", "", "TRUE"],
  ["工種", "泥作工", "", "TRUE"],
  ["工種", "電銲工", "", "TRUE"],
  ["工種", "水電工", "", "TRUE"],
  ["機具", "挖土機(50型)", "", "TRUE"],
  ["機具", "挖土機(120型)", "", "TRUE"],
  ["機具", "運土車", "", "TRUE"],
  ["機具", "吊卡車", "", "TRUE"],
  ["機具", "吊車", "", "TRUE"],
  ["機具", "壓送車", "", "TRUE"],
  ["材料", "無收縮(包)", "", "TRUE"],
  ["材料", "140kg/cm²混凝土(m³)", "", "TRUE"],
  ["材料", "210kg/cm²混凝土(m³)", "", "TRUE"],
  ["材料", "280kg/cm²混凝土(m³)", "", "TRUE"],
  ["材料", "350kg/cm²混凝土(m³)", "", "TRUE"],
  ["材料", "低強度混凝土(m³)", "", "TRUE"],
  ["材料", "砂(m³)", "", "TRUE"],
  ["材料", "水泥(包)", "", "TRUE"],
  ["材料", "黏著劑(包)", "", "TRUE"],
];

// ---------- 密碼錯誤鎖定（跟休假月曆同一套邏輯） ----------
var LOCK_THRESHOLD = 5;
var LOCK_WINDOW_SECONDS = 600;

function isSubmitLocked() {
  return !!CacheService.getScriptCache().get("submitLocked");
}
function recordWrongPassword() {
  var cache = CacheService.getScriptCache();
  var count = Number(cache.get("wrongPasswordCount") || "0") + 1;
  if (count >= LOCK_THRESHOLD) {
    cache.put("submitLocked", "1", LOCK_WINDOW_SECONDS);
    cache.remove("wrongPasswordCount");
  } else {
    cache.put("wrongPasswordCount", String(count), LOCK_WINDOW_SECONDS);
  }
}
function clearWrongPasswordCount() {
  CacheService.getScriptCache().remove("wrongPasswordCount");
}

// ---------- 找某個 key 欄位等於指定值的列（從第2列開始），找不到回傳 -1 ----------
function findRowByKey(sheet, keyCol, keyVal) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(keyVal).trim()) return i + 2;
  }
  return -1;
}

// 找「日期+項目名稱」都相符的列（日報記錄用，兩欄複合鍵）
function findRecordRow(sheet, date, itemName) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === date && String(vals[i][1]).trim() === itemName) return i + 2;
  }
  return -1;
}

function sanitizeCell(v) {
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return respond({ ok: false, error: "系統忙線中，請稍後再試" });
  }
  try {
    var data = JSON.parse(e.postData.contents);
    var date = String(data.date || "").trim();
    var password = String(data.password || "");
    var clientId = String(data.clientId || "").trim();
    var h = data.header || {};
    var items = Array.isArray(data.items) ? data.items : [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var requiredPassword = PropertiesService.getScriptProperties().getProperty("SUBMIT_PASSWORD");
    if (requiredPassword) {
      if (isSubmitLocked()) {
        return respond({ ok: false, error: "密碼錯誤次數過多，已暫時鎖定送出功能，請 10 分鐘後再試" });
      }
      if (password !== requiredPassword) {
        recordWrongPassword();
        return respond({ ok: false, error: "通關密碼錯誤，請向管理者確認" });
      }
      clearWrongPasswordCount();
    }

    if (!date) return respond({ ok: false, error: "缺少日期" });

    var now = new Date();

    // ---- upsert 日報頭：同一天已有資料就覆蓋，沒有就新增 ----
    var headerSheet = ss.getSheetByName(SHEET_HEADER);
    var hRow = findRowByKey(headerSheet, 1, date);
    var hVals = [
      date,
      sanitizeCell(String(h.weather || "")),
      sanitizeCell(String(h.status || "")),
      sanitizeCell(String(h.todayWork || "")),
      sanitizeCell(String(h.tomorrowPlan || "")),
      sanitizeCell(String(h.overtimeAM || "")),
      sanitizeCell(String(h.overtimePM || "")),
      sanitizeCell(String(h.remark || "")),
      sanitizeCell(String(h.reporter || "")),
      clientId,
      now,
    ];
    if (hRow === -1) hRow = headerSheet.getLastRow() + 1;
    headerSheet.getRange(hRow, 1, 1, hVals.length).setValues([hVals]);

    // ---- upsert 日報記錄：逐一項目，同一天同一項目已有資料就覆蓋，沒有就新增 ----
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var name = String(it.name || "").trim();
      if (!name) continue;
      var am = Number(it.am || 0);
      var pm = Number(it.pm || 0);
      if (am === 0 && pm === 0) continue; // 沒填數字的項目不寫入，避免資料表被灌滿0
      var rRow = findRecordRow(recordSheet, date, name);
      var rVals = [date, name, am, pm, clientId, now];
      if (rRow === -1) rRow = recordSheet.getLastRow() + 1;
      recordSheet.getRange(rRow, 1, 1, rVals.length).setValues([rVals]);
    }

    return respond({ ok: true, date: date });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = (e.parameter.action || "").trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "config") {
    var list = ss.getSheetByName(SHEET_LIST);
    var rows = list.getDataRange().getValues();
    rows.shift(); // 表頭
    var items = rows
      .filter(function (r) { return String(r[3]).trim().toUpperCase() === "TRUE"; })
      .map(function (r) { return { category: r[0], name: r[1], workCode: r[2] }; });
    var basic = ss.getSheetByName(SHEET_BASIC).getRange(1, 1, 5, 2).getValues();
    var basicObj = {};
    basic.forEach(function (r) { basicObj[r[0]] = r[1]; });
    return respond({ ok: true, items: items, basic: basicObj });
  }

  if (action === "day") {
    var date = String(e.parameter.date || "").trim();
    var headerSheet = ss.getSheetByName(SHEET_HEADER);
    var hRow = findRowByKey(headerSheet, 1, date);
    var header = null;
    if (hRow !== -1) {
      var hv = headerSheet.getRange(hRow, 1, 1, 9).getValues()[0];
      header = {
        date: hv[0], weather: hv[1], status: hv[2], todayWork: hv[3], tomorrowPlan: hv[4],
        overtimeAM: hv[5], overtimePM: hv[6], remark: hv[7], reporter: hv[8],
      };
    }
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    var lastRow = recordSheet.getLastRow();
    var records = [];
    if (lastRow >= 2) {
      var rv = recordSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      records = rv.filter(function (r) { return String(r[0]).trim() === date; })
        .map(function (r) { return { name: r[1], am: r[2], pm: r[3] }; });
    }
    return respond({ ok: true, header: header, records: records });
  }

  if (action === "cumulative") {
    var upTo = String(e.parameter.upTo || "").trim();

    // 材料是直接加總用量，工種/機具是「工天數」邏輯(上午+下午)/2 —— 兩種算法不同，
    // 要先查清單分頁知道每個項目屬於哪一類，才能套對公式。
    var listSheet = ss.getSheetByName(SHEET_LIST);
    var listRows = listSheet.getDataRange().getValues();
    listRows.shift();
    var categoryByName = {};
    listRows.forEach(function (r) { categoryByName[r[1]] = r[0]; });

    var recordSheet2 = ss.getSheetByName(SHEET_RECORD);
    var lastRow2 = recordSheet2.getLastRow();
    var totals = {};
    if (lastRow2 >= 2) {
      var rv2 = recordSheet2.getRange(2, 1, lastRow2 - 1, 4).getValues();
      rv2.forEach(function (r) {
        var d = String(r[0]).trim();
        if (upTo && d > upTo) return;
        var name = r[1];
        var sum = Number(r[2] || 0) + Number(r[3] || 0);
        var value = categoryByName[name] === "材料" ? sum : sum / 2;
        totals[name] = (totals[name] || 0) + value;
      });
    }
    return respond({ ok: true, totals: totals });
  }

  return respond({ ok: true, message: "daily-report API is running" });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
