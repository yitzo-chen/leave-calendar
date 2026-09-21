/**
 * 貼到「日報網頁填寫資料庫」試算表：
 * 選單「擴充功能」→「Apps Script」→ 把整份檔案內容貼進去覆蓋預設的 Code.gs → 存檔
 * → 先跑一次 setupSheets() 建立分頁結構 → 再依 README「部署」步驟部署成 Web App。
 */

var SHEET_LIST = "工種機具材料清單";
var SHEET_RECORD = "日報記錄";
var SHEET_HEADER = "日報頭";
var SHEET_BASIC = "基本資料";
var SHEET_ATTENDANCE = "本工出勤";

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
  record.getRange(2, 1, 998, 1).setNumberFormat("@"); // 日期欄強制純文字，避免Sheets自動轉成日期型別

  var header = ss.getSheetByName(SHEET_HEADER) || ss.insertSheet(SHEET_HEADER);
  // 舊版(11欄，含「加班人員上午/下午」)遷移成新版(9欄，加班改用獨立的「本工出勤」分頁)
  if (header.getRange(1, 6).getValue() === "加班人員(上午)") {
    header.getRange(1, 1, 1, 11).clearContent();
    header.getRange(1, 1).setValue(""); // 觸發下面的補建
  }
  if (header.getRange(1, 1).getValue() === "") {
    header.getRange(1, 1, 1, 9).setValues([[
      "日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目",
      "備註", "填表人", "clientId", "更新時間",
    ]]);
  }
  // 主任/工安的上午/下午姓名，補在最後4欄（不動既有欄位順序，舊資料不用遷移）
  if (header.getRange(1, 10).getValue() === "") {
    header.getRange(1, 10, 1, 4).setValues([["主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)"]]);
  }
  header.getRange(2, 1, 998, 1).setNumberFormat("@");

  var attendance = ss.getSheetByName(SHEET_ATTENDANCE) || ss.insertSheet(SHEET_ATTENDANCE);
  // 舊版(8欄「加班時數＋備註」)遷移成新版(9欄「上午加班＋下午加班＋加班原因」)。
  // 只重寫表頭，舊資料列不搬移；正式資料庫此分頁目前是空的，若有舊資料需自行處理。
  if (attendance.getRange(1, 5).getValue() === "加班時數") {
    attendance.getRange(1, 1, 1, 9).clearContent();
  }
  if (attendance.getRange(1, 1).getValue() === "") {
    attendance.getRange(1, 1, 1, 9).setValues([[
      "日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間",
    ]]);
  }
  attendance.getRange(2, 1, 998, 1).setNumberFormat("@");
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

// 沒設定 SUBMIT_PASSWORD 指令碼屬性時一律放行(測試期間)；設定了才擋密碼+錯誤鎖定。
function verifyPassword(password) {
  var requiredPassword = PropertiesService.getScriptProperties().getProperty("SUBMIT_PASSWORD");
  if (!requiredPassword) return { ok: true };
  if (isSubmitLocked()) {
    return { ok: false, error: "密碼錯誤次數過多，已暫時鎖定送出功能，請 10 分鐘後再試" };
  }
  if (password !== requiredPassword) {
    recordWrongPassword();
    return { ok: false, error: "通關密碼錯誤，請向管理者確認" };
  }
  clearWrongPasswordCount();
  return { ok: true };
}

// Sheets 寫入「2026-09-18」這種字串時常會自動轉成真正的日期型別，
// 下次 getValues() 讀出來就變成 Date 物件而不是字串，直接比對字串永遠對不上。
// 所有跟「日期」欄位有關的比對都要先過這道正規化。
function normalizeDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v || "").trim();
}

// ---------- 找日期欄位等於指定值的列（從第2列開始），找不到回傳 -1 ----------
function findRowByKey(sheet, keyCol, keyVal) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (normalizeDate(vals[i][0]) === normalizeDate(keyVal)) return i + 2;
  }
  return -1;
}

// 找「日期+第二欄」都相符的列（日報記錄/本工出勤共用，兩欄複合鍵）
function findRecordRow(sheet, date, key2) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (normalizeDate(vals[i][0]) === date && String(vals[i][1]).trim() === key2) return i + 2;
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
    var action = data.action || "submit";
    var pwCheck = verifyPassword(String(data.password || ""));
    if (!pwCheck.ok) return respond(pwCheck);

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- 新增工種/機具/材料項目（給表單上的「＋新增項目」按鈕用） ----
    if (action === "addItem") {
      var category = String(data.category || "").trim();
      var name = String(data.name || "").trim();
      if (["工種", "機具", "材料"].indexOf(category) === -1) {
        return respond({ ok: false, error: "類別錯誤" });
      }
      if (!name) return respond({ ok: false, error: "請輸入項目名稱" });

      var listSheet = ss.getSheetByName(SHEET_LIST);
      var listRows = listSheet.getDataRange().getValues();
      for (var i = 1; i < listRows.length; i++) {
        if (String(listRows[i][1]).trim() === name) {
          if (String(listRows[i][3]).trim().toUpperCase() === "TRUE") {
            return respond({ ok: false, error: "項目已存在" });
          }
          listSheet.getRange(i + 1, 4).setValue("TRUE"); // 曾被停用的項目：重新啟用
          return respond({ ok: true });
        }
      }
      listSheet.appendRow([category, sanitizeCell(name), "", "TRUE"]);
      return respond({ ok: true });
    }

    // ---- 送出／更新當天日報（原本邏輯） ----
    var date = String(data.date || "").trim();
    var clientId = sanitizeCell(String(data.clientId || "").trim());
    var h = data.header || {};
    var items = Array.isArray(data.items) ? data.items : [];
    var attendance = Array.isArray(data.attendance) ? data.attendance : [];
    // 使用者在畫面上「明確」清空/移除的項目與人員（前端比對載入時的資料算出來）。
    // 只有這兩份名單裡的資料才會被刪除，沒被列出的既有資料一律不動，避免舊頁面或兩人同天分填時互相蓋掉。
    var clearItems = Array.isArray(data.clearItems) ? data.clearItems : [];
    var removeAttendance = Array.isArray(data.removeAttendance) ? data.removeAttendance : [];

    if (!date) return respond({ ok: false, error: "缺少日期" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return respond({ ok: false, error: "日期格式錯誤" });

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
      sanitizeCell(String(h.remark || "")),
      sanitizeCell(String(h.reporter || "")),
      clientId,
      now,
      sanitizeCell(String(h.directorAm || "")),
      sanitizeCell(String(h.directorPm || "")),
      sanitizeCell(String(h.safetyAm || "")),
      sanitizeCell(String(h.safetyPm || "")),
    ];
    if (hRow === -1) hRow = headerSheet.getLastRow() + 1;
    headerSheet.getRange(hRow, 1, 1, hVals.length).setValues([hVals]);

    // ---- 日報記錄：先刪除使用者明確清空的項目，再逐一 upsert（同一天同一項目已有資料就覆蓋，沒有就新增） ----
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    for (var c = 0; c < clearItems.length; c++) {
      var clearName = String(clearItems[c] || "").trim();
      if (!clearName) continue;
      var cRow = findRecordRow(recordSheet, date, clearName);
      if (cRow !== -1) recordSheet.deleteRow(cRow);
    }
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      var itName = String(it.name || "").trim();
      if (!itName) continue;
      var am = Number(it.am || 0);
      var pm = Number(it.pm || 0);
      if (am === 0 && pm === 0) continue; // 沒填數字的項目不寫入，避免資料表被灌滿0
      var rRow = findRecordRow(recordSheet, date, itName);
      var rVals = [date, sanitizeCell(itName), am, pm, clientId, now];
      if (rRow === -1) rRow = recordSheet.getLastRow() + 1;
      recordSheet.getRange(rRow, 1, 1, rVals.length).setValues([rVals]);
    }

    // ---- 本工出勤：先刪除使用者明確移除的人，再逐一 upsert（同一天同一人已有資料就覆蓋，沒有就新增） ----
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    for (var r = 0; r < removeAttendance.length; r++) {
      var removeName = String(removeAttendance[r] || "").trim();
      if (!removeName) continue;
      var rmRow = findRecordRow(attSheet, date, removeName);
      if (rmRow !== -1) attSheet.deleteRow(rmRow);
    }
    for (var k = 0; k < attendance.length; k++) {
      var a = attendance[k];
      var aName = String(a.name || "").trim();
      if (!aName) continue;
      var aRow = findRecordRow(attSheet, date, aName);
      var aVals = [
        date, sanitizeCell(aName),
        a.am ? "V" : "", a.pm ? "V" : "",
        Number(a.amHours || 0), Number(a.pmHours || 0),
        sanitizeCell(String(a.reason || "")),
        clientId, now,
      ];
      if (aRow === -1) aRow = attSheet.getLastRow() + 1;
      attSheet.getRange(aRow, 1, 1, aVals.length).setValues([aVals]);
    }

    // ---- 更新雲端 xlsx（xlsx-drive.gs）：已持有 ScriptLock，內部不再取鎖；失敗不影響日報送出 ----
    var xlsx;
    try {
      var xr = xlsxUpdateForDate(date);
      xlsx = { ok: true, url: xr.url, warnings: xr.warnings || [] };
    } catch (xerr) {
      xlsx = { ok: false, error: String(xerr) };
    }

    return respond({ ok: true, date: date, xlsx: xlsx });
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
      var hv = headerSheet.getRange(hRow, 1, 1, 13).getValues()[0];
      header = {
        date: hv[0], weather: hv[1], status: hv[2], todayWork: hv[3], tomorrowPlan: hv[4],
        remark: hv[5], reporter: hv[6],
        directorAm: hv[9], directorPm: hv[10], safetyAm: hv[11], safetyPm: hv[12],
      };
    }
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    var lastRow = recordSheet.getLastRow();
    var records = [];
    if (lastRow >= 2) {
      var rv = recordSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      records = rv.filter(function (r) { return normalizeDate(r[0]) === date; })
        .map(function (r) { return { name: r[1], am: r[2], pm: r[3] }; });
    }
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    var attLastRow = attSheet.getLastRow();
    var attendance = [];
    if (attLastRow >= 2) {
      var av = attSheet.getRange(2, 1, attLastRow - 1, 7).getValues();
      attendance = av.filter(function (r) { return normalizeDate(r[0]) === date; })
        .map(function (r) { return { name: r[1], am: r[2] === "V", pm: r[3] === "V", amHours: r[4], pmHours: r[5], reason: r[6] }; });
    }
    return respond({ ok: true, header: header, records: records, attendance: attendance });
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
        var d = normalizeDate(r[0]);
        if (upTo && d > upTo) return;
        var name = r[1];
        var sum = Number(r[2] || 0) + Number(r[3] || 0);
        var value = categoryByName[name] === "材料" ? sum : sum / 2;
        totals[name] = (totals[name] || 0) + value;
      });
    }
    return respond({ ok: true, totals: totals });
  }

  if (action === "reporters") {
    var headerSheet3 = ss.getSheetByName(SHEET_HEADER);
    var lastRow3 = headerSheet3.getLastRow();
    var names = [];
    if (lastRow3 >= 2) {
      var seen = {};
      headerSheet3.getRange(2, 7, lastRow3 - 1, 1).getValues().forEach(function (r) {
        var n = String(r[0]).trim();
        if (n && !seen[n]) { seen[n] = true; names.push(n); }
      });
    }
    return respond({ ok: true, names: names });
  }

  if (action === "peopleNames") {
    var attSheet3 = ss.getSheetByName(SHEET_ATTENDANCE);
    var attLastRow3 = attSheet3.getLastRow();
    var pNames = [];
    if (attLastRow3 >= 2) {
      var seen2 = {};
      attSheet3.getRange(2, 2, attLastRow3 - 1, 1).getValues().forEach(function (r) {
        var n = String(r[0]).trim();
        if (n && !seen2[n]) { seen2[n] = true; pNames.push(n); }
      });
    }
    return respond({ ok: true, names: pNames });
  }

  return respond({ ok: true, message: "daily-report API is running" });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
