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
  setDateColumnText(record); // 日期欄強制純文字，避免Sheets自動轉成日期型別

  var header = ss.getSheetByName(SHEET_HEADER) || ss.insertSheet(SHEET_HEADER);
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
  setDateColumnText(header);

  var attendance = ss.getSheetByName(SHEET_ATTENDANCE) || ss.insertSheet(SHEET_ATTENDANCE);
  if (attendance.getRange(1, 1).getValue() === "") {
    attendance.getRange(1, 1, 1, 9).setValues([[
      "日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間",
    ]]);
  }
  setDateColumnText(attendance);
}

// 日期欄（A 欄）第 2 列起到分頁目前的最後一列全設純文字（不再只涵蓋前 999 列）。
// 資料列成長超過目前列數時，新增的列由 Sheets 沿用相鄰列格式；normalizeDate 仍保留作為最後防線。
function setDateColumnText(sheet) {
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rows, 1).setNumberFormat("@");
}

// 寫入前確保分頁列數足夠：真實 Sheets 對超出列數的 setValues 會丟例外（只有 appendRow 會自動擴充），
// 日報記錄約 70~100 天就會用完預設的 1000 列。新增的列沿用上方列的格式（含日期欄純文字）。
function ensureRows(sheet, row) {
  var max = sheet.getMaxRows();
  if (row > max) sheet.insertRowsAfter(max, row - max + 200);
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

var LOCKED_ERROR = "密碼錯誤次數過多，已暫時鎖定送出功能，請 10 分鐘後再試";

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
    return { ok: false, error: LOCKED_ERROR };
  }
  if (password !== requiredPassword) {
    recordWrongPassword();
    // 觸發鎖定的那一次（第 LOCK_THRESHOLD 次錯誤）就直接告知已鎖定，不用等下一次
    if (isSubmitLocked()) return { ok: false, error: LOCKED_ERROR };
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
    if (normalizeDate(vals[i][0]) === date && unsanitizeCell(vals[i][1]).trim() === key2) return i + 2;
  }
  return -1;
}

// 寫入儲存格前的公式注入防護：以 = + - @ Tab CR 開頭的字串前面加單引號，讓 Sheets 當純文字。
function sanitizeCell(v) {
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

// sanitizeCell 的反向：讀出、比對名稱時先去掉防護用的前導單引號，還原使用者輸入的原字串。
// 儲存格若被 Sheets 吃掉單引號（本來就沒有前綴）則原樣返回，兩種儲存狀態都能一致比對。
// 限制：使用者原本輸入就是「'=X」這種（單引號 + 危險字元開頭）時，讀出會少一個單引號，可接受。
function unsanitizeCell(v) {
  var s = String(v === null || v === undefined ? "" : v);
  return /^'[=+\-@\t\r]/.test(s) ? s.slice(1) : s;
}

// ---------- 送出資料驗證：在寫入任何資料之前先全部驗完，任何不合法就整個拒絕（全有或全無） ----------
var MAX_LONG_TEXT = 2000;  // 備註、本日/明日施工項目
var MAX_SHORT_TEXT = 100;  // 項目名稱、人員姓名、加班原因、天氣、施工狀況、填表人、clientId 等

function validationError(msg) {
  var err = new Error(msg);
  err.isValidation = true;
  return err;
}

// 文字欄位：空值(null/undefined/""/0/false)視為空字串；物件/陣列拒絕(會變成 [object Object])；超過長度上限拒絕。
// trim=true 時先去前後空白再算長度並回傳 trim 後的值（名稱/姓名用）；其餘欄位保留原樣。
function checkText(v, label, max, trim) {
  if (!v) return "";
  if (typeof v === "object" || typeof v === "function") throw validationError(label + "格式錯誤（必須是文字）");
  var s = String(v);
  if (trim) s = s.trim();
  if (s.length > max) throw validationError(label + "過長（上限 " + max + " 字，目前 " + s.length + " 字）");
  return s;
}

// 數量/時數欄位：空值(null/undefined/"")視為 0；非數字、無限大、負數拒絕。
function checkQty(v, label) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v !== "number" && typeof v !== "string") throw validationError(label + "格式錯誤（必須是數字）");
  var n = (typeof v === "string" && v.trim() === "") ? 0 : Number(v);
  if (!isFinite(n)) throw validationError(label + "不是有效的數字");
  if (n < 0) throw validationError(label + "不可為負數");
  return n;
}

// 找出第一個不存在的分頁（尚未執行 setupSheets 時給友善錯誤，而不是丟例外）
function firstMissingSheet(ss, names) {
  for (var i = 0; i < names.length; i++) {
    if (!ss.getSheetByName(names[i])) return names[i];
  }
  return null;
}
var NOT_SETUP_ERROR = "尚未初始化，請先執行 setupSheets";

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
    if (action !== "submit" && action !== "addItem") return respond({ ok: false, error: "不支援的操作" });

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- 新增工種/機具/材料項目（給表單上的「＋新增項目」按鈕用） ----
    if (action === "addItem") {
      var category = String(data.category || "").trim();
      var name = checkText(data.name, "項目名稱", MAX_SHORT_TEXT, true);
      if (["工種", "機具", "材料"].indexOf(category) === -1) {
        return respond({ ok: false, error: "類別錯誤" });
      }
      if (!name) return respond({ ok: false, error: "請輸入項目名稱" });
      if (firstMissingSheet(ss, [SHEET_LIST])) return respond({ ok: false, error: NOT_SETUP_ERROR });

      var listSheet = ss.getSheetByName(SHEET_LIST);
      var listRows = listSheet.getDataRange().getValues();
      for (var i = 1; i < listRows.length; i++) {
        if (unsanitizeCell(listRows[i][1]).trim() === name) {
          if (String(listRows[i][3]).trim().toUpperCase() === "TRUE") {
            return respond({ ok: false, error: "項目已存在" });
          }
          listSheet.getRange(i + 1, 1).setValue(category); // 曾被停用的項目：以新傳入的類別更新並重新啟用
          listSheet.getRange(i + 1, 4).setValue("TRUE");
          return respond({ ok: true });
        }
      }
      listSheet.appendRow([category, sanitizeCell(name), "", "TRUE"]);
      return respond({ ok: true });
    }

    // ---- 送出／更新當天日報（原本邏輯） ----
    var date = checkText(data.date, "日期", 20, true);
    var clientId = sanitizeCell(checkText(data.clientId, "clientId", MAX_SHORT_TEXT, true));
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

    // ---- 寫入前先完整驗證所有內容（任何不合法就整個拒絕，之後才開始寫入） ----
    var hVals = [
      date,
      sanitizeCell(checkText(h.weather, "天氣", MAX_SHORT_TEXT)),
      sanitizeCell(checkText(h.status, "施工狀況", MAX_SHORT_TEXT)),
      sanitizeCell(checkText(h.todayWork, "本日施工項目", MAX_LONG_TEXT)),
      sanitizeCell(checkText(h.tomorrowPlan, "預計明日施工項目", MAX_LONG_TEXT)),
      sanitizeCell(checkText(h.remark, "備註", MAX_LONG_TEXT)),
      sanitizeCell(checkText(h.reporter, "填表人", MAX_SHORT_TEXT)),
      clientId,
      now,
      sanitizeCell(checkText(h.directorAm, "主任(上午)", MAX_SHORT_TEXT)),
      sanitizeCell(checkText(h.directorPm, "主任(下午)", MAX_SHORT_TEXT)),
      sanitizeCell(checkText(h.safetyAm, "工安(上午)", MAX_SHORT_TEXT)),
      sanitizeCell(checkText(h.safetyPm, "工安(下午)", MAX_SHORT_TEXT)),
    ];

    var clearNames = [];
    for (var c = 0; c < clearItems.length; c++) {
      var clearName = checkText(clearItems[c], "清除的項目名稱", MAX_SHORT_TEXT, true);
      if (clearName) clearNames.push(clearName);
    }
    var cleanItems = [], seenItems = Object.create(null), dupItems = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      if (!it || typeof it !== "object" || Array.isArray(it)) throw validationError("第 " + (j + 1) + " 筆項目資料格式錯誤");
      var itName = checkText(it.name, "項目名稱", MAX_SHORT_TEXT, true);
      if (!itName) continue;
      var am = checkQty(it.am, "項目「" + itName + "」的上午數量");
      var pm = checkQty(it.pm, "項目「" + itName + "」的下午數量");
      if (am === 0 && pm === 0) continue; // 沒填數字的項目不寫入，避免資料表被灌滿0
      if (seenItems[itName] && dupItems.indexOf(itName) === -1) dupItems.push(itName);
      seenItems[itName] = true;
      cleanItems.push({ name: itName, am: am, pm: pm });
    }
    if (dupItems.length) throw validationError("項目重複：" + dupItems.join("、"));

    var removeNames = [];
    for (var r = 0; r < removeAttendance.length; r++) {
      var removeName = checkText(removeAttendance[r], "移除的人員姓名", MAX_SHORT_TEXT, true);
      if (removeName) removeNames.push(removeName);
    }
    var cleanAtt = [], seenNames = Object.create(null), dupNames = [];
    for (var k = 0; k < attendance.length; k++) {
      var a = attendance[k];
      if (!a || typeof a !== "object" || Array.isArray(a)) throw validationError("第 " + (k + 1) + " 筆出勤資料格式錯誤");
      var aName = checkText(a.name, "人員姓名", MAX_SHORT_TEXT, true);
      if (!aName) continue;
      if (seenNames[aName] && dupNames.indexOf(aName) === -1) dupNames.push(aName);
      seenNames[aName] = true;
      cleanAtt.push({
        name: aName, am: !!a.am, pm: !!a.pm,
        amHours: checkQty(a.amHours, "「" + aName + "」的上午加班時數"),
        pmHours: checkQty(a.pmHours, "「" + aName + "」的下午加班時數"),
        reason: checkText(a.reason, "「" + aName + "」的加班原因", MAX_SHORT_TEXT),
      });
    }
    if (dupNames.length) throw validationError("本工出勤人員姓名重複：" + dupNames.join("、"));

    if (firstMissingSheet(ss, [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE])) return respond({ ok: false, error: NOT_SETUP_ERROR });

    // ---- upsert 日報頭：同一天已有資料就覆蓋，沒有就新增 ----
    var headerSheet = ss.getSheetByName(SHEET_HEADER);
    var hRow = findRowByKey(headerSheet, 1, date);
    if (hRow === -1) hRow = headerSheet.getLastRow() + 1;
    ensureRows(headerSheet, hRow);
    headerSheet.getRange(hRow, 1, 1, hVals.length).setValues([hVals]);

    // ---- 日報記錄：先刪除使用者明確清空的項目，再逐一 upsert（同一天同一項目已有資料就覆蓋，沒有就新增） ----
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    for (var c2 = 0; c2 < clearNames.length; c2++) {
      var cRow = findRecordRow(recordSheet, date, clearNames[c2]);
      if (cRow !== -1) recordSheet.deleteRow(cRow);
    }
    for (var j2 = 0; j2 < cleanItems.length; j2++) {
      var ci = cleanItems[j2];
      var rRow = findRecordRow(recordSheet, date, ci.name);
      var rVals = [date, sanitizeCell(ci.name), ci.am, ci.pm, clientId, now];
      if (rRow === -1) rRow = recordSheet.getLastRow() + 1;
      ensureRows(recordSheet, rRow);
      recordSheet.getRange(rRow, 1, 1, rVals.length).setValues([rVals]);
    }

    // ---- 本工出勤：先刪除使用者明確移除的人，再逐一 upsert（同一天同一人已有資料就覆蓋，沒有就新增） ----
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    for (var r2 = 0; r2 < removeNames.length; r2++) {
      var rmRow = findRecordRow(attSheet, date, removeNames[r2]);
      if (rmRow !== -1) attSheet.deleteRow(rmRow);
    }
    for (var k2 = 0; k2 < cleanAtt.length; k2++) {
      var ca = cleanAtt[k2];
      var aRow = findRecordRow(attSheet, date, ca.name);
      var aVals = [
        date, sanitizeCell(ca.name),
        ca.am ? "V" : "", ca.pm ? "V" : "",
        ca.amHours, ca.pmHours,
        sanitizeCell(ca.reason),
        clientId, now,
      ];
      if (aRow === -1) aRow = attSheet.getLastRow() + 1;
      ensureRows(attSheet, aRow);
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
    if (err && err.isValidation) return respond({ ok: false, error: err.message });
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = (e.parameter.action || "").trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 尚未執行 setupSheets（缺分頁）時回友善錯誤，不丟例外（前端 res.json() 會壞）
  var neededSheets = {
    config: [SHEET_LIST, SHEET_BASIC],
    day: [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE],
    cumulative: [SHEET_LIST, SHEET_RECORD],
    reporters: [SHEET_HEADER],
    peopleNames: [SHEET_ATTENDANCE],
  };
  if (Object.prototype.hasOwnProperty.call(neededSheets, action) && firstMissingSheet(ss, neededSheets[action])) {
    return respond({ ok: false, error: NOT_SETUP_ERROR });
  }

  if (action === "config") {
    var list = ss.getSheetByName(SHEET_LIST);
    var rows = list.getDataRange().getValues();
    rows.shift(); // 表頭
    var items = rows
      .filter(function (r) { return String(r[3]).trim().toUpperCase() === "TRUE"; })
      .map(function (r) { return { category: r[0], name: unsanitizeCell(r[1]), workCode: r[2] }; });
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
        date: hv[0], weather: unsanitizeCell(hv[1]), status: unsanitizeCell(hv[2]), todayWork: unsanitizeCell(hv[3]), tomorrowPlan: unsanitizeCell(hv[4]),
        remark: unsanitizeCell(hv[5]), reporter: unsanitizeCell(hv[6]),
        directorAm: unsanitizeCell(hv[9]), directorPm: unsanitizeCell(hv[10]), safetyAm: unsanitizeCell(hv[11]), safetyPm: unsanitizeCell(hv[12]),
      };
    }
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    var lastRow = recordSheet.getLastRow();
    var records = [];
    if (lastRow >= 2) {
      var rv = recordSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      records = rv.filter(function (r) { return normalizeDate(r[0]) === date; })
        .map(function (r) { return { name: unsanitizeCell(r[1]), am: r[2], pm: r[3] }; });
    }
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    var attLastRow = attSheet.getLastRow();
    var attendance = [];
    if (attLastRow >= 2) {
      var av = attSheet.getRange(2, 1, attLastRow - 1, 7).getValues();
      attendance = av.filter(function (r) { return normalizeDate(r[0]) === date; })
        .map(function (r) { return { name: unsanitizeCell(r[1]), am: r[2] === "V", pm: r[3] === "V", amHours: r[4], pmHours: r[5], reason: unsanitizeCell(r[6]) }; });
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
    var categoryByName = Object.create(null); // 項目名稱當 key：用無原型物件，避免 constructor/__proto__ 等名稱汙染
    listRows.forEach(function (r) { categoryByName[unsanitizeCell(r[1])] = r[0]; });

    var recordSheet2 = ss.getSheetByName(SHEET_RECORD);
    var lastRow2 = recordSheet2.getLastRow();
    var totals = Object.create(null);
    var cumWarnings = [], skippedNames = Object.create(null);
    if (lastRow2 >= 2) {
      var rv2 = recordSheet2.getRange(2, 1, lastRow2 - 1, 4).getValues();
      rv2.forEach(function (r) {
        var d = normalizeDate(r[0]);
        if (upTo && d > upTo) return;
        var name = unsanitizeCell(r[1]);
        if (!(name in categoryByName)) { // 項目不在清單：不知道該用哪種算法，略過並警告
          if (!skippedNames[name]) { skippedNames[name] = true; cumWarnings.push("累計略過不在清單的項目：" + name); }
          return;
        }
        var sum = Number(r[2] || 0) + Number(r[3] || 0);
        var value = categoryByName[name] === "材料" ? sum : sum / 2;
        totals[name] = (totals[name] || 0) + value;
      });
    }
    return respond({ ok: true, totals: totals, warnings: cumWarnings });
  }

  if (action === "reporters") {
    var headerSheet3 = ss.getSheetByName(SHEET_HEADER);
    var lastRow3 = headerSheet3.getLastRow();
    var names = [];
    if (lastRow3 >= 2) {
      var seen = Object.create(null);
      headerSheet3.getRange(2, 7, lastRow3 - 1, 1).getValues().forEach(function (r) {
        var n = unsanitizeCell(r[0]).trim();
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
      var seen2 = Object.create(null);
      attSheet3.getRange(2, 2, attLastRow3 - 1, 1).getValues().forEach(function (r) {
        var n = unsanitizeCell(r[0]).trim();
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
