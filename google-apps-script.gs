/**
 * 貼到「員工休假管理表.xlsx」對應的 Google 試算表：
 * 選單「擴充功能」→「Apps Script」→ 把整份檔案內容貼進去覆蓋預設的 Code.gs → 存檔 → 部署。
 * 詳細部署步驟見 README.md「開放網頁直接填假」章節。
 */

// 找某欄第一個真正空白的列（從第2列開始）；欄位可能因為原始檔案預填了公式
// 導致 getLastRow() 不可靠，所以逐列檢查內容，而不是只看 getLastRow()。
function firstEmptyRow(sheet, col) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  var vals = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === "") return i + 2;
  }
  return lastRow + 1;
}

// 防止使用者輸入的字串被 Google Sheets 當成公式執行（CSV/公式注入）。
// 開頭是 = + - @ 或 tab/CR 的內容一律加前置單引號，強制存成純文字。
function sanitizeCell(v) {
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

// 檢查這個 clientId 是否已經寫入過（只查最近 50 筆，避免資料量大時逐列掃描太慢）。
// 用來擋前端自動重試造成的重複寫入：同一次送出不管重試幾次，clientId 都相同。
function isDuplicateSubmission(log, clientId) {
  var lastRow = log.getLastRow();
  if (lastRow < 2) return false;
  var checkFrom = Math.max(2, lastRow - 50);
  var ids = log.getRange(checkFrom, 7, lastRow - checkFrom + 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === clientId) return true;
  }
  return false;
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
    var name = sanitizeCell(String(data.name || "").trim());
    var start = String(data.start || "").trim();
    var end = String(data.end || "").trim();
    var type = String(data.type || "全天").trim();
    var note = sanitizeCell(String(data.note || "").trim());
    var password = String(data.password || "");
    var clientId = String(data.clientId || "").trim();

    // 通關密碼：實際密碼值存在「專案設定→指令碼屬性」的 SUBMIT_PASSWORD，
    // 不會出現在這份原始碼裡。如果還沒設定這個屬性，就不檢查密碼（相容舊行為）。
    var requiredPassword = PropertiesService.getScriptProperties().getProperty("SUBMIT_PASSWORD");
    if (requiredPassword && password !== requiredPassword) {
      return respond({ ok: false, error: "通關密碼錯誤，請向管理者確認" });
    }

    if (!name) return respond({ ok: false, error: "請填姓名" });
    if (!start || !end) return respond({ ok: false, error: "請選開始/結束日期" });
    if (new Date(end) < new Date(start)) {
      return respond({ ok: false, error: "結束日期不能早於開始日期" });
    }
    var validTypes = ["全天", "上午半天", "下午半天"];
    if (validTypes.indexOf(type) === -1) type = "全天";
    if (type !== "全天") end = start; // 半天假強制單日，避免被誤送成橫跨多天

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var roster = ss.getSheetByName("員工名單");
    var log = ss.getSheetByName("請假紀錄");
    if (!roster || !log) {
      return respond({ ok: false, error: "找不到「員工名單」或「請假紀錄」分頁，請確認分頁名稱" });
    }

    // 同一次送出（含前端自動重試）若已經寫入過，直接回成功，不要再寫一次
    if (clientId && isDuplicateSubmission(log, clientId)) {
      return respond({ ok: true, name: name, duplicate: true });
    }

    // 姓名不在名單就自動新增一列
    var lastRosterRow = roster.getLastRow();
    var names = lastRosterRow >= 2
      ? roster.getRange(2, 1, lastRosterRow - 1, 1).getValues().flat()
          .map(function (v) { return String(v).trim(); })
          .filter(Boolean)
      : [];
    if (names.indexOf(name) === -1) {
      roster.getRange(firstEmptyRow(roster, 1), 1).setValue(name);
    }

    // 寫入一筆請假紀錄，天數公式跟 xlsx 原本的邏輯一致
    var nextRow = firstEmptyRow(log, 1);
    log.getRange(nextRow, 1, 1, 5).setValues([[name, start, end, type, note]]);
    log.getRange(nextRow, 6).setFormula(
      "=IF(A" + nextRow + '="","",IF(D' + nextRow + '="全天",C' + nextRow + "-B" + nextRow + "+1,0.5))"
    );
    log.getRange(nextRow, 2, 1, 2).setNumberFormat("yyyy-mm-dd");
    if (clientId) log.getRange(nextRow, 7).setValue(clientId); // G欄：重複寫入防護用，勿手動編輯

    return respond({ ok: true, name: name });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 一次性設定：把「請假紀錄」分頁的「假別」欄位（D欄）變成下拉選單，只能選
// 全天／上午半天／下午半天，避免手動編輯試算表時打成「特休」「病假」等網頁
// 無法辨識的字串，導致那筆假期在月曆上悄悄消失、天數也不會被計入。
// 用法：在 Apps Script 編輯器上方選這個函式名稱，按「執行」一次即可
// （不影響已經填好的舊資料，只是加上輸入限制）。
function setupLeaveTypeValidation() {
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("請假紀錄");
  if (!log) throw new Error("找不到「請假紀錄」分頁");
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["全天", "上午半天", "下午半天"], true)
    .setAllowInvalid(false)
    .build();
  log.getRange(2, 4, 998, 1).setDataValidation(rule);
}

function doGet(e) {
  return respond({ ok: true, message: "leave-calendar API is running" });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
