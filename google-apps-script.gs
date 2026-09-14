/**
 * 貼到「員工休假管理表.xlsx」對應的 Google 試算表：
 * 選單「擴充功能」→「Apps Script」→ 把整份檔案內容貼進去覆蓋預設的 Code.gs → 存檔 → 部署。
 * 詳細部署步驟見 README.md「開放網頁直接填假」章節。
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var name = String(data.name || "").trim();
    var start = String(data.start || "").trim();
    var end = String(data.end || "").trim();
    var type = String(data.type || "全天").trim();
    var note = String(data.note || "").trim();

    if (!name) return respond({ ok: false, error: "請填姓名" });
    if (!start || !end) return respond({ ok: false, error: "請選開始/結束日期" });
    if (new Date(end) < new Date(start)) {
      return respond({ ok: false, error: "結束日期不能早於開始日期" });
    }
    var validTypes = ["全天", "上午半天", "下午半天"];
    if (validTypes.indexOf(type) === -1) type = "全天";

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var roster = ss.getSheetByName("員工名單");
    var log = ss.getSheetByName("請假紀錄");
    if (!roster || !log) {
      return respond({ ok: false, error: "找不到「員工名單」或「請假紀錄」分頁，請確認分頁名稱" });
    }

    // 姓名不在名單就自動新增一列
    var lastRosterRow = roster.getLastRow();
    var names = lastRosterRow >= 2
      ? roster.getRange(2, 1, lastRosterRow - 1, 1).getValues().flat()
          .map(function (v) { return String(v).trim(); })
          .filter(Boolean)
      : [];
    if (names.indexOf(name) === -1) {
      roster.getRange(lastRosterRow + 1, 1).setValue(name);
    }

    // 寫入一筆請假紀錄，天數公式跟 xlsx 原本的邏輯一致
    var nextRow = Math.max(log.getLastRow() + 1, 2);
    log.getRange(nextRow, 1, 1, 5).setValues([[name, start, end, type, note]]);
    log.getRange(nextRow, 6).setFormula(
      "=IF(A" + nextRow + '="","",IF(D' + nextRow + '="全天",C' + nextRow + "-B" + nextRow + "+1,0.5))"
    );
    log.getRange(nextRow, 6).setNumberFormat("0.#");
    log.getRange(nextRow, 2, 1, 2).setNumberFormat("yyyy-mm-dd");

    return respond({ ok: true, name: name });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return respond({ ok: true, message: "leave-calendar API is running" });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
