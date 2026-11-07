/**
 * 貼到「日報網頁填寫資料庫」試算表：
 * 選單「擴充功能」→「Apps Script」→ 把整份檔案內容貼進去覆蓋預設的 Code.gs → 存檔
 * → 先跑一次 setupSheets() 建立分頁結構 → 再依 README「部署」步驟部署成 Web App。
 */

var SHEET_LIST = "工種機具材料清單";
var SHEET_RECORD = "日報記錄";
var SHEET_HEADER = "日報頭";
var SHEET_BASIC = "基本資料"; // 舊版單一案場的基本資料，已被「案場設定」取代；setupSheets 不再建立，可直接刪除該分頁（僅首次遷移時才讀）
var SHEET_ATTENDANCE = "本工出勤";
var SHEET_INTERNAL = "內部記錄"; // 選工/備註兩種日誌型記錄，僅後台查看，不進列印/xlsx
var SHEET_BILLING = "請款資料"; // 使用者直接在試算表輸入；顯示在 xlsx 日報的請款表（A46:S52），欄位：案場｜請款日期｜區域｜期別｜本期請款金額｜目前進度｜備註｜發票
var SHEET_CASES = "案場設定"; // 多案場：一案場一列，取代舊版單列的「基本資料」
var SHEET_ITEMS = "合約工項"; // 工項完成量頁的項目主表（照日報原表 DPR 的項目與預估數量），欄位：案場｜序號｜項目名稱｜單位｜預估數量｜期初已完成｜啟用中；使用者可直接修改（綠色）
var SHEET_DONE = "工項完成量"; // 每日完成量（只由「工項完成量」網頁寫入），欄位：日期｜序號｜本日完成量｜clientId｜更新時間｜案場
var CASE_COL_DONE = 6;

// 各資料表「案場」欄位的欄號（一律補在既有欄位最後面，比照當初新增主任/工安欄位的作法：
// 不動既有欄位順序、舊資料不用搬移位置，只需要在最後補值）
var CASE_COL_HEADER = 14;
var CASE_COL_RECORD = 7;
var CASE_COL_ATTENDANCE = 10;
var CASE_COL_INTERNAL = 9;
var CASE_COL_LIST = 5;

var DEFAULT_CASE_ID = "lng"; // 既有資料（洲際LNG）遷移時要補上的案場代碼

// ---------- 一次性設定：建立所有分頁結構（重複執行不會清空既有資料，只補齊缺的分頁/表頭） ----------
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「案場設定」是唯一的基本資料來源（一案場一列）。首次建立時，若試算表還留著舊的「基本資料」分頁，
  // 才從它的值遷移出一列（案場代碼＝DEFAULT_CASE_ID）；沒有該分頁就用預設值，且不會再自動重建「基本資料」。
  // 之後新增案場只要在這張表加一列即可，不用改程式碼。
  var basic = ss.getSheetByName(SHEET_BASIC);
  var cases = ss.getSheetByName(SHEET_CASES) || ss.insertSheet(SHEET_CASES);
  if (cases.getRange(1, 1).getValue() === "") {
    var basicVals = basic ? basic.getRange(1, 1, 5, 2).getValues() : [["工程名稱", "洲際液化天然氣接收站"]];
    var basicObj = {};
    basicVals.forEach(function (r) { basicObj[r[0]] = r[1]; });
    cases.getRange(1, 1, 1, 7).setValues([["案場代碼", "案場名稱", "啟用中", "業主", "合約金額（元）", "開工日期（YYYY/MM/DD）", "公司名稱"]]);
    cases.getRange(2, 1, 1, 7).setValues([[
      DEFAULT_CASE_ID, basicObj["工程名稱"] || "", "TRUE",
      basicObj["業主"] || "", basicObj["合約金額（元）"] || "", basicObj["開工日期（YYYY/MM/DD）"] || "", basicObj["公司名稱"] || "",
    ]]);
    var activeRuleC = SpreadsheetApp.newDataValidation().requireValueInList(["TRUE", "FALSE"], true).build();
    cases.getRange(2, 3, 500, 1).setDataValidation(activeRuleC);
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
  ensureCaseColumn(list, CASE_COL_LIST);

  var record = ss.getSheetByName(SHEET_RECORD) || ss.insertSheet(SHEET_RECORD);
  if (record.getRange(1, 1).getValue() === "") {
    record.getRange(1, 1, 1, 6).setValues([["日期", "項目名稱", "上午", "下午", "clientId", "更新時間"]]);
  }
  setDateColumnText(record); // 日期欄強制純文字，避免Sheets自動轉成日期型別
  ensureCaseColumn(record, CASE_COL_RECORD);

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
  ensureCaseColumn(header, CASE_COL_HEADER);

  var attendance = ss.getSheetByName(SHEET_ATTENDANCE) || ss.insertSheet(SHEET_ATTENDANCE);
  if (attendance.getRange(1, 1).getValue() === "") {
    attendance.getRange(1, 1, 1, 9).setValues([[
      "日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間",
    ]]);
  }
  setDateColumnText(attendance);
  ensureCaseColumn(attendance, CASE_COL_ATTENDANCE);

  var internal = ss.getSheetByName(SHEET_INTERNAL) || ss.insertSheet(SHEET_INTERNAL);
  if (internal.getRange(1, 1).getValue() === "") {
    internal.getRange(1, 1, 1, 8).setValues([[
      "日期", "類型", "類別", "內容", "時間", "備註", "clientId", "更新時間",
    ]]);
    var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(["選工", "備註"], true).build();
    internal.getRange(2, 2, 500, 1).setDataValidation(typeRule);
    // 類別（C 欄）改自由文字，不設下拉限制：網頁端可自訂新分類，後台也能直接手打
    var pickTimeRule = SpreadsheetApp.newDataValidation().requireValueInList(["上午", "下午", "全天"], true).build();
    internal.getRange(2, 5, 500, 1).setDataValidation(pickTimeRule);
  }
  setDateColumnText(internal);
  ensureCaseColumn(internal, CASE_COL_INTERNAL);

  // 「請款資料」：一列一期請款，案場放第 1 欄（下拉選「案場設定」的代碼）。「累計至上期請款金額」不用填，
  // 產生 xlsx 時依同案場同區域的前幾期自動算；請款日期決定從哪一天的日報開始顯示（舊日期的日報不受影響）。
  var billing = ss.getSheetByName(SHEET_BILLING) || ss.insertSheet(SHEET_BILLING);
  if (billing.getRange(1, 1).getValue() === "") {
    billing.getRange(1, 1, 1, 8).setValues([["案場", "請款日期", "區域", "期別", "本期請款金額", "目前進度", "備註", "發票"]]);
    var billingCaseRule = SpreadsheetApp.newDataValidation().requireValueInRange(cases.getRange("A2:A500"), true).build();
    billing.getRange(2, 1, 500, 1).setDataValidation(billingCaseRule);
    billing.getRange(2, 2, 500, 1).setNumberFormat("@"); // 請款日期純文字（2026-10-05、2026/10/5、115.10.5 都可）
    billing.getRange(2, 5, 500, 1).setNumberFormat("#,##0");
    billing.getRange(2, 6, 500, 1).setNumberFormat("0.00%"); // 輸入 35 或 35% 都會是 35%
  }

  // 「合約工項」：工項完成量頁的項目主表。只在分頁是新建／空白時放入 DEFAULT_CASE_ID 的種子資料（日報原表 DPR 的 32 個項目），
  // 之後項目名稱、預估數量、期初已完成、啟用中都由使用者直接在這裡改；重複執行 setupSheets 不會覆蓋已有內容。
  var itemsMaster = ss.getSheetByName(SHEET_ITEMS) || ss.insertSheet(SHEET_ITEMS);
  if (itemsMaster.getRange(1, 1).getValue() === "") {
    itemsMaster.getRange(1, 1, 1, 7).setValues([["案場", "序號", "項目名稱", "單位", "預估數量", "期初已完成", "啟用中"]]);
    var seedRows = ITEM_MASTER_SEED.map(function (r) { return [DEFAULT_CASE_ID].concat(r); });
    itemsMaster.getRange(2, 1, seedRows.length, 7).setValues(seedRows);
    itemsMaster.getRange(2, 1, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(cases.getRange("A2:A500"), true).build());
    itemsMaster.getRange(2, 7, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["TRUE", "FALSE"], true).build());
  }
  // 「工項完成量」：每日完成量，由網頁寫入；日期欄純文字。只補表頭，不動任何既有資料。
  var itemsDone = ss.getSheetByName(SHEET_DONE) || ss.insertSheet(SHEET_DONE);
  if (itemsDone.getRange(1, 1).getValue() === "") {
    itemsDone.getRange(1, 1, 1, 6).setValues([["日期", "序號", "本日完成量", "clientId", "更新時間", "案場"]]);
  }
  setDateColumnText(itemsDone);

  // 給不同使用者看的操作說明：寫在資料欄右側的 J 欄（每次 setupSheets 都會重寫成最新版，請勿在此輸入資料）
  writeSheetGuide_(cases, GUIDE_COL, CASES_GUIDE);
  writeSheetGuide_(billing, GUIDE_COL, BILLING_GUIDE);

  // 顏色標示哪些分頁/格子可以人工輸入（見 COLOR_LEGEND）
  applyColorGuide_(ss);
}

var GUIDE_COL = 10; // J 欄：兩張設定分頁的資料欄都在 A~H 以內
var UPDATE_STEPS = [
  "【如何把人工輸入/修改套用到 Drive 的 xlsx 日報】",
  "1. 在「綠色」可輸入的格子填好或改好資料後，點上方選單「日報管理」→「更新全部日報（套用案場設定）」。",
  "2. 跳出確認視窗，按「是」。第一次使用 Google 會要求授權：選你的帳號→「進階」→「前往（不安全）」→「允許」，再重按一次選單。",
  "3. 等候約 1~2 分鐘（日報天數越多越久），出現「更新完成」視窗即可。每個案場各顯示一行結果，有「失敗」請截圖給系統管理員。",
  "4. 更新只會重寫 xlsx 的抬頭與請款表，不會改動日報內容與累計。Drive 上的 xlsx 只當輸出檔，請勿手動改格子（下次更新會被覆蓋）。",
];
var COLOR_LEGEND = [
  "【顏色說明：哪些可以人工輸入】",
  "🟩 綠色（分頁標籤＋表頭＋淡綠格子）＝可人工輸入/修改：「案場設定」「請款資料」「合約工項」，以及「工種機具材料清單」的「工項編號」「啟用中」兩欄。改完要照下面步驟按「更新」，Drive 的 xlsx 才會同步。",
  "🟦 藍色＝「內部記錄」：可人工修改，不進列印也不進 xlsx，不需要按更新。",
  "⬜ 灰色（分頁標籤＋表頭）＝由日報網頁寫入：「日報頭」「日報記錄」「本工出勤」「工項完成量」，以及清單的「類別」「項目名稱」「案場」欄。請勿手動輸入；要修正請回日報網頁選該日期修改，或用選單「日報管理→刪除某天日報資料」。",
  "（每個分頁左上角 A1 的備註也有同樣說明，滑鼠移上去就會顯示。）",
];
var CASES_GUIDE = [
  "📌 使用說明（程式維護，請勿在此欄輸入資料）",
  "這張表是日報「基本資料」的唯一來源：一個案場一列（案場代碼／案場名稱／啟用中／業主／合約金額／開工日期／公司名稱）。",
  "・網頁「列印日報」：修改後重新整理網頁就生效，舊日期與新日期都一樣。",
  "・Drive 的 xlsx 日報：修改後不會自動更新，需照下面步驟按一次更新。",
  "・「案場代碼」建立後請勿更改（舊資料靠它對應案場）；不用的案場把「啟用中」改成 FALSE，不要刪列。",
  "・新增案場：在最下方空白列填一列，或在日報網頁按「＋新增案場」。",
  "",
].concat(COLOR_LEGEND, [""], UPDATE_STEPS);
var BILLING_GUIDE = [
  "📌 使用說明（程式維護，請勿在此欄輸入資料）",
  "這張表是 xlsx 日報「請款表」的資料來源：一列一期請款。",
  "・案場：用下拉選單選（來自「案場設定」的代碼）。",
  "・請款日期：從這一天的日報開始顯示，之前日期的日報不受影響。可輸入 2026-10-05、2026/10/5 或民國 115.10.5。",
  "・本期請款金額：填數字（可含千分位逗號）。目前進度：填 35 或 35% 皆可。區域、期別、備註、發票：自由輸入（發票可填日期）。",
  "・「累計至上期請款金額」不用填：系統依同案場、同區域、較早的各期自動加總。",
  "・日報請款表最多顯示 5 列（最新的 5 列，超過會在更新結果警告）。超過時請自行把舊的幾期合併成一列累加金額，並刪除被合併的列。",
  "・網頁「列印日報」不印請款表；請款表只出現在 Drive 的 xlsx 日報。請勿更動第 1 列表頭與欄位順序。",
  "",
].concat(COLOR_LEGEND, [""], UPDATE_STEPS);

// 顏色標示：綠＝可人工輸入（改完按更新）、藍＝可人工修改但不進 xlsx、灰＝網頁寫入請勿手改。
// 只動格式（分頁標籤色、表頭底色、可輸入區底色、A1 備註），不動資料；每次 setupSheets 都會重套。
var GUIDE_COLORS = {
  inputHeader: "#a5d6a7", inputBody: "#e8f5e9", sysHeader: "#d9d9d9", memoHeader: "#bbdefb",
  tabInput: "#43a047", tabSys: "#9e9e9e", tabMemo: "#1e88e5",
};
function colorHeader_(sheet, firstCol, lastCol, bg) {
  var r = sheet.getRange(1, firstCol, 1, lastCol - firstCol + 1);
  r.setBackground(bg);
  r.setFontWeight("bold");
}
function applyColorGuide_(ss) {
  var C = GUIDE_COLORS;
  function put(name, fn) { var sh = ss.getSheetByName(name); if (sh) fn(sh); }
  put(SHEET_CASES, function (sh) {
    sh.setTabColor(C.tabInput);
    colorHeader_(sh, 1, 7, C.inputHeader);
    sh.getRange(2, 1, 499, 7).setBackground(C.inputBody);
    sh.getRange("A1").setNote("🟩 綠色＝可人工輸入。改完請按選單「日報管理→更新全部日報」，Drive 的 xlsx 才會同步（說明見 J 欄）。");
  });
  put(SHEET_BILLING, function (sh) {
    sh.setTabColor(C.tabInput);
    colorHeader_(sh, 1, 8, C.inputHeader);
    sh.getRange(2, 1, 499, 8).setBackground(C.inputBody);
    sh.getRange("A1").setNote("🟩 綠色＝可人工輸入。輸入後請按選單「日報管理→更新全部日報」，Drive 的 xlsx 才會顯示（說明見 J 欄）。");
  });
  put(SHEET_ITEMS, function (sh) {
    sh.setTabColor(C.tabInput);
    colorHeader_(sh, 1, 7, C.inputHeader);
    sh.getRange(2, 1, 499, 7).setBackground(C.inputBody);
    sh.getRange("A1").setNote("🟩 綠色＝可人工修改（項目名稱、單位、預估數量、期初已完成、啟用中）。序號是對應每日完成量的鍵，建立後請勿更改；不用的項目把「啟用中」改成 FALSE，不要刪列。網頁重新整理後生效，Drive 的 xlsx 在下次送出工項完成量時更新。");
  });
  put(SHEET_DONE, function (sh) {
    sh.setTabColor(C.tabSys);
    colorHeader_(sh, 1, 6, C.sysHeader);
    sh.getRange("A1").setNote("⬜ 由日報網頁「工項完成量」頁寫入，請勿手動輸入。要修正請回該頁選日期修改。");
  });
  put(SHEET_LIST, function (sh) {
    sh.setTabColor(C.tabInput);
    colorHeader_(sh, 1, 2, C.sysHeader);
    colorHeader_(sh, 3, 4, C.inputHeader);
    colorHeader_(sh, 5, 5, C.sysHeader);
    sh.getRange(2, 3, 499, 2).setBackground(C.inputBody);
    sh.getRange("A1").setNote("🟩 只有「工項編號」「啟用中」兩欄可人工輸入（綠色），改完請按選單「日報管理→更新全部日報」。⬜ 灰色欄（類別/項目名稱/案場）請勿手改，改項目名稱會讓歷史累計對不上。");
  });
  put(SHEET_INTERNAL, function (sh) {
    sh.setTabColor(C.tabMemo);
    colorHeader_(sh, 1, 9, C.memoHeader);
    sh.getRange("A1").setNote("🟦 可人工修改，不進列印也不進 xlsx，不需要按更新。");
  });
  [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE].forEach(function (name) {
    put(name, function (sh) {
      sh.setTabColor(C.tabSys);
      colorHeader_(sh, 1, Math.max(sh.getLastColumn(), 1), C.sysHeader);
      sh.getRange("A1").setNote("⬜ 由日報網頁寫入，請勿手動輸入。要修正請回日報網頁選該日期修改，或用選單「日報管理→刪除某天日報資料」。");
    });
  });
}

// 在指定欄由上而下寫入說明文字（自動換行、淡黃底、第一列粗體），欄寬放寬以便閱讀
function writeSheetGuide_(sheet, col, lines) {
  var range = sheet.getRange(1, col, lines.length, 1);
  range.setValues(lines.map(function (l) { return [l]; }));
  range.setWrap(true).setVerticalAlignment("top").setBackground("#fff8e1");
  sheet.getRange(1, col).setFontWeight("bold");
  sheet.setColumnWidth(col, 520);
}

// 補上「案場」欄位（若表頭已經是「案場」代表已migrate過，跳過）；既有資料列補上 DEFAULT_CASE_ID，
// 避免遷移前的舊資料（目前只有洲際LNG）變成讀不到案場的孤兒列。
function ensureCaseColumn(sheet, caseColIndex) {
  if (sheet.getRange(1, caseColIndex).getValue() === "案場") return;
  sheet.getRange(1, caseColIndex).setValue("案場");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, caseColIndex, lastRow - 1, 1);
  var vals = range.getValues();
  var changed = false;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === "" || vals[i][0] === null) { vals[i][0] = DEFAULT_CASE_ID; changed = true; }
  }
  if (changed) range.setValues(vals);
}

// 日期欄（A 欄）第 2 列起到分頁目前的最後一列全設純文字（不再只涵蓋前 999 列）。
// 資料列成長超過目前列數時，新增的列由 Sheets 沿用相鄰列格式；normalizeDate 仍保留作為最後防線。
function setDateColumnText(sheet) {
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rows, 1).setNumberFormat("@");
}

// 在資料欄（前 values.length 欄）的第一個空白列寫入一列。「案場設定」「請款資料」的 J 欄有說明文字，
// appendRow 會接在說明文字最後一列之後（中間空一大段），所以改用這個函式往上找空白列。回傳寫入的列號。
function appendRowTop_(sheet, values) {
  var width = values.length;
  var lastRow = sheet.getLastRow();
  var data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
  var row = Math.max(lastRow, 1) + 1;
  for (var i = 0; i < data.length; i++) {
    var blank = true;
    for (var j = 0; j < width; j++) { if (data[i][j] !== "" && data[i][j] !== null) { blank = false; break; } }
    if (blank) { row = i + 2; break; }
  }
  ensureRows(sheet, row);
  sheet.getRange(row, 1, 1, width).setValues([values]);
  return row;
}

// 寫入前確保分頁列數足夠：真實 Sheets 對超出列數的 setValues 會丟例外（只有 appendRow 會自動擴充），
// 日報記錄約 70~100 天就會用完預設的 1000 列。新增的列沿用上方列的格式（含日期欄純文字）。
function ensureRows(sheet, row) {
  var max = sheet.getMaxRows();
  if (row > max) sheet.insertRowsAfter(max, row - max + 200);
}

// 工種/機具/材料清單的起始種子資料，抄自現行 Excel 日報範本的預設項目。
// 啟用中預設 TRUE；工項編號留空，之後要接成本管理系統時再填。案場欄留空，setupSheets 會補 DEFAULT_CASE_ID。
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

// 「合約工項」種子資料：日報原表（Talin Project DPR…清理版2.xlsx 的 1106 分頁）第 11～42 列，欄位 [序號, 項目名稱, 單位, 預估數量, 期初已完成, 啟用中]。
// 預估數量＝目前合約總量，工項完成量累計超過時網頁只提醒、不擋送出。
var ITEM_MASTER_SEED = [
  [1, "樁頭及錨定處理", "PC", 325, 0, "TRUE"],
  [2, "深井排水", "PT", 26, 0, "TRUE"],
  [3, "擋土壁,鋼板樁支撐", "m2", 3057, 0, "TRUE"],
  [4, "擋土壁, 鋼鈑樁,樁長 = 10m", "m", 395, 0, "TRUE"],
  [5, "開挖", "m3", 14183, 0, "TRUE"],
  [6, "回填及夯實, 使用原土", "m3", 3685, 0, "TRUE"],
  [7, "棄土, 殘土", "m3", 1263, 0, "TRUE"],
  [8, "無筋混凝土, II型140kgf/cm2", "m3", 151, 0, "TRUE"],
  [9, "結構混凝土, II型280kgf/cm2", "m3", 7844, 0, "TRUE"],
  [10, "普通模板, 支撐高H < 7M", "m2", 20554, 0, "TRUE"],
  [11, "清水模板, 支撐高H < 7M", "m2", 5854, 0, "TRUE"],
  [12, "清水模板, 支撐高H ≧ 7M", "m2", 360, 0, "TRUE"],
  [13, "竹節鋼筋y=4200kgf/cm2", "ton", 1614, 0, "TRUE"],
  [14, "鋼筋續接器, SA級, D25", "set", 3108, 0, "TRUE"],
  [15, "鋼筋續接器, SA級, D32", "set", 90, 0, "TRUE"],
  [16, "預埋式錨栓, 碳鋼, 熱浸鍍鋅", "kg", 64, 0, "TRUE"],
  [17, "預埋鐵件, 鍍鋅", "kg", 3040, 0, "TRUE"],
  [18, "預埋、套管, 鍍鋅鋼管含法蘭", "kg", 152, 0, "TRUE"],
  [19, "預埋、套管, 不銹鋼管含法蘭", "kg", 152, 0, "TRUE"],
  [20, "預埋、套管, PVC管含法蘭", "m", 5, 0, "TRUE"],
  [21, "預埋、套管, 鍍鋅鋼管不含法蘭", "kg", 152, 0, "TRUE"],
  [22, "預埋、套管, 不銹鋼管不含法蘭", "kg", 152, 0, "TRUE"],
  [23, "預埋、套管, PVC管不含法蘭", "m", 4, 0, "TRUE"],
  [24, "預埋鐵件, 不銹鋼 SUS 304", "kg", 720, 0, "TRUE"],
  [25, "鋼構.重型型(60~150kg/m),SN490B鍍鋅油漆", "ton", 5, 0, "TRUE"],
  [26, "扶手, 鋼管, 不銹鋼(供料及預製), 油漆", "ton", 3, 0, "TRUE"],
  [27, "樓梯踏步, 供料及預製, 鍍鋅鐵格柵", "ton", 1, 0, "TRUE"],
  [28, "鋼構安裝, 重型型鋼(60~150 kg/m)", "ton", 5, 0, "TRUE"],
  [29, "樓梯踏步,安裝,鍍鋅鐵格柵", "ton", 1, 0, "TRUE"],
  [30, "50mm厚底版鋪保麗龍保護板", "m2", 3258, 0, "TRUE"],
  [31, "2.5mm厚底版舖自黏式改質瀝青防水毯", "m2", 2061, 0, "TRUE"],
  [32, "2.5mm厚牆面舖自黏式改質瀝青防水毯", "m2", 1197, 0, "TRUE"],
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

// ---------- 案場：查有效案場（依代碼），供 doGet/doPost 驗證與取基本資料用 ----------
// @return {code, name, owner, contract, startDate, company} 或 null（不存在/已停用）
function getCaseRow_(ss, caseId) {
  var sheet = ss.getSheetByName(SHEET_CASES);
  if (!sheet || !caseId) return null;
  var rows = sheet.getDataRange().getValues();
  rows.shift();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === caseId && String(rows[i][2]).trim().toUpperCase() === "TRUE") {
      return { code: String(rows[i][0]).trim(), name: rows[i][1], owner: rows[i][3], contract: rows[i][4], startDate: rows[i][5], company: rows[i][6] };
    }
  }
  return null;
}

// ---------- 找列：讀第2列起、寬度 width 欄，回傳符合 predicate 的第一列列號（找不到 -1） ----------
function findRow_(sheet, width, predicate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var vals = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (predicate(vals[i])) return i + 2;
  }
  return -1;
}

// 找「日期欄=keyVal 且 案場欄=caseId」的列（日報頭用）
function findRowByKey(sheet, keyCol, keyVal, caseCol, caseId) {
  return findRow_(sheet, caseCol, function (r) {
    return normalizeDate(r[keyCol - 1]) === normalizeDate(keyVal) && String(r[caseCol - 1] || "").trim() === caseId;
  });
}

// 找「日期+第二欄+案場」都相符的列（日報記錄/本工出勤共用，三欄複合鍵）
function findRecordRow(sheet, date, key2, caseCol, caseId) {
  return findRow_(sheet, caseCol, function (r) {
    return normalizeDate(r[0]) === date && unsanitizeCell(r[1]).trim() === key2 && String(r[caseCol - 1] || "").trim() === caseId;
  });
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
var MAX_BILLING_ITEMS = 20; // 網頁一次最多新增的請款項目數

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

// 金額欄位（請款）：必填；可含千分位逗號；不可為負、不可超過 1 兆；四捨五入到小數 2 位
function checkMoney(v, label) {
  if (v === null || v === undefined || (typeof v === "string" && v.trim() === "")) throw validationError("請輸入" + label);
  if (typeof v !== "number" && typeof v !== "string") throw validationError(label + "格式錯誤（必須是數字）");
  var n = typeof v === "number" ? v : Number(String(v).replace(/[,\s]/g, ""));
  if (!isFinite(n)) throw validationError(label + "不是有效的數字");
  if (n < 0) throw validationError(label + "不可為負數");
  if (n > 1e12) throw validationError(label + "超出合理範圍");
  return Math.round(n * 100) / 100;
}
// 進度（百分比 0~100，可含 %）：選填；回傳比例（35 → 0.35，試算表該欄是 0.00% 格式），空白回傳 ""
function checkProgressPct(v, label) {
  if (v === null || v === undefined || (typeof v === "string" && v.replace(/[%\s]/g, "") === "")) return "";
  if (typeof v !== "number" && typeof v !== "string") throw validationError(label + "格式錯誤（必須是數字）");
  var n = typeof v === "number" ? v : Number(String(v).replace(/[%\s,]/g, ""));
  if (!isFinite(n)) throw validationError(label + "不是有效的數字");
  if (n < 0 || n > 100) throw validationError(label + "必須在 0 到 100 之間");
  return Math.round(n * 10000) / 1000000;
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
    if (action !== "submit" && action !== "addItem" && action !== "addCase" && action !== "addBilling" && action !== "saveWorkItems") return respond({ ok: false, error: "不支援的操作" });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (firstMissingSheet(ss, [SHEET_CASES])) return respond({ ok: false, error: NOT_SETUP_ERROR });

    // ---- 新增案場（給案場下拉旁「＋新增案場」按鈕用）：不需要既有合法 caseId，這裡就是在建立新的 ----
    if (action === "addCase") {
      var newCode = checkText(data.code, "案場代碼", MAX_SHORT_TEXT, true);
      if (!newCode) return respond({ ok: false, error: "請輸入案場代碼" });
      if (!/^[A-Za-z0-9_-]+$/.test(newCode)) return respond({ ok: false, error: "案場代碼只能用英數字、底線、連字號" });
      var newName = checkText(data.name, "案場名稱", MAX_SHORT_TEXT, true);
      if (!newName) return respond({ ok: false, error: "請輸入案場名稱" });
      if (/[\\\/:*?"<>|]/.test(newName)) return respond({ ok: false, error: "案場名稱不能包含 \\ / : * ? \" < > | 這些字元（會用在檔名）" });

      var newOwner = checkText(data.owner, "業主", MAX_SHORT_TEXT);
      var newContract = checkText(data.contract, "合約金額", MAX_SHORT_TEXT);
      var newStartDate = checkText(data.startDate, "開工日期", MAX_SHORT_TEXT);
      var newCompany = checkText(data.company, "公司名稱", MAX_SHORT_TEXT);

      var caseSheet = ss.getSheetByName(SHEET_CASES);
      var caseRows = caseSheet.getDataRange().getValues();
      for (var ci = 1; ci < caseRows.length; ci++) {
        if (String(caseRows[ci][0]).trim() === newCode) return respond({ ok: false, error: "案場代碼已存在：" + newCode });
      }
      appendRowTop_(caseSheet, [newCode, newName, "TRUE", newOwner, newContract, newStartDate, newCompany]);

      // 選填：從既有案場複製目前的工種/機具/材料清單當新案場的起始清單（各案場清單互相獨立，之後各自維護）
      var copyFrom = String(data.copyFrom || "").trim();
      var copied = 0;
      if (copyFrom && !firstMissingSheet(ss, [SHEET_LIST])) {
        var listSheetC = ss.getSheetByName(SHEET_LIST);
        var listRowsC = listSheetC.getDataRange().getValues();
        for (var li = 1; li < listRowsC.length; li++) {
          if (String(listRowsC[li][CASE_COL_LIST - 1] || "").trim() === copyFrom) {
            listSheetC.appendRow([listRowsC[li][0], sanitizeCell(unsanitizeCell(listRowsC[li][1]).trim()), listRowsC[li][2], listRowsC[li][3], newCode]);
            copied++;
          }
        }
      }
      return respond({ ok: true, code: newCode, copied: copied });
    }

    var caseId = String(data.caseId || "").trim();
    var caseInfo = getCaseRow_(ss, caseId);
    if (!caseInfo) return respond({ ok: false, error: "案場不存在或已停用" });

    // ---- 新增工種/機具/材料項目（給表單上的「＋新增項目」按鈕用）：只在該案場範圍內比對是否已存在 ----
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
        if (unsanitizeCell(listRows[i][1]).trim() === name && String(listRows[i][CASE_COL_LIST - 1] || "").trim() === caseId) {
          if (String(listRows[i][3]).trim().toUpperCase() === "TRUE") {
            return respond({ ok: false, error: "項目已存在" });
          }
          listSheet.getRange(i + 1, 1).setValue(category); // 曾被停用的項目：以新傳入的類別更新並重新啟用
          listSheet.getRange(i + 1, 4).setValue("TRUE");
          return respond({ ok: true });
        }
      }
      listSheet.appendRow([category, sanitizeCell(name), "", "TRUE", caseId]);
      return respond({ ok: true });
    }

    // ---- 新增請款資料（網頁「請款資料」卡片）：一次可送多個項目、全有或全無；只新增、不覆蓋；修改/刪除請直接到試算表「請款資料」分頁 ----
    // data.date 為共用的請款日期；data.items 為 [{zone, period, amount, progress, remark, invoice}]（沒有 items 時視為單筆，欄位直接放在 data 上）
    if (action === "addBilling") {
      if (firstMissingSheet(ss, [SHEET_BILLING])) return respond({ ok: false, error: "尚未建立「請款資料」分頁，請先執行 setupSheets" });
      var bDate = checkText(data.date, "請款日期", 20, true);
      if (!/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(bDate) || parseAdminDate(bDate) !== bDate) return respond({ ok: false, error: "請款日期格式錯誤" });
      var bIn = Array.isArray(data.items) ? data.items : [data];
      if (!bIn.length) return respond({ ok: false, error: "沒有要新增的請款項目" });
      if (bIn.length > MAX_BILLING_ITEMS) return respond({ ok: false, error: "一次最多新增 " + MAX_BILLING_ITEMS + " 筆請款項目" });

      // 先全部驗證（任何一項不合法就整批拒絕，不寫入）
      var bNew = [], bSeen = {};
      for (var bn = 0; bn < bIn.length; bn++) {
        var it = bIn[bn] || {};
        var tag = bIn.length > 1 ? "請款項目 " + (bn + 1) + " 的" : "";
        var bZone = checkText(it.zone, tag + "區域", MAX_SHORT_TEXT, true);
        var bPeriod = checkText(it.period, tag + "期別", MAX_SHORT_TEXT, true);
        if (!bZone) return respond({ ok: false, error: "請輸入" + tag + "區域" });
        if (!bPeriod) return respond({ ok: false, error: "請輸入" + tag + "期別" });
        var key = JSON.stringify([bZone, bPeriod]);
        if (bSeen[key]) return respond({ ok: false, error: "請款項目 " + (bn + 1) + " 與前面的項目重複（區域「" + bZone + "」、期別「" + bPeriod + "」）" });
        bSeen[key] = true;
        bNew.push({
          zone: bZone, period: bPeriod,
          amount: checkMoney(it.amount, tag + "本期請款金額"),
          progress: checkProgressPct(it.progress, tag + "目前進度"),
          remark: checkText(it.remark, tag + "備註", MAX_SHORT_TEXT, true),
          invoice: checkText(it.invoice, tag + "發票", MAX_SHORT_TEXT, true),
        });
      }

      var billingSheet = ss.getSheetByName(SHEET_BILLING);
      var billingRows = billingSheet.getDataRange().getValues();
      for (var bi = 1; bi < billingRows.length; bi++) {
        var br = billingRows[bi];
        if (String(br[0] || "").trim() !== caseId || xlsxBillingDate_(br[1]) !== bDate) continue;
        var exZone = unsanitizeCell(br[2]).trim(), exPeriod = unsanitizeCell(br[3]).trim();
        for (var bk = 0; bk < bNew.length; bk++) {
          if (bNew[bk].zone === exZone && bNew[bk].period === exPeriod) {
            return respond({ ok: false, error: "這筆請款資料已存在（" + bDate + "、區域「" + exZone + "」、期別「" + exPeriod + "」），如需修改請到試算表「請款資料」分頁；本批資料都未新增" });
          }
        }
      }
      for (var bw = 0; bw < bNew.length; bw++) {
        var nr = bNew[bw];
        appendRowTop_(billingSheet, [caseId, bDate, sanitizeCell(nr.zone), sanitizeCell(nr.period), nr.amount, nr.progress, sanitizeCell(nr.remark), sanitizeCell(nr.invoice)]);
      }

      // 更新雲端 xlsx（只更新一次）：請款日期當天及之後的分頁；失敗不影響已寫入的請款資料
      var bxlsx;
      try {
        var bxr = xlsxUpdateForDate(bDate, caseId, caseInfo);
        bxlsx = { ok: true, url: bxr.url, warnings: bxr.warnings || [] };
      } catch (bxerr) {
        // 該案場還沒有任何日報分頁可寫（尚未送出過日報）：沒有東西可更新，不算失敗
        bxlsx = /沒有任何日期可寫入/.test(String(bxerr)) ? { ok: true, skipped: true, warnings: [] } : { ok: false, error: String(bxerr) };
      }
      return respond({ ok: true, added: bNew.length, xlsx: bxlsx });
    }

    // ---- 送出／更新當天的工項完成量（「工項完成量」網頁）----
    // 只動這一天、被明確列出的序號：items＝要寫入的 [{seq, qty}]（qty 為 0 視為沒有）、clearItems＝使用者明確清空的序號；
    // 沒被列出的既有資料一律不動（舊頁面或兩人分填時不會互相蓋掉）。完全不碰日報頭／日報記錄等日報資料。
    // 寫完後只更新 xlsx 的「工項.xxx」分頁（xlsxUpdateItems_ 不重寫、不移除任何日報分頁）；xlsx 失敗不影響已寫入的資料。
    if (action === "saveWorkItems") {
      if (firstMissingSheet(ss, [SHEET_ITEMS, SHEET_DONE])) return respond({ ok: false, error: "尚未建立「合約工項」「工項完成量」分頁，請先執行 setupSheets" });
      var wsDate = checkText(data.date, "日期", 20, true);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(wsDate) || parseAdminDate(wsDate) !== wsDate) return respond({ ok: false, error: "日期格式錯誤" });
      var wsClientId = sanitizeCell(checkText(data.clientId, "clientId", MAX_SHORT_TEXT, true));
      var wsIn = Array.isArray(data.items) ? data.items : [];
      var wsClear = Array.isArray(data.clearItems) ? data.clearItems : [];
      if (wsIn.length > 200 || wsClear.length > 200) return respond({ ok: false, error: "一次送出的項目過多" });
      var wsMaster = {};
      xlsxLoadItems_(ss, caseId).master.forEach(function (m) { wsMaster[m.seq] = m; });

      // 先全部驗證，任何不合法就整個拒絕、不寫入
      var wsWrites = [], wsDeletes = [], wsSeen = {};
      for (var wi2 = 0; wi2 < wsIn.length; wi2++) {
        var wit = wsIn[wi2];
        if (!wit || typeof wit !== "object" || Array.isArray(wit)) throw validationError("第 " + (wi2 + 1) + " 筆工項資料格式錯誤");
        var wSeq = Number(wit.seq);
        if (!isFinite(wSeq) || Math.floor(wSeq) !== wSeq || !wsMaster[wSeq]) throw validationError("第 " + (wi2 + 1) + " 筆的工項序號不在「合約工項」內");
        if (wsSeen[wSeq]) throw validationError("工項序號 " + wSeq + " 重複");
        wsSeen[wSeq] = true;
        var wQty = checkQty(wit.qty, "工項 " + wSeq + " 的完成量");
        if (wQty > 1e9) throw validationError("工項 " + wSeq + " 的完成量超出合理範圍");
        wQty = Math.round(wQty * 100) / 100;
        if (wQty > 0) wsWrites.push({ seq: wSeq, qty: wQty }); else wsDeletes.push(wSeq);
      }
      for (var wc = 0; wc < wsClear.length; wc++) {
        var wcSeq = Number(wsClear[wc]);
        if (!isFinite(wcSeq) || Math.floor(wcSeq) !== wcSeq || !wsMaster[wcSeq]) throw validationError("要清除的工項序號不在「合約工項」內");
        if (wsSeen[wcSeq] && wsWrites.some(function (x) { return x.seq === wcSeq; })) throw validationError("工項序號 " + wcSeq + " 同時要寫入與清除");
        if (wsDeletes.indexOf(wcSeq) === -1) wsDeletes.push(wcSeq);
      }

      var wsSheet = ss.getSheetByName(SHEET_DONE), wsNow = new Date();
      function wsFindRow(seq) {
        return findRow_(wsSheet, CASE_COL_DONE, function (r) {
          return normalizeDate(r[0]) === wsDate && Number(r[1]) === seq && String(r[CASE_COL_DONE - 1] || "").trim() === caseId;
        });
      }
      wsWrites.forEach(function (w) {
        var vals = [wsDate, w.seq, w.qty, wsClientId, wsNow, caseId];
        var row = wsFindRow(w.seq);
        if (row === -1) row = wsSheet.getLastRow() + 1;
        ensureRows(wsSheet, row);
        wsSheet.getRange(row, 1, 1, vals.length).setValues([vals]);
      });
      var wsDeleted = 0;
      wsDeletes.forEach(function (seq) {
        var row = wsFindRow(seq);
        if (row !== -1) { wsSheet.deleteRow(row); wsDeleted++; }
      });

      var wsx;
      try {
        var wxr = xlsxUpdateItems_(caseId, caseInfo);
        wsx = { ok: true, url: wxr.url || "", skipped: !!wxr.skipped, warnings: wxr.warnings || [] };
      } catch (wxerr) {
        wsx = { ok: false, error: String(wxerr) };
      }
      return respond({ ok: true, date: wsDate, saved: wsWrites.length, cleared: wsDeleted, xlsx: wsx });
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
    var internalNotes = Array.isArray(data.internalNotes) ? data.internalNotes : [];

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
      caseId,
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

    // ---- 內部記錄：選工（結構化）或備註（純文字）兩種列，僅後台查看，不進列印/xlsx ----
    var cleanInternal = [];
    for (var n = 0; n < internalNotes.length; n++) {
      var note = internalNotes[n];
      if (!note || typeof note !== "object" || Array.isArray(note)) throw validationError("第 " + (n + 1) + " 筆內部記錄格式錯誤");
      var noteType = String(note.type || "").trim();
      if (["選工", "備註"].indexOf(noteType) === -1) throw validationError("第 " + (n + 1) + " 筆內部記錄類型錯誤");
      var noteContent = checkText(note.content, "內部記錄內容", MAX_LONG_TEXT, true);
      if (!noteContent) continue; // 沒填內容的列不寫入
      if (noteType === "選工") {
        var noteCategory = checkText(note.category, "選工記錄類別", MAX_SHORT_TEXT, true);
        if (!noteCategory) throw validationError("第 " + (n + 1) + " 筆選工記錄缺少類別");
        var noteTime = String(note.time || "").trim();
        if (["上午", "下午", "全天"].indexOf(noteTime) === -1) throw validationError("第 " + (n + 1) + " 筆選工記錄時間錯誤");
        cleanInternal.push({
          type: noteType, category: noteCategory, content: noteContent, time: noteTime,
          remark: checkText(note.remark, "選工記錄備註", MAX_SHORT_TEXT),
        });
      } else {
        cleanInternal.push({ type: noteType, category: "", content: noteContent, time: "", remark: "" });
      }
    }

    if (firstMissingSheet(ss, [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE, SHEET_INTERNAL])) return respond({ ok: false, error: NOT_SETUP_ERROR });

    // ---- upsert 日報頭：同一天同一案場已有資料就覆蓋，沒有就新增 ----
    var headerSheet = ss.getSheetByName(SHEET_HEADER);
    var hRow = findRowByKey(headerSheet, 1, date, CASE_COL_HEADER, caseId);
    if (hRow === -1) hRow = headerSheet.getLastRow() + 1;
    ensureRows(headerSheet, hRow);
    headerSheet.getRange(hRow, 1, 1, hVals.length).setValues([hVals]);

    // ---- 日報記錄：先刪除使用者明確清空的項目，再逐一 upsert（同一天同一項目同一案場已有資料就覆蓋，沒有就新增） ----
    var recordSheet = ss.getSheetByName(SHEET_RECORD);
    for (var c2 = 0; c2 < clearNames.length; c2++) {
      var cRow = findRecordRow(recordSheet, date, clearNames[c2], CASE_COL_RECORD, caseId);
      if (cRow !== -1) recordSheet.deleteRow(cRow);
    }
    for (var j2 = 0; j2 < cleanItems.length; j2++) {
      var ci = cleanItems[j2];
      var rRow = findRecordRow(recordSheet, date, ci.name, CASE_COL_RECORD, caseId);
      var rVals = [date, sanitizeCell(ci.name), ci.am, ci.pm, clientId, now, caseId];
      if (rRow === -1) rRow = recordSheet.getLastRow() + 1;
      ensureRows(recordSheet, rRow);
      recordSheet.getRange(rRow, 1, 1, rVals.length).setValues([rVals]);
    }

    // ---- 本工出勤：先刪除使用者明確移除的人，再逐一 upsert（同一天同一人同一案場已有資料就覆蓋，沒有就新增） ----
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    for (var r2 = 0; r2 < removeNames.length; r2++) {
      var rmRow = findRecordRow(attSheet, date, removeNames[r2], CASE_COL_ATTENDANCE, caseId);
      if (rmRow !== -1) attSheet.deleteRow(rmRow);
    }
    for (var k2 = 0; k2 < cleanAtt.length; k2++) {
      var ca = cleanAtt[k2];
      var aRow = findRecordRow(attSheet, date, ca.name, CASE_COL_ATTENDANCE, caseId);
      var aVals = [
        date, sanitizeCell(ca.name),
        ca.am ? "V" : "", ca.pm ? "V" : "",
        ca.amHours, ca.pmHours,
        sanitizeCell(ca.reason),
        clientId, now, caseId,
      ];
      if (aRow === -1) aRow = attSheet.getLastRow() + 1;
      ensureRows(attSheet, aRow);
      attSheet.getRange(aRow, 1, 1, aVals.length).setValues([aVals]);
    }

    // ---- 內部記錄：只新增，不刪除/覆蓋既有資料——這張表刻意獨立於日報的送出/修改之外，
    // 要修改或刪除既有內容，請直接到後台試算表編輯，不受日報重新送出影響。
    var internalSheet = ss.getSheetByName(SHEET_INTERNAL);
    for (var n2 = 0; n2 < cleanInternal.length; n2++) {
      var ni = cleanInternal[n2];
      internalSheet.appendRow([
        date, ni.type, sanitizeCell(ni.category), sanitizeCell(ni.content), ni.time, sanitizeCell(ni.remark), clientId, now, caseId,
      ]);
    }

    // ---- 更新雲端 xlsx（xlsx-drive.gs）：已持有 ScriptLock，內部不再取鎖；失敗不影響日報送出 ----
    var xlsx;
    try {
      var xr = xlsxUpdateForDate(date, caseId, caseInfo);
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
  // 除了「案場設定」本身，其餘依案場過濾的 action 都同時需要「案場設定」分頁存在
  var neededSheets = {
    cases: [SHEET_CASES],
    config: [SHEET_LIST, SHEET_CASES],
    day: [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE, SHEET_CASES],
    cumulative: [SHEET_LIST, SHEET_RECORD, SHEET_CASES],
    reporters: [SHEET_HEADER, SHEET_CASES],
    peopleNames: [SHEET_ATTENDANCE, SHEET_CASES],
    internalCategories: [SHEET_INTERNAL, SHEET_CASES],
    xlsxUrl: [SHEET_CASES],
    workItems: [SHEET_ITEMS, SHEET_DONE, SHEET_CASES],
    leaderNames: [SHEET_HEADER, SHEET_CASES],
  };
  if (Object.prototype.hasOwnProperty.call(neededSheets, action) && firstMissingSheet(ss, neededSheets[action])) {
    return respond({ ok: false, error: NOT_SETUP_ERROR });
  }

  // 除了「cases」（列出可選案場，本身不需要指定案場）與尚未列出的 action，其餘一律要求合法的案場
  var caseId = String(e.parameter.caseId || "").trim();
  var caseInfo = null;
  var CASE_REQUIRED = ["config", "day", "cumulative", "reporters", "peopleNames", "internalCategories", "xlsxUrl", "workItems", "leaderNames"];
  if (CASE_REQUIRED.indexOf(action) !== -1) {
    caseInfo = getCaseRow_(ss, caseId);
    if (!caseInfo) return respond({ ok: false, error: "案場不存在或已停用" });
  }

  if (action === "cases") {
    var caseSheet = ss.getSheetByName(SHEET_CASES);
    var caseRows = caseSheet.getDataRange().getValues();
    caseRows.shift();
    var cases = caseRows
      .filter(function (r) { return String(r[2]).trim().toUpperCase() === "TRUE"; })
      .map(function (r) { return { code: String(r[0]).trim(), name: r[1] }; });
    return respond({ ok: true, cases: cases });
  }

  if (action === "config") {
    var list = ss.getSheetByName(SHEET_LIST);
    var rows = list.getDataRange().getValues();
    rows.shift(); // 表頭
    var items = rows
      .filter(function (r) { return String(r[3]).trim().toUpperCase() === "TRUE" && String(r[CASE_COL_LIST - 1] || "").trim() === caseId; })
      .map(function (r) { return { category: r[0], name: unsanitizeCell(r[1]), workCode: r[2] }; });
    var basicObj = {
      "業主": caseInfo.owner || "",
      "工程名稱": caseInfo.name || "",
      "合約金額（元）": caseInfo.contract || "",
      "開工日期（YYYY/MM/DD）": caseInfo.startDate || "",
      "公司名稱": caseInfo.company || "",
    };
    return respond({ ok: true, items: items, basic: basicObj });
  }

  if (action === "day") {
    var date = String(e.parameter.date || "").trim();
    var headerSheet = ss.getSheetByName(SHEET_HEADER);
    var hRow = findRowByKey(headerSheet, 1, date, CASE_COL_HEADER, caseId);
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
      var rv = recordSheet.getRange(2, 1, lastRow - 1, CASE_COL_RECORD).getValues();
      records = rv.filter(function (r) { return normalizeDate(r[0]) === date && String(r[CASE_COL_RECORD - 1] || "").trim() === caseId; })
        .map(function (r) { return { name: unsanitizeCell(r[1]), am: r[2], pm: r[3] }; });
    }
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    var attLastRow = attSheet.getLastRow();
    var attendance = [];
    if (attLastRow >= 2) {
      var av = attSheet.getRange(2, 1, attLastRow - 1, CASE_COL_ATTENDANCE).getValues();
      attendance = av.filter(function (r) { return normalizeDate(r[0]) === date && String(r[CASE_COL_ATTENDANCE - 1] || "").trim() === caseId; })
        .map(function (r) { return { name: unsanitizeCell(r[1]), am: r[2] === "V", pm: r[3] === "V", amHours: r[4], pmHours: r[5], reason: unsanitizeCell(r[6]) }; });
    }
    return respond({ ok: true, header: header, records: records, attendance: attendance });
  }

  if (action === "cumulative") {
    var upTo = String(e.parameter.upTo || "").trim();

    // 材料是直接加總用量，工種/機具是「工天數」邏輯(上午+下午)/2 —— 兩種算法不同，
    // 要先查清單分頁知道每個項目屬於哪一類，才能套對公式。清單本身依案場過濾（工種/機具/材料清單各案場獨立）。
    var listSheet = ss.getSheetByName(SHEET_LIST);
    var listRows = listSheet.getDataRange().getValues();
    listRows.shift();
    var categoryByName = Object.create(null); // 項目名稱當 key：用無原型物件，避免 constructor/__proto__ 等名稱汙染
    listRows.forEach(function (r) {
      if (String(r[CASE_COL_LIST - 1] || "").trim() === caseId) categoryByName[unsanitizeCell(r[1])] = r[0];
    });

    var recordSheet2 = ss.getSheetByName(SHEET_RECORD);
    var lastRow2 = recordSheet2.getLastRow();
    var totals = Object.create(null);
    var cumWarnings = [], skippedNames = Object.create(null);
    if (lastRow2 >= 2) {
      var rv2 = recordSheet2.getRange(2, 1, lastRow2 - 1, CASE_COL_RECORD).getValues();
      rv2.forEach(function (r) {
        if (String(r[CASE_COL_RECORD - 1] || "").trim() !== caseId) return;
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
      headerSheet3.getRange(2, 1, lastRow3 - 1, CASE_COL_HEADER).getValues().forEach(function (r) {
        if (String(r[CASE_COL_HEADER - 1] || "").trim() !== caseId) return;
        var n = unsanitizeCell(r[6]).trim();
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
      attSheet3.getRange(2, 1, attLastRow3 - 1, CASE_COL_ATTENDANCE).getValues().forEach(function (r) {
        if (String(r[CASE_COL_ATTENDANCE - 1] || "").trim() !== caseId) return;
        var n = unsanitizeCell(r[1]).trim();
        if (n && !seen2[n]) { seen2[n] = true; pNames.push(n); }
      });
    }
    return respond({ ok: true, names: pNames });
  }

  // 選工記錄「類別」輸入框的建議清單：固定 4 個預設值 + 該案場後台曾經用過的自訂值（依出現順序，不重複）
  if (action === "internalCategories") {
    var categories = ["工種", "機具", "材料", "其他"];
    var seenCat = Object.create(null);
    categories.forEach(function (c) { seenCat[c] = true; });
    var internalSheet3 = ss.getSheetByName(SHEET_INTERNAL);
    var internalLastRow3 = internalSheet3.getLastRow();
    if (internalLastRow3 >= 2) {
      internalSheet3.getRange(2, 1, internalLastRow3 - 1, CASE_COL_INTERNAL).getValues().forEach(function (r) {
        if (String(r[CASE_COL_INTERNAL - 1] || "").trim() !== caseId) return;
        var c = unsanitizeCell(r[2]).trim();
        if (c && !seenCat[c]) { seenCat[c] = true; categories.push(c); }
      });
    }
    return respond({ ok: true, names: categories });
  }

  // 主任／工安「填寫過的名字」清單（上午、下午欄合併、不重複），依最近一次出現的日期由新到舊排列，給網頁的下拉選單用
  if (action === "leaderNames") {
    var leaderSheet = ss.getSheetByName(SHEET_HEADER);
    var leaderLast = leaderSheet.getLastRow();
    var dirLatest = Object.create(null), safeLatest = Object.create(null);
    if (leaderLast >= 2) {
      leaderSheet.getRange(2, 1, leaderLast - 1, CASE_COL_HEADER).getValues().forEach(function (r) {
        if (String(r[CASE_COL_HEADER - 1] || "").trim() !== caseId) return;
        var ld = normalizeDate(r[0]);
        [[9, dirLatest], [10, dirLatest], [11, safeLatest], [12, safeLatest]].forEach(function (p) {
          var n = unsanitizeCell(r[p[0]]).trim();
          if (n && (!(n in p[1]) || ld > p[1][n])) p[1][n] = ld;
        });
      });
    }
    var newestFirst = function (m) { return Object.keys(m).sort(function (a, b) { return m[a] < m[b] ? 1 : m[a] > m[b] ? -1 : 0; }); };
    return respond({ ok: true, directors: newestFirst(dirLatest), safeties: newestFirst(safeLatest) });
  }

  // 工項完成量頁讀取：項目主表（啟用中，或該日已有資料的停用項目）＋ 該日已存的完成量 ＋ 到前一天為止的累計（期初已完成＋之前各日加總）
  if (action === "workItems") {
    var wDate = String(e.parameter.date || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(wDate)) return respond({ ok: false, error: "日期格式錯誤" });
    var wi = xlsxLoadItems_(ss, caseId);
    var wPrior = {}, wValues = {};
    wi.master.forEach(function (m) { wPrior[m.seq] = m.opening; });
    Object.keys(wi.byDate).forEach(function (d) {
      Object.keys(wi.byDate[d]).forEach(function (sq) {
        if (!(sq in wPrior)) return; // 不在主表的序號（主表被改掉）略過
        if (d < wDate) wPrior[sq] += wi.byDate[d][sq];
        else if (d === wDate) wValues[sq] = wi.byDate[d][sq];
      });
    });
    Object.keys(wPrior).forEach(function (sq) { wPrior[sq] = Math.round(wPrior[sq] * 100) / 100; });
    var wItems = wi.master
      .filter(function (m) { return m.active || (m.seq in wValues); })
      .map(function (m) { return { seq: m.seq, name: m.name, unit: m.unit, qty: m.qty }; });
    return respond({ ok: true, items: wItems, prior: wPrior, values: wValues });
  }

  // 目前該案場 xlsx 輸出檔的網址；檔案還沒產生（尚未送出過日報）時 url 為空字串
  if (action === "xlsxUrl") {
    try {
      var xlsxFile = xlsxFindOutput_(caseId, caseInfo);
      return respond({ ok: true, url: xlsxFile ? xlsxFile.getUrl() : "" });
    } catch (err) {
      return respond({ ok: false, error: String(err) });
    }
  }

  return respond({ ok: true, message: "daily-report API is running" });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 後台管理：一鍵刪除某天資料（試算表上方「日報管理」選單）----------
// 給接手的後台管理人員用：不用逐張資料表找列、也不用碰 xlsx。只有能編輯這份試算表的人看得到選單。

function onOpen() {
  SpreadsheetApp.getUi().createMenu("日報管理")
    .addItem("更新全部日報（套用案場設定）", "adminRebuildAll")
    .addItem("刪除某天日報資料…", "adminDeleteDay")
    .addToUi();
}

// 選單「更新全部日報」：修改「案場設定」的公司名稱/業主/工程名稱/合約金額/開工日期後按這個，
// 依試算表資料重建所有啟用中案場的 xlsx（舊日期＋新日期的抬頭一併更新；日報內容與累計不變）。
// 逐案場各自 try/catch：某案場失敗（例如還沒有任何日報）不影響其他案場。
function adminRebuildAll() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (firstMissingSheet(ss, [SHEET_CASES])) {
    ui.alert("無法更新", NOT_SETUP_ERROR, ui.ButtonSet.OK);
    return;
  }
  var rows = ss.getSheetByName(SHEET_CASES).getDataRange().getValues();
  rows.shift();
  var active = rows.filter(function (r) { return String(r[2]).trim().toUpperCase() === "TRUE"; });
  if (!active.length) {
    ui.alert("無法更新", "「案場設定」目前沒有啟用中的案場。", ui.ButtonSet.OK);
    return;
  }
  var names = active.map(function (r) { return "・" + String(r[0]).trim() + "（" + r[1] + "）"; }).join("\n");
  if (ui.alert("更新全部日報",
    "將依「案場設定」的最新資料，重建以下案場的 xlsx（每天的分頁都會更新抬頭）：\n\n" + names +
    "\n\n日報內容與累計不會改變。Drive 上的 xlsx 只當輸出檔，若曾手動改過格子會被覆蓋。\n確定要更新嗎？",
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var lines = active.map(function (r) {
    var id = String(r[0]).trim();
    try {
      var res = xlsxRun_(null, id, getCaseRow_(ss, id));
      if (res.removed) return "・" + id + "：沒有日報，略過";
      return "・" + id + "：已更新 " + res.sheets + " 個分頁" + (res.warnings.length ? "（有 " + res.warnings.length + " 則警告，見執行記錄）" : "");
    } catch (e) {
      return "・" + id + "：⚠ 失敗：" + e;
    }
  });
  ui.alert("更新完成", lines.join("\n"), ui.ButtonSet.OK);
}

// 接受 2026-09-21、2026/9/21、115.9.21（民國，年為 3 位數）；不合法回傳空字串，合法回傳 yyyy-mm-dd
function parseAdminDate(input) {
  var m = /^(\d{3,4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/.exec(String(input || "").trim());
  if (!m) return "";
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if (y < 1000) y += 1911;
  var dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return "";
  return y + "-" + ("0" + mo).slice(-2) + "-" + ("0" + d).slice(-2);
}

// 該分頁中「日期欄（A欄）等於 date 且 案場欄等於 caseId」的所有列號（由小到大）
function adminDayRows_(sheet, date, caseCol, caseId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var vals = sheet.getRange(2, 1, lastRow - 1, caseCol).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    if (normalizeDate(vals[i][0]) === date && String(vals[i][caseCol - 1] || "").trim() === caseId) rows.push(i + 2);
  }
  return rows;
}

// 統計某案場某天現有多少資料（不修改任何東西）：{header, record, attendance, xlsx(該日在該案場 xlsx 有沒有分頁)}
function adminCountDay(date, caseId, caseInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var xlsx = false;
  try { xlsx = xlsxHasSheetForDate(date, caseId, caseInfo); } catch (err) { xlsx = false; }
  return {
    header: adminDayRows_(ss.getSheetByName(SHEET_HEADER), date, CASE_COL_HEADER, caseId).length,
    record: adminDayRows_(ss.getSheetByName(SHEET_RECORD), date, CASE_COL_RECORD, caseId).length,
    attendance: adminDayRows_(ss.getSheetByName(SHEET_ATTENDANCE), date, CASE_COL_ATTENDANCE, caseId).length,
    xlsx: xlsx,
  };
}

// 刪除某案場某天在三張資料表的所有列，並同步該案場的 xlsx（移除該日分頁、重算之後日期的累計）。
// 與日報送出共用同一把鎖，避免有人正好在送出時互相干擾。
// @return {ok, deleted:{header,record,attendance}, xlsx:{ok, skipped?, removedFile?, sheets?, error?}} 或 {ok:false, error}
function adminDeleteDayCore(date, caseId, caseInfo) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) return { ok: false, error: "日期格式錯誤" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (firstMissingSheet(ss, [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE])) return { ok: false, error: NOT_SETUP_ERROR };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return { ok: false, error: "系統忙碌中（可能有人正在送出日報），請稍後再試" };
  }
  try {
    var deleted = {};
    [[SHEET_HEADER, "header", CASE_COL_HEADER], [SHEET_RECORD, "record", CASE_COL_RECORD], [SHEET_ATTENDANCE, "attendance", CASE_COL_ATTENDANCE]].forEach(function (p) {
      var sheet = ss.getSheetByName(p[0]);
      var rows = adminDayRows_(sheet, date, p[2], caseId);
      for (var i = rows.length - 1; i >= 0; i--) sheet.deleteRow(rows[i]); // 由下往上刪，列號才不會位移
      deleted[p[1]] = rows.length;
    });

    var xlsx;
    try {
      if (!xlsxFindOutput_(caseId, caseInfo)) {
        xlsx = { ok: true, skipped: true }; // xlsx 還沒建立過，不需處理
      } else {
        var xr = xlsxUpdateForDate(date, caseId, caseInfo);
        xlsx = { ok: true, removedFile: !!xr.removed, sheets: xr.sheets };
      }
    } catch (xerr) {
      xlsx = { ok: false, error: String(xerr) }; // 資料已刪除；xlsx 失敗時可再執行 xlsxRebuildAll(caseId) 補救
    }
    return { ok: true, deleted: deleted, xlsx: xlsx };
  } finally {
    lock.releaseLock();
  }
}

// 選單「刪除某天日報資料…」的畫面流程：先選案場 → 輸入日期 → 顯示將刪除的筆數並二次確認 → 執行 → 顯示結果
function adminDeleteDay() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (firstMissingSheet(ss, [SHEET_CASES])) {
    ui.alert("無法刪除", NOT_SETUP_ERROR, ui.ButtonSet.OK);
    return;
  }
  var caseSheet = ss.getSheetByName(SHEET_CASES);
  var caseRows = caseSheet.getDataRange().getValues();
  caseRows.shift();
  var activeCases = caseRows.filter(function (r) { return String(r[2]).trim().toUpperCase() === "TRUE"; });
  if (!activeCases.length) {
    ui.alert("無法刪除", "「案場設定」目前沒有啟用中的案場。", ui.ButtonSet.OK);
    return;
  }
  var caseList = activeCases.map(function (r) { return String(r[0]).trim() + "（" + r[1] + "）"; }).join("\n");
  var caseResp = ui.prompt(
    "刪除某天日報資料 - 第1步：選案場",
    "請輸入案場代碼：\n\n" + caseList,
    ui.ButtonSet.OK_CANCEL
  );
  if (caseResp.getSelectedButton() !== ui.Button.OK) return;
  var caseId = caseResp.getResponseText().trim();
  var caseInfo = getCaseRow_(ss, caseId);
  if (!caseInfo) {
    ui.alert("案場代碼不正確", "請輸入清單中列出的案場代碼。", ui.ButtonSet.OK);
    return;
  }

  var resp = ui.prompt(
    "刪除某天日報資料 - 第2步：選日期（案場：" + caseInfo.name + "）",
    "請輸入要刪除的日期（例如 2026-09-21 或民國 115.9.21）。\n\n" +
    "會刪除該天在「日報頭」「日報記錄」「本工出勤」的資料，並從該案場的 xlsx 移除該天的分頁。\n" +
    "（「內部記錄」分頁不受影響）",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var date = parseAdminDate(resp.getResponseText());
  if (!date) {
    ui.alert("日期格式不正確", "請輸入像 2026-09-21 或 115.9.21 的日期。", ui.ButtonSet.OK);
    return;
  }
  if (firstMissingSheet(ss, [SHEET_HEADER, SHEET_RECORD, SHEET_ATTENDANCE])) {
    ui.alert("無法刪除", NOT_SETUP_ERROR, ui.ButtonSet.OK);
    return;
  }

  var n = adminCountDay(date, caseId, caseInfo);
  if (!n.header && !n.record && !n.attendance && !n.xlsx) {
    ui.alert("找不到資料", date + "（" + caseInfo.name + "）在三張資料表和 xlsx 都沒有資料，沒有刪除任何東西。", ui.ButtonSet.OK);
    return;
  }
  var sheetName = XlsxBuilder.sheetNameForDate(date);
  var confirmText = "即將刪除【" + caseInfo.name + "】" + date + " 的資料：\n\n" +
    "・日報頭　　 " + n.header + " 列\n" +
    "・日報記錄　 " + n.record + " 列\n" +
    "・本工出勤　 " + n.attendance + " 列\n" +
    "・xlsx 分頁「" + sheetName + "」：" + (n.xlsx ? "會移除（之後日期的累計會重新計算）" : "沒有") + "\n" +
    "（「內部記錄」分頁不受影響）\n\n" +
    "刪除後無法在網頁上復原（可用 Google 試算表的「檔案 → 版本記錄」還原）。\n確定要刪除嗎？";
  if (ui.alert("確認刪除 " + date, confirmText, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var r = adminDeleteDayCore(date, caseId, caseInfo);
  if (!r.ok) {
    ui.alert("刪除失敗", r.error, ui.ButtonSet.OK);
    return;
  }
  var msg = "已刪除【" + caseInfo.name + "】" + date + "：日報頭 " + r.deleted.header + " 列、日報記錄 " + r.deleted.record + " 列、本工出勤 " + r.deleted.attendance + " 列。\n\n";
  if (!r.xlsx.ok) msg += "⚠ xlsx 更新失敗：" + r.xlsx.error + "\n資料已刪除，請到 Apps Script 執行 xlsxRebuildAll(\"" + caseId + "\") 補救。";
  else if (r.xlsx.skipped) msg += "xlsx 尚未建立，不需處理。";
  else if (r.xlsx.removedFile) msg += "xlsx 已沒有任何分頁，已把檔案移到 Drive 垃圾桶（下次送出日報時會自動重新建立）。";
  else msg += "xlsx 已更新（移除該日分頁，並重算之後日期的累計）。";
  ui.alert("刪除完成", msg, ui.ButtonSet.OK);
}
