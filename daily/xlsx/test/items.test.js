// 工項完成量＋主任/工安名單 的測試：node items.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
// 重點是「資料保護」：送出工項完成量不能動到既有日報資料與 xlsx 內任何日報分頁。
const fs = require("fs");
const XLSX = require("xlsx");
const { makeGlueEnv, readZip, writeZip } = require("./fake-gas.js");
const TEMPLATE_BYTES = fs.readFileSync("D:/洲際液化天然氣接收站AI/日報範本.xlsx");

const R = [];
function T(id, name, fn) {
  try { const r = fn(); const pass = r === true || r === undefined; R.push([id, pass ? "PASS" : "FAIL", name, pass ? "" : (typeof r === "string" ? r : JSON.stringify(r))]); }
  catch (e) { R.push([id, "ERROR", name, String(e && e.stack || e)]); }
}
const eq = (a, b) => (JSON.stringify(a) === JSON.stringify(b) ? true : "got " + JSON.stringify(a) + " expected " + JSON.stringify(b));

const CASE_ID = "lng", OTHER = "h2";
const OUTPUT_NAME = "施工日報彙整_測試工程.xlsx";
const HDR = ["日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目", "備註", "填表人", "clientId", "更新時間", "主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)", "案場"];
function baseSheets() {
  return {
    "案場設定": [["案場代碼", "案場名稱", "啟用中", "業主", "合約金額（元）", "開工日期（YYYY/MM/DD）", "公司名稱"], [CASE_ID, "測試工程", "TRUE", "中油", "", "", ""], [OTHER, "別案場", "TRUE", "", "", "", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中", "案場"], ["工種", "公司工", "", "TRUE", CASE_ID], ["工種", "模板工", "", "TRUE", CASE_ID]],
    "日報頭": [HDR,
      ["2026-09-20", "雨", "施工", "第三天", "", "", "王", "c", "t", "甲主任", "甲主任", "乙工安", "乙工安", CASE_ID],
      ["2026-09-22", "晴", "施工", "", "", "", "王", "c", "t", "丙主任", "丙主任", "戊工安", "戊工安", CASE_ID],
      ["2026-09-21", "晴", "施工", "", "", "", "王", "c", "t", "甲主任", "", "", "", CASE_ID],
      ["2026-09-25", "晴", "施工", "", "", "", "王", "c", "t", "他案主任", "", "他案工安", "", OTHER]],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間", "案場"], ["2026-09-20", "公司工", 4, 4, "c", "t", CASE_ID]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間", "案場"]],
    "請款資料": [["案場", "請款日期", "區域", "期別", "本期請款金額", "目前進度", "備註", "發票"]],
    "內部記錄": [["日期", "類型", "類別", "內容", "時間", "備註", "clientId", "更新時間", "案場"]],
    "合約工項": [["案場", "序號", "項目名稱", "單位", "預估數量", "期初已完成", "啟用中"],
      [CASE_ID, 1, "樁頭及錨定處理", "PC", 325, 0, "TRUE"],
      [CASE_ID, 2, "深井排水", "PT", 26, 5, "TRUE"],
      [CASE_ID, 5, "開挖", "m3", 14183, 0, "TRUE"],
      [CASE_ID, 9, "已停用項目", "m3", 100, 0, "FALSE"],
      [OTHER, 1, "別案場項目", "m", 50, 0, "TRUE"]],
    "工項完成量": [["日期", "序號", "本日完成量", "clientId", "更新時間", "案場"]],
  };
}

function makeEnv(opts) {
  opts = opts || {};
  const env = makeGlueEnv({ sheetsData: opts.sheets || baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  const data = env.state.sheetsData, cache = {};
  const sheet = (n) => ({
    getMaxRows: () => 1000,
    insertRowsAfter: () => {},
    getLastRow: () => { let l = 0; data[n].forEach((r, i) => { if (r.some((v) => v !== "" && v != null)) l = i + 1; }); return l; },
    getDataRange: () => ({ getValues: () => JSON.parse(JSON.stringify(data[n])) }),
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = data[n][r - 1 + i] || []; const x = []; for (let j = 0; j < nc; j++) x.push(row[c - 1 + j] === undefined ? "" : row[c - 1 + j]); o.push(x); } return o; },
        setValues(vs) { for (let i = 0; i < nr; i++) { while (data[n].length < r + i) data[n].push([]); for (let j = 0; j < nc; j++) data[n][r - 1 + i][c - 1 + j] = vs[i][j]; } },
      };
    },
    deleteRow: (r) => { data[n].splice(r - 1, 1); },
    appendRow: (arr) => { data[n].push(arr.slice()); },
  });
  env.ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getId: () => "SS1", getSheetByName: (n) => (data[n] ? sheet(n) : null) }) };
  env.ctx.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  env.ctx.CacheService = { getScriptCache: () => ({ get: (k) => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
  env.ctx.ContentService = { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: "JSON" } };
  env.post = (payload) => JSON.parse(env.call("doPost")({ postData: { contents: JSON.stringify(payload) } }).s);
  env.get = (params) => JSON.parse(env.call("doGet")({ parameter: params }).s);
  return env;
}
const F = (env) => Object.values(env.state.files).find((f) => f.name === OUTPUT_NAME && !f.trashed);
const book = (env) => XLSX.read(F(env).bytes);
const sheetXmls = (env) => { // {分頁名: xml}，拿掉 tabSelected 以便比對
  const parts = {}; readZip(F(env).bytes).forEach((e) => { parts[e.name] = e.data.toString("utf8"); });
  const wb = parts["xl/workbook.xml"], rels = parts["xl/_rels/workbook.xml.rels"], out = {}, re = /<sheet name="([^"]*)" sheetId="\d+" r:id="([^"]+)"\/>/g; let m;
  while ((m = re.exec(wb))) { const t = new RegExp('<Relationship Id="' + m[2] + '"[^>]*Target="([^"]+)"').exec(rels)[1]; out[m[1]] = parts["xl/" + t].replace(/\s+tabSelected="1"/g, ""); }
  return out;
};
const saveItems = (date, items, clear, extra) => Object.assign({ action: "saveWorkItems", date, caseId: CASE_ID, clientId: "cid", items, clearItems: clear || [] }, extra || {});
const submitDay = (date, cong) => ({
  date, caseId: CASE_ID, clientId: "cid", header: { weather: "晴", status: "施工", todayWork: "新的一天", reporter: "王", directorAm: "甲", directorPm: "甲", safetyAm: "乙", safetyPm: "乙" },
  items: [{ name: "公司工", am: cong, pm: cong }], attendance: [],
});
const doneRows = (env, caseId) => env.state.sheetsData["工項完成量"].slice(1).filter((r) => r[5] === (caseId || CASE_ID)).map((r) => [r[0], r[1], r[2]]);

(function main() {
  // ---------- 讀取 ----------
  const env = makeEnv();
  T("I1", "workItems：只回該案場啟用中的項目（序號/名稱/單位/預估數量）；累計＝期初已完成；該日無資料", () => {
    const r = env.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-07" });
    return eq([r.ok, r.items, r.prior, r.values], [true, [{ seq: 1, name: "樁頭及錨定處理", unit: "PC", qty: 325 }, { seq: 2, name: "深井排水", unit: "PT", qty: 26 }, { seq: 5, name: "開挖", unit: "m3", qty: 14183 }], { 1: 0, 2: 5, 5: 0, 9: 0 }, {}]);
  });
  T("I2", "workItems：日期格式錯誤、案場不存在 → 友善錯誤", () => eq([env.get({ action: "workItems", caseId: CASE_ID, date: "11/07" }).ok, env.get({ action: "workItems", caseId: "nope", date: "2026-11-07" }).ok], [false, false]));

  // ---------- 先送出日報，建立日報分頁，供之後比對「有沒有被動到」 ----------
  env.post(submitDay("2026-09-20", 4));
  const rd = env.post(submitDay("2026-09-21", 5));
  const snapDay = {}; // 送出工項前，所有日報分頁的 xml
  let dbBefore;
  T("I3", "前置：送出日報後 xlsx 有 115.9.20、115.9.21，且還沒有工項分頁", () => {
    const names = book(env).SheetNames; Object.assign(snapDay, sheetXmls(env));
    dbBefore = JSON.stringify([env.state.sheetsData["日報頭"], env.state.sheetsData["日報記錄"], env.state.sheetsData["本工出勤"], env.state.sheetsData["請款資料"], env.state.sheetsData["內部記錄"], env.state.sheetsData["合約工項"]]);
    return eq([rd.ok, names.includes("115.9.20"), names.includes("115.9.21"), names.some((n) => n.startsWith("工項."))], [true, true, true, false]);
  });

  // ---------- 儲存 ----------
  let r1;
  T("I4", "saveWorkItems 11/07：寫入 2 列（0 值不寫）、回 saved=2、xlsx ok 且有「工項.115.11.07」分頁", () => {
    r1 = env.post(saveItems("2026-11-07", [{ seq: 1, qty: 12 }, { seq: 5, qty: "700.5" }, { seq: 2, qty: 0 }]));
    return eq([r1.ok, r1.saved, r1.cleared, r1.xlsx.ok, /^https:\/\//.test(r1.xlsx.url), doneRows(env), book(env).SheetNames.includes("工項.115.11.07")], [true, 2, 0, true, true, [["2026-11-07", 1, 12], ["2026-11-07", 5, 700.5]], true]);
  });
  T("I5", "工項分頁內容：標題、欄位、前日完成＝期初、本日、累計；停用項目不列；0 值留白", () => {
    const s = book(env).Sheets["工項.115.11.07"];
    return eq([s.A1.v.includes("115.11.7"), s.A4.v, s.H4.v, s.A5.v, s.B5.v, s.D5.v, s.F5.v, s.G5.v, s.B6.v, s.E6.v, s.F6 === undefined, s.G6.v, s.B7.v, s.F7.v, s.A8 === undefined || s.A8.v !== 9], [true, "序號", "備註", 1, "樁頭及錨定處理", 325, 12, 12, "深井排水", 5, true, 5, "開挖", 700.5, true]);
  });
  T("I6", "【資料保護】送出工項後：日報頭／日報記錄／本工出勤／請款資料／內部記錄／合約工項 的內容完全沒變", () => eq(JSON.stringify([env.state.sheetsData["日報頭"], env.state.sheetsData["日報記錄"], env.state.sheetsData["本工出勤"], env.state.sheetsData["請款資料"], env.state.sheetsData["內部記錄"], env.state.sheetsData["合約工項"]]) === dbBefore, true));
  T("I7", "【資料保護】xlsx 內既有日報分頁（115.9.20／115.9.21）的 XML 與送出工項前逐字相同", () => {
    const now = sheetXmls(env);
    return eq(["115.9.20", "115.9.21"].map((n) => now[n] === snapDay[n]), [true, true]);
  });
  T("I8", "同一天再送一次＝更新（不重複新增列）；累計與分頁跟著更新", () => {
    const r = env.post(saveItems("2026-11-07", [{ seq: 1, qty: 20 }]));
    const s = book(env).Sheets["工項.115.11.07"];
    return eq([r.ok, r.saved, doneRows(env), s.F5.v, s.F7.v], [true, 1, [["2026-11-07", 1, 20], ["2026-11-07", 5, 700.5]], 20, 700.5]);
  });
  T("I9", "沒被列出的序號不會被動（只送序號 1，序號 5 的 700.5 仍在）", () => eq(doneRows(env).find((x) => x[1] === 5)[2], 700.5));
  T("I10", "11/08：前日完成＝11/07 累計；累計超過預估 → 備註「超出預估」（只提醒仍寫入）", () => {
    const r = env.post(saveItems("2026-11-08", [{ seq: 1, qty: 310 }, { seq: 2, qty: 3 }]));
    const s = book(env).Sheets["工項.115.11.08"];
    return eq([r.ok, s.E5.v, s.F5.v, s.G5.v, s.H5.v, s.E6.v, s.G6.v, s.H6 === undefined || !s.H6.v], [true, 20, 310, 330, "超出預估 +5", 5, 8, true]);
  });
  T("I11", "workItems 11/09：累計＝期初＋11/07＋11/08；11/08 的 values 回傳該日資料", () => {
    const a = env.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-09" }), b = env.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-08" });
    return eq([a.prior["1"], a.prior["2"], a.prior["5"], a.values, b.values, b.prior["1"]], [330, 8, 700.5, {}, { 1: 310, 2: 3 }, 20]);
  });
  T("I12", "clearItems 只刪被點名的那一列；其他列與其他日期不動", () => {
    const r = env.post(saveItems("2026-11-08", [], [2]));
    return eq([r.ok, r.cleared, doneRows(env).filter((x) => x[0] === "2026-11-08"), doneRows(env).filter((x) => x[0] === "2026-11-07").length], [true, 1, [["2026-11-08", 1, 310]], 2]);
  });
  T("I13", "案場隔離：另一案場有自己的項目與資料，互不影響", () => {
    const r = env.post(saveItems("2026-11-07", [{ seq: 1, qty: 7 }], [], { caseId: OTHER }));
    const a = env.get({ action: "workItems", caseId: OTHER, date: "2026-11-08" });
    return eq([r.ok, doneRows(env, OTHER), a.items.map((x) => x.name), a.prior, doneRows(env).length], [true, [["2026-11-07", 1, 7]], ["別案場項目"], { 1: 7 }, 3]);
  });

  // ---------- 驗證：全有或全無 ----------
  T("I14", "不合法一律整批拒絕、不寫入：序號不在主表、重複序號、負數、非數字、日期錯誤、同時寫入與清除", () => {
    const before = JSON.stringify(env.state.sheetsData["工項完成量"]), n = env.state.patches;
    const bad = [
      saveItems("2026-11-09", [{ seq: 1, qty: 1 }, { seq: 99, qty: 1 }]),
      saveItems("2026-11-09", [{ seq: 1, qty: 1 }, { seq: 1, qty: 2 }]),
      saveItems("2026-11-09", [{ seq: 1, qty: -3 }]),
      saveItems("2026-11-09", [{ seq: 1, qty: "abc" }]),
      saveItems("2026-13-45", [{ seq: 1, qty: 1 }]),
      saveItems("2026-11-09", [{ seq: 1, qty: 5 }], [1]),
      saveItems("2026-11-09", [], [99]),
    ].map((p) => env.post(p).ok);
    return eq([bad, JSON.stringify(env.state.sheetsData["工項完成量"]) === before, env.state.patches === n], [[false, false, false, false, false, false, false], true, true]);
  });
  T("I15", "設了 SUBMIT_PASSWORD：沒帶／帶錯密碼被擋、不寫入；正確才通過", () => {
    env.state.props.SUBMIT_PASSWORD = "pw";
    const before = JSON.stringify(env.state.sheetsData["工項完成量"]);
    const a = env.post(saveItems("2026-11-09", [{ seq: 1, qty: 1 }])), b = env.post(saveItems("2026-11-09", [{ seq: 1, qty: 1 }], [], { password: "x" }));
    const same = JSON.stringify(env.state.sheetsData["工項完成量"]) === before;
    const c = env.post(saveItems("2026-11-09", [{ seq: 1, qty: 1 }], [], { password: "pw" }));
    delete env.state.props.SUBMIT_PASSWORD;
    return eq([a.ok, b.ok, same, c.ok], [false, false, true, true]);
  });

  // ---------- 與日報送出／全量重建的互動 ----------
  T("I16", "之後再送出日報（9/22）：工項分頁仍在且內容正確、新日報分頁產生、舊日報分頁仍在", () => {
    const r = env.post(submitDay("2026-09-22", 6));
    const b = book(env), s = b.Sheets["工項.115.11.08"];
    return eq([r.ok, r.xlsx.ok, b.SheetNames.includes("115.9.22"), b.SheetNames.includes("115.9.21"), b.SheetNames.includes("工項.115.11.07"), s.G5.v], [true, true, true, true, true, 330]);
  });
  T("I17", "分頁順序：日報分頁依日期在前、工項分頁依日期在後；開檔停在最新日報分頁", () => {
    const names = book(env).SheetNames, wb = readZip(F(env).bytes).find((e) => e.name === "xl/workbook.xml").data.toString("utf8");
    const activeTab = Number(/activeTab="(\d+)"/.exec(wb)[1]);
    return eq([names.filter((n) => !n.startsWith("工項.")).concat(names.filter((n) => n.startsWith("工項."))).join() === names.join(), names[activeTab]], [true, "115.9.22"]);
  });
  T("I18", "全量重建（xlsxRebuildAll）：工項分頁由資料庫重新產生，不會消失", () => {
    env.call("xlsxRebuildAll")(CASE_ID);
    const b = book(env);
    return eq([b.SheetNames.filter((n) => n.startsWith("工項.")), b.Sheets["工項.115.11.07"].F5.v], [["工項.115.11.07", "工項.115.11.08", "工項.115.11.09"], 20]);
  });
  T("I19", "範本改版（樣式表不同）：不覆蓋雲端檔、回 xlsx 錯誤，但工項資料已存入試算表", () => {
    const f = F(env), entries = readZip(f.bytes).map((e) => ({ name: e.name, data: e.name === "xl/styles.xml" ? Buffer.from(e.data.toString("utf8") + "<!-- changed -->") : e.data }));
    f.bytes = writeZip(entries);
    const bytesBefore = Buffer.from(f.bytes), rows = doneRows(env).length;
    const r = env.post(saveItems("2026-11-10", [{ seq: 5, qty: 1 }]));
    return eq([r.ok, r.xlsx.ok, /範本已改版/.test(r.xlsx.error || ""), doneRows(env).length === rows + 1, Buffer.compare(bytesBefore, f.bytes) === 0], [true, false, true, true, true]);
  });
  T("I20", "尚未有 xlsx 檔（沒送出過日報）：送出工項會建立只含工項分頁的檔案；之後送日報接續使用同一個檔", () => {
    const e2 = makeEnv();
    const r = e2.post(saveItems("2026-11-07", [{ seq: 1, qty: 3 }]));
    const first = F(e2), onlyItems = book(e2).SheetNames.every((n) => n.startsWith("工項."));
    const r2 = e2.post(submitDay("2026-09-21", 4));
    const b = book(e2);
    return eq([r.ok, r.xlsx.ok, onlyItems, r2.ok, F(e2) === first, b.SheetNames.includes("115.9.21"), b.SheetNames.includes("工項.115.11.07")], [true, true, true, true, true, true, true]);
  });
  T("I21", "沒有任何資料可寫（刪光工項、也沒有 xlsx）：回 ok 且 skipped，不建立空檔", () => {
    const e3 = makeEnv();
    const r = e3.post(saveItems("2026-11-07", [], [1]));
    return eq([r.ok, r.xlsx.ok, r.xlsx.skipped, !!F(e3)], [true, true, true, false]);
  });
  T("I22", "既有日報被刪光最後一個分頁的流程不受影響：工項分頁存在時 xlsx 不會被丟垃圾桶（檔案仍可用）", () => {
    const e4 = makeEnv();
    e4.post(submitDay("2026-09-21", 4)); e4.post(saveItems("2026-11-07", [{ seq: 1, qty: 3 }]));
    const ok = e4.call("adminDeleteDayCore")("2026-09-21", CASE_ID, { name: "測試工程" });
    return eq([ok.ok, !!F(e4), book(e4).SheetNames.some((n) => n.startsWith("工項."))], [true, true, true]);
  });

  T("I28", "【資料保護】第一次寫入工項分頁前，自動在同資料夾留一份「上線前備份」（位元組與原檔相同），之後不再重複備份", () => {
    const e8 = makeEnv();
    e8.post(submitDay("2026-09-21", 5));
    const before = Buffer.from(F(e8).bytes), count = () => Object.values(e8.state.files).filter((x) => /上線前備份/.test(x.name)).length;
    const none = count();
    e8.post(saveItems("2026-11-07", [{ seq: 1, qty: 3 }]));
    const bk = Object.values(e8.state.files).find((x) => /上線前備份/.test(x.name));
    e8.post(saveItems("2026-11-08", [{ seq: 1, qty: 4 }]));
    return eq([none, count(), !!bk && Buffer.compare(bk.bytes, before) === 0, Buffer.compare(F(e8).bytes, before) !== 0], [0, 1, true, true]);
  });

  // ---------- 主任／工安名單 ----------
  T("I23", "leaderNames：上午下午合併、不重複、依最近日期由新到舊；不含其他案場", () => {
    const r = makeEnv().get({ action: "leaderNames", caseId: CASE_ID });
    return eq([r.ok, r.directors, r.safeties], [true, ["丙主任", "甲主任"], ["戊工安", "乙工安"]]);
  });
  T("I24", "leaderNames：另一案場只看到自己的名字；沒有資料時回空陣列", () => {
    const a = env.get({ action: "leaderNames", caseId: OTHER }), b = env.get({ action: "leaderNames", caseId: "nope" });
    return eq([a.directors, a.safeties, b.ok], [["他案主任"], ["他案工安"], false]);
  });
  T("I25", "尚未執行 setupSheets（缺分頁）：workItems／leaderNames 回友善錯誤而不是丟例外", () => {
    const sheets = baseSheets(); delete sheets["合約工項"]; delete sheets["工項完成量"];
    const e5 = makeEnv({ sheets });
    const a = e5.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-07" }), p = e5.post(saveItems("2026-11-07", [{ seq: 1, qty: 1 }]));
    return eq([a.ok, /setupSheets/.test(a.error), p.ok, /setupSheets/.test(p.error)], [false, true, false, true]);
  });
  T("I26", "種子資料：ITEM_MASTER_SEED 共 32 列、序號 1～32，與日報原表 DPR 的名稱／單位／預估數量一致", () => {
    const seed = env.call("ITEM_MASTER_SEED");
    const dpr = XLSX.readFile("D:/洲際液化天然氣接收站/2.施工日報/Talin Project DPR for SS01_02_1151106_清理版2.xlsx").Sheets["1106"];
    const bad = [];
    seed.forEach((s, i) => {
      const r = 11 + i;
      const name = String(dpr["K" + r].v).trim().replace(/[,，]+$/, "").trim(), unit = String(dpr["M" + r].v).trim().toLowerCase();
      if (s[0] !== dpr["Q" + r].v || s[1] !== name || String(s[2]).toLowerCase() !== unit || s[3] !== dpr["L" + r].v) bad.push(s[0]);
    });
    return eq([seed.length, bad], [32, []]);
  });
  T("I27", "既有功能不受影響：送出日報的回應與舊版相同欄位（ok/date/xlsx）", () => {
    const e6 = makeEnv(), r = e6.post(submitDay("2026-09-21", 5));
    return eq([r.ok, r.date, r.xlsx.ok], [true, "2026-09-21", true]);
  });

  // ---------- 計算式（2026-11-13）：量體類工項的計算式文字隨完成量一起存、讀、輸出到 xlsx ----------
  const fxEnv = makeEnv();
  fxEnv.state.sheetsData["工項完成量"].push(["2026-11-06", 5, 50, "old", "t", CASE_ID]); // 舊格式資料列（只有 6 欄、沒有計算式）
  T("I29", "舊格式資料列（無計算式欄）：workItems 回 formulas 為空，values 照舊", () => {
    const r = fxEnv.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-06" });
    return eq([r.ok, r.values, r.formulas], [true, { 5: 50 }, {}]);
  });
  T("I30", "saveWorkItems 帶計算式：存到第 7 欄、workItems 讀回、xlsx「工項」分頁 I 欄顯示；沒計算式的列 I 欄留白", () => {
    const r = fxEnv.post(saveItems("2026-11-07", [{ seq: 1, qty: 3 }, { seq: 5, qty: 112.8, formula: "12.5×8+3.2×4" }]));
    const g = fxEnv.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-07" });
    const s = book(fxEnv).Sheets["工項.115.11.07"];
    return eq([r.ok, r.saved, fxEnv.state.sheetsData["工項完成量"].filter((x) => x[0] === "2026-11-07" && x[1] === 5)[0][6], g.formulas, s.I4.v, s.I5 === undefined || !s.I5.v, s.I7.v, s.F7.v],
      [true, 2, "12.5×8+3.2×4", { 5: "12.5×8+3.2×4" }, "本日計算式", true, "12.5×8+3.2×4", 112.8]);
  });
  T("I31", "計算式開頭是 - 時存入防護用前導單引號，讀回與 xlsx 都還原成原文字", () => {
    fxEnv.post(saveItems("2026-11-08", [{ seq: 5, qty: 7, formula: "-3+10" }]));
    const raw = fxEnv.state.sheetsData["工項完成量"].filter((x) => x[0] === "2026-11-08")[0][6];
    const g = fxEnv.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-08" }), s = book(fxEnv).Sheets["工項.115.11.08"];
    return eq([raw, g.formulas, s.I7.v], ["'-3+10", { 5: "-3+10" }, "-3+10"]);
  });
  T("I32", "計算式含不允許字元（字母、分號、=、函式）或超過 200 字：整批拒絕、不寫入", () => {
    const before = JSON.stringify(fxEnv.state.sheetsData["工項完成量"]);
    const bad = [{ seq: 5, qty: 1, formula: "abc" }, { seq: 5, qty: 1, formula: "1;2" }, { seq: 5, qty: 1, formula: "=1+1" }, { seq: 5, qty: 1, formula: "SUM(1,2)" }, { seq: 5, qty: 1, formula: "1+".repeat(101) }]
      .map((it) => fxEnv.post(saveItems("2026-11-09", [{ seq: 1, qty: 2 }, it])).ok);
    return eq([bad, JSON.stringify(fxEnv.state.sheetsData["工項完成量"]) === before], [[false, false, false, false, false], true]);
  });
  T("I33", "同一天再送一次但不帶計算式＝計算式清空（直接改量的情況）；全形字元與 ÷ 可存", () => {
    fxEnv.post(saveItems("2026-11-07", [{ seq: 5, qty: 100 }]));
    const a = fxEnv.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-07" });
    fxEnv.post(saveItems("2026-11-10", [{ seq: 5, qty: 5, formula: "（１０＋２）÷２.４" }]));
    const b = fxEnv.get({ action: "workItems", caseId: CASE_ID, date: "2026-11-10" });
    return eq([a.values[5], a.formulas, b.formulas], [100, {}, { 5: "（１０＋２）÷２.４" }]);
  });
  T("I34", "【資料保護】舊資料列在新增計算式後保持原樣（6 欄值不變），其他日期的計算式不被動", () => {
    const old = fxEnv.state.sheetsData["工項完成量"].filter((x) => x[0] === "2026-11-06")[0], fx8 = fxEnv.state.sheetsData["工項完成量"].filter((x) => x[0] === "2026-11-08")[0];
    return eq([old.slice(0, 6), old[6] === undefined || old[6] === "", fx8[6]], [["2026-11-06", 5, 50, "old", "t", CASE_ID], true, "'-3+10"]);
  });

  console.log(R.map((r) => "[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      → " + r[3] : "")).join("\n"));
  console.log({ total: R.length, notPass: R.filter((r) => r[1] !== "PASS").length });
})();
