// doPost 接 xlsx 的測試：node dopost.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
// 多案場改版後 doPost 一律要求合法的 caseId；這裡固定用一個假案場 "lng"（案場設定裡的唯一啟用列）。
const fs = require("fs");
const XLSX = require("xlsx");
const { makeGlueEnv } = require("./fake-gas.js");
const TEMPLATE_BYTES = fs.readFileSync("D:/洲際液化天然氣接收站AI/日報範本.xlsx");

const R = [];
function T(id, name, fn) {
  try { const r = fn(); const pass = r === true || r === undefined; R.push([id, pass ? "PASS" : "FAIL", name, pass ? "" : (typeof r === "string" ? r : JSON.stringify(r))]); }
  catch (e) { R.push([id, "ERROR", name, String(e && e.message || e)]); }
}
const eq = (a, b) => (JSON.stringify(a) === JSON.stringify(b) ? true : "got " + JSON.stringify(a) + " expected " + JSON.stringify(b));

const CASE_ID = "lng";
const OUTPUT_NAME = "施工日報彙整_測試工程.xlsx";
const HDR = ["日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目", "備註", "填表人", "clientId", "更新時間", "主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)"];
function baseSheets() {
  return {
    "基本資料": [["業主", "中油"], ["工程名稱", "測試工程"], ["合約金額（元）", ""], ["開工日期（YYYY/MM/DD）", ""], ["公司名稱", ""]],
    "案場設定": [["案場代碼", "案場名稱", "啟用中", "業主", "合約金額（元）", "開工日期（YYYY/MM/DD）", "公司名稱"], [CASE_ID, "測試工程", "TRUE", "中油", "", "", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中", "案場"], ["工種", "公司工", "", "TRUE", CASE_ID], ["工種", "模板工", "", "TRUE", CASE_ID], ["材料", "砂(m³)", "", "TRUE", CASE_ID]],
    "日報頭": [HDR, ["2026-09-20", "雨", "施工", "第三天", "", "", "王", "c", "t", "甲", "甲", "乙", "乙", CASE_ID]],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間", "案場"], ["2026-09-20", "公司工", 4, 4, "c", "t", CASE_ID]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間", "案場"]],
    "請款資料": [["案場", "請款日期", "區域", "期別", "本期請款金額", "目前進度", "備註", "發票", "", "使用說明"], ["", "", "", "", "", "", "", "", "", "說明1"], ["", "", "", "", "", "", "", "", "", "說明2"], ["", "", "", "", "", "", "", "", "", "說明3"], ["", "", "", "", "", "", "", "", "", "說明4"]],
    "內部記錄": [["日期", "類型", "類別", "內容", "時間", "備註", "clientId", "更新時間", "案場"]],
  };
}

// 在假 Drive 環境上補一個可寫入的試算表、鎖、快取與 ContentService，讓真正的 doPost 能跑
function makeEnv(opts) {
  opts = opts || {};
  const env = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
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
  env.ctx.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() { env.state.released = (env.state.released || 0) + 1; } }) };
  env.ctx.CacheService = { getScriptCache: () => ({ get: (k) => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
  env.ctx.ContentService = { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: "JSON" } };
  env.post = (payload) => JSON.parse(env.call("doPost")({ postData: { contents: JSON.stringify(payload) } }).s);
  return env;
}
const F = (env) => Object.values(env.state.files).find((f) => f.name === OUTPUT_NAME && !f.trashed);
const book = (env) => XLSX.read(F(env).bytes);
const submit = (date, cong) => ({
  date, caseId: CASE_ID, clientId: "cid", header: { weather: "晴", status: "施工", todayWork: "新的一天", reporter: "王", directorAm: "甲", directorPm: "甲", safetyAm: "乙", safetyPm: "乙" },
  items: [{ name: "公司工", am: cong, pm: cong }], attendance: [{ name: "林二", am: true, pm: true, amHours: 1, pmHours: 0, reason: "趕工" }],
});

(function main() {
  const env = makeEnv();
  let r1;
  T("P1", "送出 9/21：日報寫入成功，回應含 xlsx:{ok,url,warnings}，檔案有 115.9.20 與 115.9.21", () => {
    r1 = env.post(submit("2026-09-21", 5));
    return eq([r1.ok, r1.date, r1.xlsx.ok, /^https:\/\//.test(r1.xlsx.url), r1.xlsx.warnings, book(env).SheetNames.includes("115.9.21"), env.state.sheetsData["日報頭"].length], [true, "2026-09-21", true, true, [], true, 3]);
  });
  T("P2", "9/21 分頁內容來自剛寫入的資料（公司工 5/5、累計 4+5=9、林二加班1 趕工）", () => {
    const s = book(env).Sheets["115.9.21"];
    return eq([s.A8.v, s.D8.v, s.F8.v, s.M34.v, s.O34.v, s.P34.v], ["公司工", 5, 9, "林二", 1, "趕工"]);
  });
  T("P3", "改較早日期 9/20（4/4→10/10）：9/21 分頁累計一併重寫（10+5=15）", () => {
    env.post(submit("2026-09-20", 10));
    const wb = book(env);
    return eq([wb.Sheets["115.9.20"].F8.v, wb.Sheets["115.9.21"].F8.v, wb.SheetNames], [10, 15, ["115.9.20", "115.9.21"]]);
  });
  T("P4", "xlsx 內部丟例外：日報仍寫入且回 ok:true，xlsx:{ok:false,error}；鎖有釋放", () => {
    const e2 = makeEnv();
    e2.ctx.xlsxUpdateForDate = () => { throw new Error("Drive 掛了"); };
    const r = e2.post(submit("2026-09-22", 3));
    return eq([r.ok, r.date, /Drive 掛了/.test(r.xlsx.error), r.xlsx.ok, e2.state.sheetsData["日報頭"].length, e2.state.sheetsData["日報記錄"].length, e2.state.released], [true, "2026-09-22", true, false, 3, 3, 1]);
  });
  T("P5", "找不到範本檔（真實錯誤路徑）：日報照樣送出成功", () => {
    const e2 = makeEnv();
    Object.values(e2.state.files).forEach((f) => { f.trashed = true; });
    const r = e2.post(submit("2026-09-22", 3));
    return eq([r.ok, r.xlsx.ok, /找不到範本檔/.test(r.xlsx.error), e2.state.sheetsData["日報頭"].length], [true, false, true, 3]);
  });
  T("P6", "警告透傳：xlsxUpdateForDate 回傳 warnings → 原樣出現在 xlsx.warnings", () => {
    const e2 = makeEnv();
    e2.ctx.xlsxUpdateForDate = () => ({ ok: true, url: "https://drive.google.com/file/d/X", warnings: ["2026-09-22：工種超出容量 11 列，已略過 2 項"] });
    const r = e2.post(submit("2026-09-22", 3));
    return eq([r.ok, r.xlsx.ok, r.xlsx.url, r.xlsx.warnings], [true, true, "https://drive.google.com/file/d/X", ["2026-09-22：工種超出容量 11 列，已略過 2 項"]]);
  });
  T("P7", "真實容量警告：工種項目超過容量時，warnings 非空且透傳", () => {
    const e2 = makeEnv();
    const list = e2.state.sheetsData["工種機具材料清單"];
    for (let i = 0; i < 14; i++) list.push(["工種", "工種" + i, "", "TRUE", CASE_ID]);
    const p = submit("2026-09-22", 1);
    for (let i = 0; i < 14; i++) p.items.push({ name: "工種" + i, am: 1, pm: 1 });
    const r = e2.post(p);
    return eq([r.ok, r.xlsx.ok, r.xlsx.warnings.length > 0], [true, true, true]);
  });
  T("P8", "缺日期／日期格式錯誤：回錯誤，不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const a = e2.post({ date: "", caseId: CASE_ID, header: {} }), b = e2.post({ date: "9/21", caseId: CASE_ID, header: {} });
    return eq([a.ok, b.ok, called], [false, false, 0]);
  });
  T("P9", "密碼錯誤：被擋，不寫入也不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => (k === "SUBMIT_PASSWORD" ? "pw" : null) }) };
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const r = e2.post(Object.assign(submit("2026-09-22", 3), { password: "bad" }));
    return eq([r.ok, called, e2.state.sheetsData["日報頭"].length], [false, 0, 2]);
  });
  T("P10", "案場不存在：回錯誤，不寫入也不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const r = e2.post(Object.assign(submit("2026-09-22", 3), { caseId: "不存在的案場" }));
    return eq([r.ok, /案場不存在或已停用/.test(r.error), called, e2.state.sheetsData["日報頭"].length], [false, true, 0, 2]);
  });
  T("P11", "addItem 不觸發 xlsx 更新", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    // addItem 需要 appendRow；本測試假分頁不支援，只驗證進入 addItem 分支前不會走到 xlsx
    try { e2.post({ action: "addItem", caseId: CASE_ID, category: "無效", name: "x" }); } catch (e) {}
    return called === 0;
  });


  // ---------- 網頁「請款資料」卡片：addBilling ----------
  const bill = (extra) => Object.assign({ action: "addBilling", caseId: CASE_ID, date: "2026-09-20", zone: "土建", period: "第1期", amount: "1,500,000", progress: "35%", remark: "含保留款", invoice: "AB-1" }, extra || {});
  const billRows = (e) => e.state.sheetsData["請款資料"];
  const eb = makeEnv();
  T("P12", "新增請款：寫入「請款資料」第一個空白列（第2列，不是接在 J 欄說明之後）；金額去千分位、進度35%→0.35；xlsx 當天與之後分頁出現請款表", () => {
    eb.post(submit("2026-09-21", 5));
    const r = eb.post(bill());
    const row = billRows(eb)[1], wb = book(eb);
    const a = wb.Sheets["115.9.20"], b = wb.Sheets["115.9.21"];
    return eq([r.ok, r.xlsx.ok, row.slice(0, 8), billRows(eb).length, a.A47.v, a.E47.v, a.N47.v, a.P47.v, a.R47.v, b.C47.v], [true, true, [CASE_ID, "2026-09-20", "土建", "第1期", 1500000, 0.35, "含保留款", "AB-1"], 5, "土建", 1500000, 0.35, "含保留款", "AB-1", "第1期"]);
  });
  T("P13", "第二筆新增在第3列（往下找空白列）；同區域下一期的累計至上期＝第1期金額", () => {
    const r = eb.post(bill({ date: "2026-09-21", period: "第2期", amount: 700000, progress: 50 }));
    const b = book(eb).Sheets["115.9.21"];
    return eq([r.ok, billRows(eb)[2].slice(2, 6), b.C48.v, b.E48.v, b.I48.v, b.N48.v], [true, ["土建", "第2期", 700000, 0.5], "第2期", 700000, 1500000, 0.5]);
  });
  T("P14", "重複（同案場、請款日期、區域、期別）：拒絕、不新增", () => {
    const n = billRows(eb).length;
    const r = eb.post(bill());
    return eq([r.ok, /已存在/.test(r.error), billRows(eb).length], [false, true, n]);
  });
  T("P15", "欄位驗證：缺區域/期別/金額、金額負數或非數字、進度>100、日期格式錯、備註過長 → 都拒絕且不新增、不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const bad = [{ zone: "" }, { period: " " }, { amount: "" }, { amount: -1 }, { amount: "abc" }, { progress: 150 }, { progress: "x" }, { date: "2026/9/20" }, { date: "2026-02-30" }, { remark: "x".repeat(101) }];
    const res = bad.map((b) => e2.post(bill(b)).ok);
    return eq([res.every((x) => x === false), called, billRows(e2).length], [true, 0, 5]);
  });
  T("P16", "進度可空白（存成空字串）、金額 0 可以；備註/發票可空白；開頭是 = 的文字會被加單引號防公式注入", () => {
    const e2 = makeEnv();
    const r = e2.post(bill({ progress: "", amount: 0, remark: "", invoice: "", zone: "=SUM(A1)" }));
    const row = billRows(e2)[1];
    return eq([r.ok, row[4], row[5], row[6], row[7], row[2]], [true, 0, "", "", "", "'=SUM(A1)"]);
  });
  T("P17", "xlsx 更新丟例外：請款資料仍寫入且回 ok:true、xlsx:{ok:false,error}", () => {
    const e2 = makeEnv();
    e2.ctx.xlsxUpdateForDate = () => { throw new Error("Drive 掛了"); };
    const r = e2.post(bill());
    return eq([r.ok, r.xlsx.ok, /Drive 掛了/.test(r.xlsx.error), billRows(e2)[1][2]], [true, false, true, "土建"]);
  });
  T("P18", "該案場還沒有任何日報分頁可寫（請款日期早於所有日報且 xlsx 尚未建立）：請款資料照寫，xlsx 視為略過而不是失敗", () => {
    const e2 = makeEnv();
    const r = e2.post(bill({ date: "2026-09-01" }));
    return eq([r.ok, r.xlsx.ok, r.xlsx.skipped === true, billRows(e2).length], [true, true, true, 5]);
  });
  T("P19", "密碼錯誤／案場不存在：被擋、不新增；沒有「請款資料」分頁：回清楚錯誤", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => (k === "SUBMIT_PASSWORD" ? "pw" : null) }) };
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const a = e2.post(Object.assign(bill(), { password: "bad" }));
    const e3 = makeEnv(); e3.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const b = e3.post(bill({ caseId: "不存在的案場" }));
    const e4 = makeEnv(); delete e4.state.sheetsData["請款資料"];
    const c = e4.post(bill());
    return eq([a.ok, billRows(e2).length, b.ok, /案場不存在/.test(b.error), billRows(e3).length, c.ok, /請款資料/.test(c.error), called], [false, 5, false, true, 5, false, true, 0]);
  });
  T("P20", "新增案場(addCase)也寫在第一個空白列（案場設定 J 欄有說明文字時不會接在說明後面）", () => {
    const e2 = makeEnv();
    const cs = e2.state.sheetsData["案場設定"];
    cs[0][9] = "使用說明"; for (let i = 1; i <= 4; i++) { cs[i] = cs[i] || []; while (cs[i].length < 9) cs[i].push(""); cs[i][9] = "說明" + i; }
    const r = e2.post({ action: "addCase", code: "kh", name: "高雄案", owner: "甲", contract: "", startDate: "", company: "乙" });
    return eq([r.ok, cs[2].slice(0, 3), cs.length], [true, ["kh", "高雄案", "TRUE"], 5]);
  });


  // ---------- 一次新增多個請款項目（共用請款日期，全有或全無，只更新一次 xlsx）----------
  const items3 = [
    { zone: "土建", period: "第3期", amount: "1,000,000", progress: "10", remark: "", invoice: "A1" },
    { zone: "機電", period: "第3期", amount: 500000, progress: "", remark: "含保留款", invoice: "" },
    { zone: "景觀", period: "第3期", amount: "300,000.5", progress: "5%", remark: "", invoice: "" },
  ];
  const multi = (items, extra) => Object.assign({ action: "addBilling", caseId: CASE_ID, date: "2026-09-20", items }, extra || {});
  T("P21", "多項目：3 筆一次寫入（第2、3、4 列）、回 added=3、xlsx 只更新一次、請款表依序顯示、累計各區域獨立", () => {
    const e2 = makeEnv(); let called = 0; const real = e2.ctx.xlsxUpdateForDate;
    e2.post(submit("2026-09-21", 5));
    e2.ctx.xlsxUpdateForDate = function () { called++; return real.apply(this, arguments); };
    const r = e2.post(multi(items3));
    const rows = billRows(e2), a = book(e2).Sheets["115.9.20"];
    return eq([r.ok, r.added, r.xlsx.ok, called, rows[1].slice(2, 6), rows[2].slice(2, 6), rows[3].slice(2, 6), a.A47.v, a.A48.v, a.A49.v, a.E52.v, a.I52.v],
      [true, 3, true, 1, ["土建", "第3期", 1000000, 0.1], ["機電", "第3期", 500000, ""], ["景觀", "第3期", 300000.5, 0.05], "土建", "機電", "景觀", 1800000.5, 0]);
  });
  T("P22", "全有或全無：第2項金額不合法 → 整批拒絕、一筆都不新增、不呼叫 xlsx，錯誤訊息指出是哪一項", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const items = [items3[0], Object.assign({}, items3[1], { amount: "abc" }), items3[2]];
    const r = e2.post(multi(items));
    return eq([r.ok, /請款項目 2/.test(r.error), billRows(e2).length, called], [false, true, 5, 0]);
  });
  T("P23", "批次內重複（同區域＋期別）或與既有資料重複：整批拒絕，不新增", () => {
    const e2 = makeEnv();
    const dup = e2.post(multi([items3[0], items3[1], Object.assign({}, items3[0], { amount: 1 })]));
    e2.post(bill());                                            // 先寫入一筆 土建 第1期
    const n = billRows(e2).length;
    const exist = e2.post(multi([items3[1], { zone: "土建", period: "第1期", amount: 1 }]));
    return eq([dup.ok, /重複/.test(dup.error), exist.ok, /已存在/.test(exist.error), billRows(e2).length], [false, true, false, true, n]);
  });
  T("P24", "空項目清單、超過 20 筆：拒絕；單筆舊格式(欄位直接放在 data)仍可用", () => {
    const e2 = makeEnv();
    const none = e2.post(multi([]));
    const tooMany = e2.post(multi(Array.from({ length: 21 }, (_, i) => ({ zone: "z" + i, period: "p", amount: 1 }))));
    const legacy = e2.post(bill());
    return eq([none.ok, tooMany.ok, /最多/.test(tooMany.error), legacy.ok, legacy.added, billRows(e2)[1][2]], [false, false, true, true, 1, "土建"]);
  });


  T("P25", "請款日期與日報日期無關：請款日期晚於所有既有日報 → 既有分頁不變、xlsx 回 ok；之後送出該日期之後的日報，分頁才顯示這筆", () => {
    const e2 = makeEnv();
    e2.post(submit("2026-09-21", 5));
    const before = book(e2).Sheets["115.9.21"].A47;
    const r = e2.post(bill({ date: "2026-12-31", amount: 900000, progress: 70 }));
    const mid = book(e2);
    const later = e2.post(submit("2027-01-02", 3));
    const s2 = book(e2).Sheets["116.1.2"];
    return eq([before === undefined, r.ok, r.xlsx.ok, mid.Sheets["115.9.21"].A47 === undefined, later.xlsx.ok, s2.A47.v, s2.E47.v, s2.N47.v], [true, true, true, true, true, "土建", 900000, 0.7]);
  });

  R.forEach((r) => console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + r[3] : "")));
  console.log({ total: R.length, notPass: R.filter((r) => r[1] !== "PASS").length });
})();
