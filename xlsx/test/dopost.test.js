// doPost 接 xlsx 的測試：node dopost.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
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

const HDR = ["日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目", "備註", "填表人", "clientId", "更新時間", "主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)"];
function baseSheets() {
  return {
    "基本資料": [["業主", "中油"], ["工程名稱", "測試工程"], ["合約金額（元）", ""], ["開工日期（YYYY/MM/DD）", ""], ["公司名稱", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中"], ["工種", "公司工", "", "TRUE"], ["工種", "模板工", "", "TRUE"], ["材料", "砂(m³)", "", "TRUE"]],
    "日報頭": [HDR, ["2026-09-20", "雨", "施工", "第三天", "", "", "王", "c", "t", "甲", "甲", "乙", "乙"]],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間"], ["2026-09-20", "公司工", 4, 4, "c", "t"]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間"]],
  };
}

// 在假 Drive 環境上補一個可寫入的試算表、鎖、快取與 ContentService，讓真正的 doPost 能跑
function makeEnv(opts) {
  opts = opts || {};
  const env = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  const data = env.state.sheetsData, cache = {};
  const sheet = (n) => ({
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
  });
  env.ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getId: () => "SS1", getSheetByName: (n) => (data[n] ? sheet(n) : null) }) };
  env.ctx.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() { env.state.released = (env.state.released || 0) + 1; } }) };
  env.ctx.CacheService = { getScriptCache: () => ({ get: (k) => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
  env.ctx.ContentService = { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: "JSON" } };
  env.post = (payload) => JSON.parse(env.call("doPost")({ postData: { contents: JSON.stringify(payload) } }).s);
  return env;
}
const F = (env) => Object.values(env.state.files).find((f) => f.name === "施工日報彙整.xlsx" && !f.trashed);
const book = (env) => XLSX.read(F(env).bytes);
const submit = (date, cong) => ({
  date, clientId: "cid", header: { weather: "晴", status: "施工", todayWork: "新的一天", reporter: "王", directorAm: "甲", directorPm: "甲", safetyAm: "乙", safetyPm: "乙" },
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
    for (let i = 0; i < 14; i++) list.push(["工種", "工種" + i, "", "TRUE"]);
    const p = submit("2026-09-22", 1);
    for (let i = 0; i < 14; i++) p.items.push({ name: "工種" + i, am: 1, pm: 1 });
    const r = e2.post(p);
    return eq([r.ok, r.xlsx.ok, r.xlsx.warnings.length > 0], [true, true, true]);
  });
  T("P8", "缺日期／日期格式錯誤：回錯誤，不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const a = e2.post({ date: "", header: {} }), b = e2.post({ date: "9/21", header: {} });
    return eq([a.ok, b.ok, called], [false, false, 0]);
  });
  T("P9", "密碼錯誤：被擋，不寫入也不呼叫 xlsx", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => (k === "SUBMIT_PASSWORD" ? "pw" : null) }) };
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    const r = e2.post(Object.assign(submit("2026-09-22", 3), { password: "bad" }));
    return eq([r.ok, called, e2.state.sheetsData["日報頭"].length], [false, 0, 2]);
  });
  T("P10", "addItem 不觸發 xlsx 更新", () => {
    const e2 = makeEnv(); let called = 0;
    e2.ctx.xlsxUpdateForDate = () => { called++; return { url: "u" }; };
    // addItem 需要 appendRow；本測試假分頁不支援，只驗證進入 addItem 分支前不會走到 xlsx
    try { e2.post({ action: "addItem", category: "無效", name: "x" }); } catch (e) {}
    return called === 0;
  });

  R.forEach((r) => console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + r[3] : "")));
  console.log({ total: R.length, notPass: R.filter((r) => r[1] !== "PASS").length });
})();
