// 多案場隔離測試：node multicase.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
// 專門驗證「案場設定」上線後，兩個案場（lng／h2）的清單、當天資料、累計、xlsx 完全互不干擾，
// 以及不存在／已停用的案場代碼會被拒絕。其餘輸入驗證細節見 validation.test.js。
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
    "案場設定": [["案場代碼", "案場名稱", "啟用中", "業主", "合約金額（元）", "開工日期（YYYY/MM/DD）", "公司名稱"],
      ["lng", "洲際LNG", "TRUE", "中油", "1000", "2026/01/01", "坤鎔營造"],
      ["h2", "中鼎H2", "TRUE", "台肥", "2000", "2026/02/01", "坤鎔營造"],
      ["old", "已停用案場", "FALSE", "X", "", "", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中", "案場"],
      ["工種", "公司工", "", "TRUE", "lng"], ["工種", "公司工", "", "TRUE", "h2"]],
    "日報頭": [HDR],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間", "案場"]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間", "案場"]],
    "內部記錄": [["日期", "類型", "類別", "內容", "時間", "備註", "clientId", "更新時間", "案場"]],
  };
}

function makeEnv() {
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
        setValue(v) { while (data[n].length < r) data[n].push([]); data[n][r - 1][c - 1] = v; },
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

const D = "2026-09-21";
const sub = (caseId, o) => Object.assign({
  date: D, caseId, clientId: "cid",
  header: { weather: "晴", status: "施工", reporter: "王" },
  items: [{ name: "公司工", am: 1, pm: 1 }], attendance: [],
}, o || {});

(function main() {
  const e = makeEnv();

  T("M1", "cases：只列出啟用中的案場，停用的（old）不出現", () => {
    const r = e.get({ action: "cases" });
    return eq([r.ok, r.cases.map((c) => c.code).sort()], [true, ["h2", "lng"]]);
  });

  T("M2", "config：各案場只看到自己的「公司工」清單項目與基本資料，不會混到另一案", () => {
    const rl = e.get({ action: "config", caseId: "lng" });
    const rh = e.get({ action: "config", caseId: "h2" });
    return eq([rl.ok, rl.items.length, rl.basic["工程名稱"], rh.ok, rh.items.length, rh.basic["工程名稱"]], [true, 1, "洲際LNG", true, 1, "中鼎H2"]);
  });

  T("M3", "同一天、同一項目名稱，兩案場各自送出：day/cumulative 互不干擾", () => {
    e.post(sub("lng", { items: [{ name: "公司工", am: 3, pm: 3 }] }));
    e.post(sub("h2", { items: [{ name: "公司工", am: 9, pm: 9 }] }));
    const dl = e.get({ action: "day", caseId: "lng", date: D });
    const dh = e.get({ action: "day", caseId: "h2", date: D });
    const cl = e.get({ action: "cumulative", caseId: "lng" });
    const ch = e.get({ action: "cumulative", caseId: "h2" });
    return eq([dl.records[0].am, dh.records[0].am, cl.totals["公司工"], ch.totals["公司工"]], [3, 9, 3, 9]);
  });

  T("M4", "addItem：lng 新增的項目不會出現在 h2 的清單裡", () => {
    e.post({ action: "addItem", caseId: "lng", category: "材料", name: "水泥" });
    const rl = e.get({ action: "config", caseId: "lng" });
    const rh = e.get({ action: "config", caseId: "h2" });
    return eq([rl.items.some((i) => i.name === "水泥"), rh.items.some((i) => i.name === "水泥")], [true, false]);
  });

  T("M5", "案場代碼不存在／已停用：doGet/doPost 都拒絕", () => {
    const g1 = e.get({ action: "config", caseId: "old" });
    const g2 = e.get({ action: "config", caseId: "不存在" });
    const p1 = e.post(sub("old"));
    return eq([g1.ok, g2.ok, p1.ok], [false, false, false]);
  });

  T("M6", "xlsx：兩案場各自產生獨立檔案，內容與基本資料互不混淆", () => {
    const wbLng = XLSX.read(Object.values(e.state.files).find((f) => f.name === "施工日報彙整_洲際LNG.xlsx" && !f.trashed).bytes);
    const wbH2 = XLSX.read(Object.values(e.state.files).find((f) => f.name === "施工日報彙整_中鼎H2.xlsx" && !f.trashed).bytes);
    const sl = wbLng.Sheets["115.9.21"], sh = wbH2.Sheets["115.9.21"];
    return eq([sl.A3.v, sh.A3.v, sl.D8.v, sh.D8.v], ["業主：中油", "業主：台肥", 3, 9]);
  });

  T("M7", "adminDeleteDayCore：刪 lng 當天資料，h2 同一天資料不受影響", () => {
    const jr = (v) => JSON.parse(JSON.stringify(v));
    const r = jr(e.call(`adminDeleteDayCore(${JSON.stringify(D)},'lng',{name:'洲際LNG',owner:'中油',contract:'1000',startDate:'2026/01/01',company:'坤鎔營造'})`));
    const dl = e.get({ action: "day", caseId: "lng", date: D });
    const dh = e.get({ action: "day", caseId: "h2", date: D });
    return eq([r.ok, r.deleted.header, dl.header, dh.header && dh.header.date], [true, 1, null, D]);
  });

  T("M8", "addCase：建立新案場並沿用 lng 的清單，新案場立刻出現在 cases 清單、清單項目已複製且彼此獨立", () => {
    const r = e.post({ action: "addCase", code: "new1", name: "測試新案場", owner: "業主X", contract: "500", startDate: "2027/01/01", company: "公司X", copyFrom: "lng" });
    const cases = e.get({ action: "cases" }).cases.map((c) => c.code).sort();
    const cfgNew = e.get({ action: "config", caseId: "new1" });
    const cfgLng = e.get({ action: "config", caseId: "lng" });
    return eq(
      [r.ok, r.code, r.copied, cases, cfgNew.ok, cfgNew.items.map((i) => i.name), cfgNew.basic["工程名稱"], cfgNew.basic["業主"], cfgLng.items.length],
      [true, "new1", 2, ["h2", "lng", "new1"], true, ["公司工", "水泥"], "測試新案場", "業主X", 2]
    );
  });

  T("M9", "addCase：案場代碼重複被拒，不寫入", () => {
    const r = e.post({ action: "addCase", code: "lng", name: "重複代碼" });
    const cases = e.get({ action: "cases" }).cases.length;
    return eq([r.ok, /已存在/.test(r.error), cases], [false, true, 3]);
  });

  T("M10", "addCase：代碼含非法字元 / 名稱含檔名危險字元 / 缺代碼或名稱 都被拒", () => {
    const bad1 = e.post({ action: "addCase", code: "有 空格", name: "x" });
    const bad2 = e.post({ action: "addCase", code: "ok2", name: 'a/b' });
    const bad3 = e.post({ action: "addCase", code: "", name: "x" });
    const bad4 = e.post({ action: "addCase", code: "ok3", name: "" });
    const cases = e.get({ action: "cases" }).cases.length;
    return eq([bad1.ok, bad2.ok, bad3.ok, bad4.ok, cases], [false, false, false, false, 3]);
  });

  T("M11", "addCase：不複製（copyFrom 空）→ 新案場清單為空；不影響其他案場", () => {
    const r = e.post({ action: "addCase", code: "empty1", name: "空清單案場" });
    const cfg = e.get({ action: "config", caseId: "empty1" });
    return eq([r.ok, r.copied, cfg.items.length], [true, 0, 0]);
  });

  R.forEach((r) => console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + r[3] : "")));
  console.log({ total: R.length, notPass: R.filter((r) => r[1] !== "PASS").length });
})();
