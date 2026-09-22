// doPost/doGet 輸入驗證與防汙染測試：node validation.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
// 涵蓋：未知 action、數量/時數驗證、文字型別與長度上限、全有或全無、同名人員重複、
//       項目名稱當 key 的汙染（constructor/__proto__）、累計略過不在清單的項目、addItem 重啟換類別、GET 未初始化、
//       同請求項目重名、= 開頭名稱的存/讀/比對一致（unsanitizeCell）、第 5 次密碼錯誤即告知鎖定。
const fs = require("fs");
const XLSX = require("xlsx");
const XB = require("../xlsx-builder.js");
const { makeGlueEnv } = require("./fake-gas.js");
const TEMPLATE_BYTES = fs.readFileSync("D:/洲際液化天然氣接收站AI/日報範本.xlsx");

const R = [];
function T(id, name, fn) {
  try { const r = fn(); const pass = r === true || r === undefined; R.push([id, pass ? "PASS" : "FAIL", name, pass ? "" : (typeof r === "string" ? r : JSON.stringify(r))]); }
  catch (e) { R.push([id, "ERROR", name, String(e && e.stack || e).split("\n").slice(0, 3).join(" | ")]); }
}
const eq = (a, b) => (JSON.stringify(a) === JSON.stringify(b) ? true : "got " + JSON.stringify(a) + " expected " + JSON.stringify(b));

const HDR = ["日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目", "備註", "填表人", "clientId", "更新時間", "主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)"];
function baseSheets() {
  return {
    "基本資料": [["業主", "中油"], ["工程名稱", "測試工程"], ["合約金額（元）", ""], ["開工日期（YYYY/MM/DD）", ""], ["公司名稱", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中"], ["工種", "公司工", "", "TRUE"], ["工種", "模板工", "", "TRUE"], ["材料", "砂(m³)", "", "TRUE"]],
    "日報頭": [HDR],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間"]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間"]],
    "內部記錄": [["日期", "類型", "類別", "內容", "時間", "備註", "clientId", "更新時間"]],
  };
}

// 假環境：在 fake-gas 的 Drive 上補可讀寫的試算表（含 setValue / appendRow / deleteRow）、鎖、快取、ContentService
function makeEnv(sheetsData) {
  const env = makeGlueEnv({ sheetsData: sheetsData || baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  const data = env.state.sheetsData, cache = {};
  const maxRows = env.maxRows = {}; // 各分頁目前的列數上限（預設 1000），超出時 setValues 會丟例外（同真實 Sheets）
  const sheet = (n) => ({
    getMaxRows: () => (n in maxRows ? maxRows[n] : 1000),
    insertRowsAfter: (after, cnt) => { maxRows[n] = (n in maxRows ? maxRows[n] : 1000) + cnt; },
    getLastRow: () => { let l = 0; data[n].forEach((r, i) => { if (r.some((v) => v !== "" && v != null)) l = i + 1; }); return l; },
    getDataRange: () => ({ getValues: () => JSON.parse(JSON.stringify(data[n])) }),
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = data[n][r - 1 + i] || []; const x = []; for (let j = 0; j < nc; j++) x.push(row[c - 1 + j] === undefined ? "" : row[c - 1 + j]); o.push(x); } return o; },
        setValues(vs) { if (r + nr - 1 > (n in maxRows ? maxRows[n] : 1000)) throw new Error("範圍超出分頁大小"); for (let i = 0; i < nr; i++) { while (data[n].length < r + i) data[n].push([]); for (let j = 0; j < nc; j++) data[n][r - 1 + i][c - 1 + j] = vs[i][j]; } },
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
  env.rows = (n) => data[n].length - 1; // 不含表頭的資料列數
  return env;
}
const D = "2026-09-21";
const sub = (o) => Object.assign({
  date: D, clientId: "cid",
  header: { weather: "晴", status: "施工", todayWork: "w", tomorrowPlan: "t", remark: "r", reporter: "王", directorAm: "甲", directorPm: "甲", safetyAm: "乙", safetyPm: "乙" },
  items: [{ name: "公司工", am: 1, pm: 1 }], attendance: [{ name: "林二", am: true, pm: true, amHours: 1, pmHours: 0, reason: "趕工" }],
}, o || {});
// 拒絕且完全沒寫入任何資料
const rejected = (env, r, re) => (r.ok === false && (!re || re.test(r.error)) && env.rows("日報頭") === 0 && env.rows("日報記錄") === 0 && env.rows("本工出勤") === 0 && env.rows("內部記錄") === 0) || JSON.stringify({ r, h: env.rows("日報頭"), rec: env.rows("日報記錄"), att: env.rows("本工出勤"), internal: env.rows("內部記錄") });

(function main() {
  T("V1", "未知 action（deleteAll）被拒，不寫入；submit / addItem 仍正常", () => {
    const e = makeEnv();
    const bad = e.post(sub({ action: "deleteAll" }));
    const ok1 = e.post(sub({ action: "submit" })), ok2 = e.post(sub({ date: "2026-09-22" }));
    return eq([bad.ok, !!bad.error, ok1.ok, ok2.ok, e.rows("日報頭")], [false, true, true, true, 2]);
  });
  T("V2", "數量非數字 / 負數 / 無限大 / 布林 / 物件 → 整個請求被拒且不寫入任何資料", () => {
    const bads = ["abc", -1, "-0.5", "Infinity", true, { a: 1 }, [5]];
    for (const v of bads) {
      const e = makeEnv();
      const r1 = rejected(e, e.post(sub({ items: [{ name: "公司工", am: v, pm: 0 }] })), /上午數量/);
      if (r1 !== true) return "am=" + JSON.stringify(v) + " " + r1;
      const r2 = rejected(e, e.post(sub({ items: [{ name: "公司工", am: 1, pm: v }] })), /下午數量/);
      if (r2 !== true) return "pm=" + JSON.stringify(v) + " " + r2;
    }
    return true;
  });
  T("V3", "出勤時數 NaN / 負 → 拒絕且不寫入", () => {
    const e = makeEnv();
    const a = rejected(e, e.post(sub({ attendance: [{ name: "甲", amHours: "x" }] })), /上午加班時數/);
    const b = rejected(e, e.post(sub({ attendance: [{ name: "甲", pmHours: -3 }] })), /下午加班時數/);
    return a === true ? b : a;
  });
  T("V4", "空值語意不變：null/\"\"/未填視為 0（項目全 0 不寫入；出勤時數存 0）；字串數字 \"2.5\" 照收", () => {
    const e = makeEnv();
    const r = e.post(sub({ items: [{ name: "公司工", am: null, pm: "" }, { name: "模板工", pm: "2.5" }], attendance: [{ name: "甲", am: true, amHours: null, pmHours: "" }] }));
    const s = e.state.sheetsData;
    return eq([r.ok, e.rows("日報記錄"), s["日報記錄"][1].slice(1, 4), s["本工出勤"][1].slice(4, 6)], [true, 1, ["模板工", 0, 2.5], [0, 0]]);
  });
  T("V5", "文字欄位收到物件 / 陣列 → 拒絕且不寫入", () => {
    const cases = [{ remark: { a: 1 } }, { weather: ["x"] }, { todayWork: { t: 1 } }, { reporter: [] || 0, directorAm: {} }];
    for (const h of cases) {
      const e = makeEnv();
      const r = rejected(e, e.post(sub({ header: Object.assign({}, sub().header, h) })), /格式錯誤/);
      if (r !== true) return JSON.stringify(h) + " " + r;
    }
    const e2 = makeEnv();
    const r2 = rejected(e2, e2.post(sub({ items: [{ name: { x: 1 }, am: 1 }] })), /項目名稱/);
    const e3 = makeEnv();
    const r3 = rejected(e3, e3.post(sub({ attendance: [{ name: "甲", reason: ["a"] }] })), /加班原因/);
    return r2 === true ? r3 : r2;
  });
  T("V6", "items / attendance 含 null（或非物件）元素 → 整個拒絕；日報頭也沒寫入（全有或全無）", () => {
    const e = makeEnv();
    const a = rejected(e, e.post(sub({ items: [{ name: "公司工", am: 1, pm: 1 }, null] })), /項目資料格式錯誤/);
    const e2 = makeEnv();
    const b = rejected(e2, e2.post(sub({ attendance: [null] })), /出勤資料格式錯誤/);
    const e3 = makeEnv();
    const c = rejected(e3, e3.post(sub({ items: ["x"], attendance: [[]] })));
    return a === true ? (b === true ? c : b) : a;
  });
  T("V7", "後段（出勤）才不合法時，前面的日報頭 / 項目也不會寫入", () => {
    const e = makeEnv();
    return rejected(e, e.post(sub({ items: [{ name: "公司工", am: 5, pm: 5 }], attendance: [{ name: "甲" }, { name: "乙", amHours: -1 }] })));
  });
  T("V8", "長度上限：備註/本日/明日 2000 字（剛好可過、+1 被拒）；名稱/姓名/原因/填表人 100 字；錯誤訊息指出欄位", () => {
    const L = (n) => "字".repeat(n);
    const e = makeEnv();
    const ok = e.post(sub({ header: Object.assign({}, sub().header, { remark: L(2000), todayWork: L(2000), tomorrowPlan: L(2000), reporter: L(100) }), items: [{ name: L(100), am: 1 }], attendance: [{ name: L(100), reason: L(100) }] }));
    const tests = [
      [{ header: Object.assign({}, sub().header, { remark: L(2001) }) }, /備註/],
      [{ header: Object.assign({}, sub().header, { todayWork: L(2001) }) }, /本日施工項目/],
      [{ header: Object.assign({}, sub().header, { tomorrowPlan: L(2001) }) }, /預計明日施工項目/],
      [{ header: Object.assign({}, sub().header, { reporter: L(101) }) }, /填表人/],
      [{ items: [{ name: L(101), am: 1 }] }, /項目名稱/],
      [{ attendance: [{ name: L(101) }] }, /人員姓名/],
      [{ attendance: [{ name: "甲", reason: L(101) }] }, /加班原因/],
    ];
    for (const [p, re] of tests) {
      const e2 = makeEnv();
      const r = rejected(e2, e2.post(sub(p)), re);
      if (r !== true) return re + " " + r;
    }
    return ok.ok === true || JSON.stringify(ok);
  });
  T("V9", "同名人員重複 → 拒絕並回報姓名（trim 後比對，含全形空白）；不同名可過", () => {
    const e = makeEnv();
    const a = rejected(e, e.post(sub({ attendance: [{ name: "甲", am: true }, { name: " 甲　", pm: true }, { name: "乙" }, { name: "乙" }] })), /重複：甲、乙/);
    const e2 = makeEnv();
    const ok = e2.post(sub({ attendance: [{ name: "甲" }, { name: "乙" }, { name: "" }, { name: "  " }, { name: "" }] }));
    return a === true ? eq([ok.ok, e2.rows("本工出勤")], [true, 2]) : a;
  });
  T("V10", "姓名 trim（含全形空白）後存入，後續同名 upsert 不新增列", () => {
    const e = makeEnv();
    e.post(sub({ attendance: [{ name: " 甲　", am: true, amHours: 1 }] }));
    e.post(sub({ attendance: [{ name: "甲", am: true, amHours: 2 }] }));
    const s = e.state.sheetsData["本工出勤"];
    return eq([e.rows("本工出勤"), s[1][1], s[1][4]], [1, "甲", 2]);
  });
  T("V11", "不得破壞 clearItems / removeAttendance：明確清空的項目與人員照樣刪除，沒列出的不動", () => {
    const e = makeEnv();
    e.post(sub({ items: [{ name: "公司工", am: 5, pm: 5 }, { name: "模板工", am: 2, pm: 2 }], attendance: [{ name: "甲", am: true }, { name: "乙", am: true }] }));
    const r = e.post(sub({ items: [{ name: "公司工", am: 0, pm: 0 }], attendance: [{ name: "甲", am: true }], clearItems: ["公司工"], removeAttendance: ["乙"] }));
    const g = e.get({ action: "day", date: D });
    return eq([r.ok, g.records.map((x) => x.name), g.attendance.map((x) => x.name)], [true, ["模板工"], ["甲"]]);
  });
  T("V12", "clearItems / removeAttendance 內含物件 → 拒絕且不寫入", () => {
    const e = makeEnv();
    const a = rejected(e, e.post(sub({ clearItems: [{ x: 1 }] })), /清除的項目名稱/);
    const e2 = makeEnv();
    const b = rejected(e2, e2.post(sub({ removeAttendance: [["x"]] })), /移除的人員姓名/);
    return a === true ? b : a;
  });
  T("V13", "addItem：停用項目用不同類別重啟 → 類別以新傳入的為準；同類別重啟不變", () => {
    const e = makeEnv();
    const list = e.state.sheetsData["工種機具材料清單"];
    list[1][3] = "FALSE"; list[2][3] = "FALSE";
    const r1 = e.post({ action: "addItem", category: "材料", name: "公司工" });
    const r2 = e.post({ action: "addItem", category: "工種", name: "模板工" });
    return eq([r1.ok, list[1].slice(0, 4), r2.ok, list[2].slice(0, 4), list.length], [true, ["材料", "公司工", "", "TRUE"], true, ["工種", "模板工", "", "TRUE"], 4]);
  });
  T("V14", "addItem：名稱過長 / 物件被拒，不新增", () => {
    const e = makeEnv();
    const a = e.post({ action: "addItem", category: "工種", name: "x".repeat(101) });
    const b = e.post({ action: "addItem", category: "工種", name: { a: 1 } });
    const c = e.post({ action: "addItem", category: "工種", name: "x".repeat(100) });
    return eq([a.ok, /項目名稱/.test(a.error), b.ok, c.ok, e.state.sheetsData["工種機具材料清單"].length], [false, true, false, true, 5]);
  });
  T("V15", "GET / POST 於未初始化（缺分頁）的試算表：回友善錯誤、不丟例外、不寫入", () => {
    const e = makeEnv({}); // 沒有任何分頁
    const out = [];
    for (const a of ["config", "day", "cumulative", "reporters", "peopleNames"]) {
      const r = e.get({ action: a, date: D });
      out.push([r.ok, r.error]);
    }
    const p = e.post(sub()), q = e.post({ action: "addItem", category: "工種", name: "x" });
    const idle = e.get({});
    const want = [false, "尚未初始化，請先執行 setupSheets"];
    return eq([out, p.ok, p.error, q.ok, q.error, idle.ok], [out.map(() => want), false, want[1], false, want[1], true]);
  });
  T("V16", "GET cumulative：不在清單的項目略過並回警告；constructor / __proto__ 名稱（在清單內）不汙染", () => {
    const sd = baseSheets();
    sd["工種機具材料清單"].push(["工種", "constructor", "", "TRUE"], ["工種", "__proto__", "", "TRUE"]);
    sd["日報記錄"].push([D, "constructor", 2, 2, "c", "t"], [D, "__proto__", 4, 4, "c", "t"], [D, "神秘", 3, 3, "c", "t"], [D, "砂(m³)", 3, 0, "c", "t"]);
    const e = makeEnv(sd);
    const c = e.get({ action: "cumulative" });
    return eq([c.ok, c.totals.constructor, Object.prototype.hasOwnProperty.call(c.totals, "__proto__") && c.totals["__proto__"], "神秘" in c.totals, c.totals["砂(m³)"], c.warnings], [true, 2, 4, false, 3, ["累計略過不在清單的項目：神秘"]]);
  });
  T("V17", "GET reporters / peopleNames：姓名 constructor / toString 不被去重誤判", () => {
    const sd = baseSheets();
    sd["日報頭"].push([D, "晴", "施工", "", "", "", "constructor", "c", "t"]);
    sd["本工出勤"].push([D, "toString", "V", "", 0, 0, "", "c", "t"]);
    const e = makeEnv(sd);
    return eq([e.get({ action: "reporters" }).names, e.get({ action: "peopleNames" }).names], [["constructor"], ["toString"]]);
  });
  T("V18", "XlsxBuilder.computeCumulatives：不在清單者略過並寫入 warnings（同一項目只警告一次）；不傳 warnings 也不拋例外", () => {
    const recs = [{ date: "2026-09-18", name: "公司工", am: 2, pm: 2 }, { date: "2026-09-18", name: "神秘", am: 9, pm: 9 }, { date: "2026-09-19", name: "神秘", am: 1, pm: 1 }, { date: "2026-09-19", name: "公司工", am: 2, pm: 2 }];
    const cat = { "公司工": "工種" }, w = [];
    const by = XB.computeCumulatives(recs, cat, w);
    const noW = XB.computeCumulatives(recs, cat);
    return eq([w, by["2026-09-19"]["公司工"], "神秘" in by["2026-09-19"], noW["2026-09-18"]["公司工"]], [["累計略過不在清單的項目：神秘"], 4, false, 2]);
  });
  T("V19", "XlsxBuilder.computeCumulatives：項目名 constructor/toString/__proto__ 不汙染（categoryByName 為一般 {} 也可）", () => {
    const recs = [{ date: "2026-09-18", name: "constructor", am: 2, pm: 2 }, { date: "2026-09-18", name: "toString", am: 1, pm: 1 }, { date: "2026-09-18", name: "__proto__", am: 2, pm: 2 }, { date: "2026-09-18", name: "hasOwnProperty", am: 5, pm: 5 }];
    const cat = Object.create(null);
    ["constructor", "toString", "__proto__"].forEach((n) => { cat[n] = "工種"; });
    const w = [];
    const by = XB.computeCumulatives(recs, cat, w);
    const snap = XB.cumulativeAsOf(by, "2026-09-18");
    const w2 = [], plain = XB.computeCumulatives(recs, {}, w2); // 一般 {} 清單：繼承來的 constructor/toString 等不算「在清單內」，全部略過
    return eq([snap.constructor, snap.toString, snap["__proto__"], w, Object.keys(plain), w2.length], [2, 1, 2, ["累計略過不在清單的項目：hasOwnProperty"], [], 4]);
  });
  T("V20", "端到端：清單內項目 constructor 寫進 xlsx 數字正確；記錄中不在清單的項目 → xlsx.warnings 含略過訊息，日報照送出", () => {
    const sd = baseSheets();
    sd["工種機具材料清單"].push(["工種", "constructor", "", "TRUE"]);
    sd["日報頭"].push(["2026-09-20", "雨", "施工", "x", "", "", "王", "c", "t", "甲", "甲", "乙", "乙"]);
    sd["日報記錄"].push(["2026-09-20", "神秘", 3, 3, "c", "t"]);
    const e = makeEnv(sd);
    const r = e.post(sub({ items: [{ name: "constructor", am: 2, pm: 2 }, { name: "公司工", am: 1, pm: 1 }] }));
    const wb = XLSX.read(Object.values(e.state.files).find((f) => f.name === "施工日報彙整.xlsx" && !f.trashed).bytes);
    const s = wb.Sheets["115.9.21"];
    // 清單順序：公司工(A8)、模板工(A9)、constructor(A10)
    return eq([r.ok, r.xlsx.ok, r.xlsx.warnings.some((w) => /累計略過不在清單的項目：神秘/.test(w)), s.A8.v, s.A10.v, s.D10.v, s.E10.v, s.F10.v], [true, true, true, "公司工", "constructor", 2, 2, 2]);
  });

  T("V21", "同一請求 items trim 後同名（兩筆都有非 0 數字）→ 拒絕並回報項目名，不寫入；其中一筆全 0 則略過不算重複", () => {
    const e = makeEnv();
    const a = rejected(e, e.post(sub({ items: [{ name: "公司工", am: 1 }, { name: " 公司工　", pm: 2 }, { name: "模板工", am: 1 }, { name: "模板工", pm: 1 }] })), /項目重複：公司工、模板工/);
    const e2 = makeEnv();
    const ok = e2.post(sub({ items: [{ name: "公司工", am: 1, pm: 1 }, { name: "公司工", am: 0, pm: 0 }, { name: "模板工", am: 2 }] }));
    return a === true ? eq([ok.ok, e2.rows("日報記錄"), e2.state.sheetsData["日報記錄"][1].slice(1, 4)], [true, 2, ["公司工", 1, 1]]) : a;
  });
  T("V22", "= + - @ 開頭姓名/項目名：儲存加單引號、讀出還原、再送同名不重複、不寫入公式（B33b/B38/B104）", () => {
    const e = makeEnv();
    const p = () => sub({ items: [{ name: "=X", am: 1, pm: 1 }, { name: "-Y", am: 1 }], attendance: [{ name: "=X", am: true, reason: "@r" }, { name: "+Z", am: true }] });
    e.post(p()); const r2 = e.post(p());
    const s = e.state.sheetsData;
    const a1 = e.post({ action: "addItem", category: "工種", name: "=N" }), a2 = e.post({ action: "addItem", category: "工種", name: "=N" });
    const g = e.get({ action: "day", date: D }), pn = e.get({ action: "peopleNames" }).names, cfg = e.get({ action: "config" }).items.map((i) => i.name);
    // 儲存格內容仍是加前綴的字面文字（不會被當公式）；同一天重送不新增列
    return eq([r2.ok, e.rows("日報記錄"), e.rows("本工出勤"), s["日報記錄"][1][1], s["本工出勤"][1][1], s["本工出勤"][1][6],
      a1.ok, a2.ok, a2.error, s["工種機具材料清單"].filter((r) => r[1] === "'=N").length,
      g.records.map((x) => x.name), g.attendance.map((x) => x.name), g.attendance[0].reason, pn, cfg.indexOf("=N") >= 0],
      [true, 2, 2, "'=X", "'=X", "'@r", true, false, "項目已存在", 1, ["=X", "-Y"], ["=X", "+Z"], "@r", ["=X", "+Z"], true]);
  });
  T("V22b", "儲存格被 Sheets 吃掉單引號（存成 =X）時同樣能比對、讀出", () => {
    const sd = baseSheets();
    sd["本工出勤"].push([D, "=X", "V", "", 0, 0, "", "c", "t"]);
    const e = makeEnv(sd);
    const r = e.post(sub({ items: [], attendance: [{ name: "=X", am: true, amHours: 3 }] }));
    return eq([r.ok, e.rows("本工出勤"), e.get({ action: "peopleNames" }).names], [true, 1, ["=X"]]);
  });
  T("V22c", "端到端：= 開頭姓名寫進 xlsx 時去掉前綴（顯示為使用者原輸入）", () => {
    const e = makeEnv();
    const r = e.post(sub({ attendance: [{ name: "=X", am: true, pm: true }] }));
    const wb = XLSX.read(Object.values(e.state.files).find((f) => f.name === "施工日報彙整.xlsx" && !f.trashed).bytes);
    const vals = [];
    Object.values(wb.Sheets["115.9.21"]).forEach((c) => { if (c && typeof c.v === "string") vals.push(c.v); });
    return eq([r.ok, r.xlsx.ok, vals.indexOf("=X") >= 0, vals.indexOf("'=X") >= 0, r.xlsx.error], [true, true, true, false, undefined]);
  });
  T("V23", "密碼：第 5 次錯誤（觸發鎖定）就回「已鎖定」，第 1~4 次仍是密碼錯誤；鎖定後正確密碼也被拒", () => {
    const e = makeEnv();
    e.state.props.SUBMIT_PASSWORD = "pw";
    const msgs = [];
    for (let i = 0; i < 5; i++) msgs.push(e.post(sub({ password: "bad" })).error);
    const after = e.post(sub({ password: "pw" }));
    return eq([msgs.slice(0, 4).every((m) => m === "通關密碼錯誤，請向管理者確認"), /已暫時鎖定/.test(msgs[4]), after.ok, /已暫時鎖定/.test(after.error)], [true, true, false, true]);
  });

  T("V24", "分頁列數用完時自動擴充（真實 Sheets 對超出列數的 setValues 會丟例外），三個資料分頁都不失敗", () => {
    const e = makeEnv();
    e.maxRows["日報頭"] = 1; e.maxRows["日報記錄"] = 1; e.maxRows["本工出勤"] = 1;
    const r = e.post(sub({ items: [{ name: "公司工", am: 1, pm: 1 }, { name: "模板工", am: 2, pm: 2 }, { name: "砂(m³)", am: 3, pm: 3 }],
      attendance: [{ name: "林二", am: true }, { name: "王三", pm: true }] }));
    return eq([r.ok, e.rows("日報頭"), e.rows("日報記錄"), e.rows("本工出勤"), e.maxRows["日報記錄"] >= 4], [true, 1, 3, 2, true]);
  });

  T("V25", "xlsxUrl：檔案尚未產生回空字串；送出日報後回檔案網址（與送出結果的 url 相同）；輸出檔被丟垃圾桶後又回空字串", () => {
    const e = makeEnv();
    const before = e.get({ action: "xlsxUrl" });
    const r = e.post(sub());
    const after = e.get({ action: "xlsxUrl" });
    Object.values(e.state.files).forEach((f) => { if (f.name === "施工日報彙整.xlsx") f.trashed = true; });
    const trashed = e.get({ action: "xlsxUrl" });
    return eq([before, r.ok && r.xlsx.ok, after.ok && /^https:\/\/drive\.google\.com\//.test(after.url), after.url === r.xlsx.url, trashed],
      [{ ok: true, url: "" }, true, true, true, { ok: true, url: "" }]);
  });

  // ---- 後台管理：一鍵刪除某天資料（選單「日報管理」）----
  const D1 = "2026-09-21", D2 = "2026-09-22";
  const jr = (v) => JSON.parse(JSON.stringify(v)); // 跨 vm 領域的物件轉成一般物件
  const core = (e, d) => jr(e.call("adminDeleteDayCore(" + JSON.stringify(d) + ")"));
  const xlsxBytes = (e) => Object.values(e.state.files).find((f) => f.name === "施工日報彙整.xlsx" && !f.trashed);
  const twoDays = () => { // 兩天資料：D1 公司工 1/1，D2 公司工 2/2（D2 累計 = 1 + 2 = 3）
    const e = makeEnv();
    e.post(sub({ date: D1, items: [{ name: "公司工", am: 1, pm: 1 }] }));
    e.post(sub({ date: D2, items: [{ name: "公司工", am: 2, pm: 2 }] }));
    return e;
  };

  T("V26", "parseAdminDate：接受西元/民國與 - / . 分隔；不合法（不存在的日期、2 位數年、亂字）回空字串", () => {
    const e = makeEnv();
    const p = (s) => e.call("parseAdminDate(" + JSON.stringify(s) + ")");
    return eq(["2026-09-21", "2026/9/1", "115.9.21", "115/09/21", " 2026-9-5 "].map(p).concat(["2026-02-30", "26-9-21", "abc", "", "2026-13-01"].map(p)),
      ["2026-09-21", "2026-09-01", "2026-09-21", "2026-09-21", "2026-09-05", "", "", "", "", ""]);
  });
  T("V27", "刪除某天：三張表該天的列都刪、別天不動；xlsx 移除該日分頁並重算之後日期的累計（3 → 2）", () => {
    const e = twoDays();
    const before = XLSX.read(xlsxBytes(e).bytes);
    const cumBefore = before.Sheets["115.9.22"].F8.v;
    const cnt = jr(e.call("adminCountDay(" + JSON.stringify(D1) + ")"));
    const r = core(e, D1);
    const after = XLSX.read(xlsxBytes(e).bytes);
    return eq([cumBefore, cnt, r.ok, r.deleted, r.xlsx.ok, e.rows("日報頭"), e.rows("日報記錄"), e.rows("本工出勤"), after.SheetNames, after.Sheets["115.9.22"].F8.v],
      [3, { header: 1, record: 1, attendance: 1, xlsx: true }, true, { header: 1, record: 1, attendance: 1 }, true, 1, 1, 1, ["115.9.22"], 2]);
  });
  T("V28", "刪除最後一天：xlsx 已無分頁 → 檔案丟垃圾桶並清掉記錄的檔案 ID；之後再送出日報會重新建立、xlsxUrl 指到新檔", () => {
    const e = twoDays();
    core(e, D1);
    const r = core(e, D2);
    const gone = xlsxBytes(e) === undefined, propGone = !("XLSX_FILE_ID" in e.state.props);
    e.post(sub({ date: D2 }));
    const url = e.get({ action: "xlsxUrl" }).url;
    return eq([r.ok, r.xlsx.removedFile, gone, propGone, e.rows("日報頭") === 1 && e.rows("日報記錄") === 1, /^https:\/\/drive\.google\.com\//.test(url)], [true, true, true, true, true, true]);
  });
  T("V29", "刪除沒有資料的日期：xlsx 尚未建立 → 略過；xlsx 已存在 → 不動其他分頁；日報頭已被手動刪掉的殘留分頁與資料也一併清除", () => {
    const fresh = makeEnv();
    const r0 = core(fresh, D1);
    const e = twoDays();
    const r1 = core(e, "2030-01-01");
    const sheetsKept = XLSX.read(xlsxBytes(e).bytes).SheetNames;
    // 模擬手動刪掉 D1 的日報頭：資料表殘留該天的記錄／出勤，xlsx 仍有分頁
    const e2 = twoDays();
    const hdr = e2.state.sheetsData["日報頭"]; hdr.splice(1, 1);
    const cnt = jr(e2.call("adminCountDay(" + JSON.stringify(D1) + ")"));
    const r2 = core(e2, D1);
    return eq([r0.ok, r0.deleted, r0.xlsx.skipped, r1.ok, r1.deleted, sheetsKept, cnt, r2.deleted, XLSX.read(xlsxBytes(e2).bytes).SheetNames, e2.rows("日報記錄"), e2.rows("本工出勤")],
      [true, { header: 0, record: 0, attendance: 0 }, true, true, { header: 0, record: 0, attendance: 0 }, ["115.9.21", "115.9.22"],
        { header: 0, record: 1, attendance: 1, xlsx: true }, { header: 0, record: 1, attendance: 1 }, ["115.9.22"], 1, 1]);
  });
  T("V30", "有人正在送出（鎖被占用）→ 回「系統忙碌」，不刪任何資料；日期格式錯誤也直接拒絕", () => {
    const e = twoDays();
    e.ctx.LockService = { getScriptLock: () => ({ waitLock() { throw new Error("busy"); }, releaseLock() {} }) };
    const busy = core(e, D1), bad = core(e, "115.9.21");
    return eq([busy.ok, /忙碌/.test(busy.error), e.rows("日報頭"), e.rows("日報記錄"), bad.ok], [false, true, 2, 2, false]);
  });

  R.forEach((r) => console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + String(r[3]).slice(0, 500) : "")));
  console.log({ total: R.length, notPass: R.filter((r) => r[1] !== "PASS").length });
})();
