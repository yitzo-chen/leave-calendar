// Apps Script 串接程式測試：node glue.test.js（需 xlsx(SheetJS) 於 NODE_PATH）
// 多案場改版後，xlsx-drive.gs 的入口都多了 caseId/caseInfo 參數；這裡固定用一個假案場 "lng" 測試，
// 資料列（日報頭/日報記錄/本工出勤/工種機具材料清單）也都補上對應的「案場」欄位值。
const fs = require("fs");
const XLSX = require("xlsx");
const { makeGlueEnv } = require("./fake-gas.js");
const TEMPLATE_BYTES = fs.readFileSync("D:/洲際液化天然氣接收站AI/日報範本.xlsx");

const R = [];
function T(id, name, fn) {
  try {
    const r = fn();
    const pass = r === true || r === undefined;
    R.push([id, pass ? "PASS" : "FAIL", name, pass ? "" : (typeof r === "string" ? r : JSON.stringify(r))]);
  } catch (e) {
    R.push([id, "ERROR", name, String(e && e.message || e)]);
  }
}
const eq = (a, b) => (JSON.stringify(a) === JSON.stringify(b) ? true : "got " + JSON.stringify(a) + " expected " + JSON.stringify(b));

const CASE_ID = "lng";
const CASE_LITERAL = "{name:'測試工程',owner:'中油',contract:'',startDate:'',company:''}";
const OUTPUT_NAME = "施工日報彙整_測試工程.xlsx";
const call2 = (fn, date) => `${fn}('${date}','${CASE_ID}',${CASE_LITERAL})`;

const HDR = (extra) => Object.assign(["日期", "天氣", "施工狀況", "本日施工項目", "預計明日施工項目", "備註", "填表人", "clientId", "更新時間", "主任(上午)", "主任(下午)", "工安(上午)", "工安(下午)"]);
function baseSheets() {
  return {
    "基本資料": [["業主", "中油"], ["工程名稱", "測試工程"], ["合約金額（元）", ""], ["開工日期（YYYY/MM/DD）", ""], ["公司名稱", ""]],
    "案場設定": [["案場代碼", "案場名稱", "啟用中", "業主", "合約金額（元）", "開工日期（YYYY/MM/DD）", "公司名稱"], [CASE_ID, "測試工程", "TRUE", "中油", "", "", ""]],
    "工種機具材料清單": [["類別", "項目名稱", "工項編號", "啟用中", "案場"],
      ["工種", "公司工", "", "TRUE", CASE_ID], ["工種", "模板工", "", "TRUE", CASE_ID], ["工種", "舊工種", "", "FALSE", CASE_ID], ["材料", "砂(m³)", "", "TRUE", CASE_ID]],
    "日報頭": [HDR(), ["2026-09-18", "晴", "施工", "第一天", "", "", "王", "c", "t", "甲", "甲", "乙", "乙", CASE_ID], ["2026-09-20", "雨", "施工", "第三天", "", "", "王", "c", "t", "甲", "甲", "乙", "乙", CASE_ID]],
    "日報記錄": [["日期", "項目名稱", "上午", "下午", "clientId", "更新時間", "案場"],
      ["2026-09-18", "公司工", 3, 2, "c", "t", CASE_ID], ["2026-09-18", "舊工種", 1, 1, "c", "t", CASE_ID], ["2026-09-18", "砂(m³)", 5, 0, "c", "t", CASE_ID], ["2026-09-20", "公司工", 4, 4, "c", "t", CASE_ID]],
    "本工出勤": [["日期", "人員名稱", "上午", "下午", "上午加班", "下午加班", "加班原因", "clientId", "更新時間", "案場"],
      ["2026-09-18", "陳一", "V", "V", 0, 0, "", "c", "t", CASE_ID], ["2026-09-20", "林二", "V", "", 1, 0, "材料進場", "c", "t", CASE_ID]],
  };
}
const sheetVals = (bytes) => XLSX.read(bytes);
const F = (env) => Object.values(env.state.files).find((f) => f.name === OUTPUT_NAME && !f.trashed);

(async () => {
  T("G1", "自我測試：可建立檔案並讀回 2 個分頁", () => {
    const env = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
    const r = env.call("xlsxSelfTest()");
    return eq([r.ok, r.sheets], [true, ["115.9.18", "115.9.19"]]);
  });

  const env = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  let r1;
  T("G2", "首次更新 2026-09-20：建立該案場檔案；只有該日分頁（其他日期未指定不重建）", () => {
    r1 = env.call(call2("xlsxUpdateForDate", "2026-09-20"));
    const wb = sheetVals(F(env).bytes);
    return eq([r1.ok, r1.sheets, r1.full, wb.SheetNames, env.state.patches], [true, 1, false, ["115.9.20"], 0]);
  });
  T("G3", "分頁內容：公司工 4/4、累計＝(3+2)/2+(4+4)/2＝6.5、人員林二 加班1 材料進場", () => {
    const s = sheetVals(F(env).bytes).Sheets["115.9.20"];
    return eq([s.A8.v, s.D8.v, s.E8.v, s.F8.v, s.M34.v, s.O34.v, s.P34.v, s.A3.v, s.J34.v], ["公司工", 4, 4, 6.5, "林二", 1, "材料進場", "業主：中油", "甲"]);
  });
  T("G4", "停用項目：當天無紀錄不列出；有紀錄才列出（9/18 有『舊工種』）", () => {
    const s20 = sheetVals(F(env).bytes).Sheets["115.9.20"];
    const names20 = ["A8", "A9", "A10"].map((k) => s20[k] && s20[k].v);
    return eq(names20, ["公司工", "模板工", undefined]);
  });
  let fileId;
  T("G5", "更新較早的 9/18：同一檔案（用 REST 覆蓋，ID 不變）、兩個分頁、9/20 累計不變", () => {
    fileId = F(env).id;
    const r = env.call(call2("xlsxUpdateForDate", "2026-09-18"));
    const wb = sheetVals(F(env).bytes);
    return eq([r.sheets, r.rebuilt, F(env).id === fileId, env.state.patches, wb.SheetNames, wb.Sheets["115.9.18"].A10.v], [2, 2, true, 1, ["115.9.18", "115.9.20"], "舊工種"]);
  });
  T("G6", "9/18 資料修改後再更新：9/20（較晚）的累計一併重算（3+2 → 10+10）", () => {
    env.state.sheetsData["日報記錄"][1] = ["2026-09-18", "公司工", 10, 10, "c", "t", CASE_ID];
    const r = env.call(call2("xlsxUpdateForDate", "2026-09-18"));
    const wb = sheetVals(F(env).bytes);
    return eq([r.rebuilt, wb.Sheets["115.9.18"].F8.v, wb.Sheets["115.9.20"].F8.v], [2, 10, 14]);
  });
  T("G7", "未來日期沒有日報頭：不會憑空產生分頁，已存在的分頁保持", () => {
    const r = env.call(call2("xlsxUpdateForDate", "2026-09-25"));
    return eq([r.rebuilt, sheetVals(F(env).bytes).SheetNames], [0, ["115.9.18", "115.9.20"]]);
  });
  T("G8", "範本樣式改版（styles 不同）：自動全部重建", () => {
    // 模擬舊檔：把已存檔內的 styles 換成不同內容
    const f = F(env); const entries = env.readZip(f.bytes).map((e) => (e.name === "xl/styles.xml" ? { name: e.name, data: Buffer.from(e.data.toString("utf8") + " ") } : e));
    f.bytes = env.writeZip(entries);
    const r = env.call(call2("xlsxUpdateForDate", "2026-09-18"));
    return eq([r.full, r.sheets, r.rebuilt], [true, 2, 2]);
  });
  T("G9", "xlsxRebuildAll(caseId)：依日報頭全部日期重建（案場設定的基本資料一併帶入）", () => { const r = env.call(`xlsxRebuildAll('${CASE_ID}')`); return eq([r.full, r.sheets], [true, 2]); });
  T("G10", "輸出檔被移到垃圾桶：改依名稱找不到→建立新檔並更新記錄的 ID（該案場專屬的 property key）", () => {
    F(env).trashed = true;
    const r = env.call(call2("xlsxUpdateForDate", "2026-09-20"));
    return eq([r.ok, F(env).id !== fileId, env.state.props["XLSX_FILE_ID_" + CASE_ID] === F(env).id], [true, true, true]);
  });
  T("G11", "找不到範本檔：丟出清楚的錯誤訊息", () => {
    const e2 = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: {} });
    try { e2.call(call2("xlsxUpdateForDate", "2026-09-20")); return "應該要丟錯"; } catch (e) { return /找不到範本檔/.test(String(e.message)) || String(e.message); }
  });
  T("G12", "範本結構不對：丟出錯誤，且不動既有輸出檔", () => {
    const junk = env.writeZip([{ name: "a.txt", data: Buffer.from("x") }]);
    const e2 = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": junk } });
    try { e2.call(call2("xlsxUpdateForDate", "2026-09-20")); return "應該要丟錯"; } catch (e) { return /範本檔結構不符/.test(String(e.message)) || String(e.message); }
  });
  T("G13", "REST 更新失敗（HTTP 404）：丟出錯誤", () => {
    const e2 = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
    e2.call(call2("xlsxUpdateForDate", "2026-09-20"));
    e2.state.files[e2.state.props["XLSX_FILE_ID_" + CASE_ID]] && (e2.ctx.UrlFetchApp.fetch = () => ({ getResponseCode: () => 500, getContentText: () => "boom" }));
    try { e2.call(call2("xlsxUpdateForDate", "2026-09-20")); return "應該要丟錯"; } catch (e) { return /HTTP 500/.test(String(e.message)) || String(e.message); }
  });
  T("G14", "產出的 xlsx 可被 SheetJS 完整讀回（分頁數與合併數 221）", () => {
    const e2 = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
    e2.call(`xlsxRebuildAll('${CASE_ID}')`);
    const wb = sheetVals(Object.values(e2.state.files).find((f) => f.name === OUTPUT_NAME).bytes);
    return eq([wb.SheetNames, wb.Sheets["115.9.18"]["!merges"].length], [["115.9.18", "115.9.20"], 221]);
  });
  // ---------- 請款資料與備註累加（2026-10-07）----------
  const billingSheets = () => {
    const sh = baseSheets();
    sh["日報頭"][1][5] = "測量人員:甲、乙"; // 9/18 備註
    sh["日報頭"][2][5] = ["業主臨時巡場", "鋼板樁進場"].join(String.fromCharCode(10)); // 9/20 備註（兩行）
    sh["請款資料"] = [["案場", "請款日期", "區域", "期別", "本期請款金額", "目前進度", "備註", "發票"],
      [CASE_ID, "2026-09-18", "土建", "第1期", 1000000, 0.2, "含保留款", "AB-1"],
      [CASE_ID, new Date(Date.UTC(2026, 8, 19)), "土建", "第2期", "2,500,000", "35%", "", new Date(Date.UTC(2026, 9, 3))], // 日期物件、千分位字串、%字串、發票日期
      ["other", "2026-09-18", "別案", "第1期", 999, 0.1, "", ""],       // 別的案場：不能出現
      [CASE_ID, "2026-09-25", "土建", "第3期", 700000, 0.5, "", ""],     // 晚於 9/20：9/20 的分頁不能出現
      [CASE_ID, "abc", "土建", "壞日期", 1, 0, "", ""],                  // 日期無法辨識：略過並警告
      [CASE_ID, "115.9.18", "機電", "第1期", 300000, "", "", ""]];       // 民國日期
    return sh;
  };
  const eb = makeGlueEnv({ sheetsData: billingSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  let rb;
  T("G15", "請款資料：重建後 9/18 只有當天的 2 列（土建第1期、機電第1期[民國日期]），別案場與之後日期的列不出現；壞日期警告 1 則", () => {
    rb = eb.call(`xlsxRebuildAll('${CASE_ID}')`);
    const s = sheetVals(F(eb).bytes).Sheets["115.9.18"];
    return eq([s.A47.v, s.C47.v, s.E47.v, s.I47.v, s.N47.v, s.P47.v, s.R47.v, s.A48.v, s.C48.v, s.E48.v, s.N48 === undefined, s.A49 === undefined, s.E52.v,
      rb.warnings.filter((w) => /請款資料第 6 列/.test(w)).length],
      ["土建", "第1期", 1000000, 0, 0.2, "含保留款", "AB-1", "機電", "第1期", 300000, true, true, 1300000, 1]);
  });
  T("G16", "請款資料 9/20：依請款日期排序＝土建第1期、機電第1期、土建第2期（日期物件、\"2,500,000\"、\"35%\"→0.35）；第2期累計至上期＝第1期100萬；9/25 那列不出現", () => {
    const s = sheetVals(F(eb).bytes).Sheets["115.9.20"];
    const dump = Object.keys(s).filter((k) => /^[A-Z]+(4[6-9]|5[0-2])$/.test(k)).map((k) => k + "=" + s[k].v).join(" ");
    if (!s.A47 || !s.A48 || !s.A49) return "請款表區：" + dump;
    return eq([s.A47.v, s.C47.v, s.A48.v, s.A49.v, s.C49.v, s.E49.v, s.I49.v, s.N49.v, s.A50 === undefined, s.E52.v, s.I52.v],
      ["土建", "第1期", "機電", "土建", "第2期", 2500000, 1000000, 0.35, true, 3800000, 1000000]);
  });
  T("G17", "發票欄收到日期物件：轉成「月/日」文字 10/3", () => eq(sheetVals(F(eb).bytes).Sheets["115.9.20"].R49 && sheetVals(F(eb).bytes).Sheets["115.9.20"].R49.v, "10/3"));
  T("G18", "備註累加：9/18 分頁 U8＝9/18 備註；9/20 分頁依序列出 9/18、9/20（兩行拆兩列），都在 U 欄", () => {
    const wb = sheetVals(F(eb).bytes);
    const a = wb.Sheets["115.9.18"], b = wb.Sheets["115.9.20"];
    return eq([a.U7.v, a.U8.v, a.U9 === undefined, b.U8.v, b.U9.v, b.U10.v, b.U11 === undefined], ["註：", "115.9.18 測量人員:甲、乙", true, "115.9.18 測量人員:甲、乙", "115.9.20 業主臨時巡場", "115.9.20 鋼板樁進場", true]);
  });
  T("G19", "沒有「請款資料」分頁（尚未 setupSheets）：照常重建，請款表是空的、沒有警告", () => {
    const e2 = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
    const r = e2.call(`xlsxRebuildAll('${CASE_ID}')`);
    const s = sheetVals(Object.values(e2.state.files).find((f) => f.name === OUTPUT_NAME).bytes).Sheets["115.9.18"];
    return eq([r.ok, r.warnings.length, s.A47 === undefined, s.E52.v], [true, 0, true, 0]);
  });
  // 輸出檔留一份供 python 驗證
  const last = makeGlueEnv({ sheetsData: baseSheets(), driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
  last.call(`xlsxRebuildAll('${CASE_ID}')`);
  const out = Object.values(last.state.files).find((f) => f.name === OUTPUT_NAME).bytes;
  fs.writeFileSync(process.argv[2] || "glue-out.xlsx", out);

  let f = 0;
  R.forEach((r) => { console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + r[3] : "")); if (r[1] !== "PASS") f++; });
  console.log({ total: R.length, notPass: f });
})();
