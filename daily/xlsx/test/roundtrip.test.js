// 進階測試：node roundtrip.test.js（需 jszip、xlsx(SheetJS) 於 NODE_PATH）
const fs = require("fs");
const JSZip = require("jszip");
const XLSX = require("xlsx");
const XB = require("../xlsx-builder.js");
const TEMPLATE = "D:/洲際液化天然氣接收站AI/日報範本.xlsx";

const R = [];
function T(id, name, fn) {
  try {
    const r = fn();
    const pass = r === true || r === undefined;
    R.push([id, pass ? "PASS" : "FAIL", name, pass ? "" : (typeof r === "string" ? r : JSON.stringify(r))]);
  } catch (e) {
    R.push([id, "ERROR", name, String(e.stack || e).split("\n").slice(0, 3).join(" | ")]);
  }
}
const eq = (a, b) => (JSON.stringify(a) === JSON.stringify(b) ? true : "got " + JSON.stringify(a) + " expected " + JSON.stringify(b));

async function loadParts(buf) {
  const z = await JSZip.loadAsync(buf);
  const p = {};
  for (const n of Object.keys(z.files)) {
    if (z.files[n].dir || n.endsWith(".bin")) continue;
    p[n] = await z.file(n).async("string");
  }
  return p;
}
async function toBuf(parts) {
  const z = new JSZip();
  Object.keys(parts).forEach((n) => z.file(n, parts[n], { createFolders: false }));
  return z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

(async () => {
  const base = XB.prepareTemplate(await loadParts(fs.readFileSync(TEMPLATE)));
  const tsheet = base["xl/worksheets/sheet1.xml"];
  const O = {};
  const mk = (date, extra) => Object.assign({ date, basic: {}, header: { weather: "晴", status: "施工", reporter: "王" }, items: [], attendance: [] }, extra || {});
  const build = (date, extra) => {
    const r = XB.buildDaySheet(tsheet, mk(date, extra), O);
    return { name: XB.sheetNameForDate(date), xml: r.xml, w: r.warnings };
  };
  const strip = (x) => x.replace(/ tabSelected="1"/g, "");

  const d1 = build("2026-09-18", { items: [{ category: "工種", name: "公司工", am: 1, pm: 1, cumulative: 1 }] });
  const d2 = build("2026-09-19", { items: [{ category: "工種", name: "公司工", am: 2, pm: 2, cumulative: 3 }] });
  const d3 = build("2026-09-21", { items: [{ category: "工種", name: "公司工", am: 3, pm: 3, cumulative: 6 }] });
  const wb1 = XB.assembleWorkbook(base, [d1, d2, d3]);

  T("R01", "重複組裝穩定：組裝→讀出分頁→再組裝，所有零件逐字相同", () => {
    const wb2 = XB.assembleWorkbook(base, XB.readSheets(wb1), "115.9.21");
    const diff = Object.keys(wb1).filter((k) => wb1[k] !== wb2[k]);
    return diff.length === 0 || diff;
  });
  T("R02", "覆蓋單日：只改 115.9.19，其他分頁 XML 逐字不變", () => {
    const cur = XB.readSheets(wb1);
    const nd = build("2026-09-19", { items: [{ category: "工種", name: "公司工", am: 9, pm: 9, cumulative: 9 }] });
    const next = cur.map((s) => (s.name === nd.name ? { name: nd.name, xml: nd.xml } : s));
    const wb2 = XB.assembleWorkbook(base, next, "115.9.21");
    const a = XB.readSheets(wb1), b = XB.readSheets(wb2);
    const same = a.filter((s, i) => s.name !== "115.9.19" && strip(s.xml) === strip(b[i].xml)).length;
    return eq([a.length, b.length, same, b[1].xml.includes("<v>9</v>")], [3, 3, 2, true]);
  });
  T("R03", "中間插入較早日期：分頁重新排序、Print_Area 的 localSheetId 全部重算", () => {
    const wb2 = XB.assembleWorkbook(base, XB.readSheets(wb1).concat([build("2026-09-17", {})]), "115.9.21");
    const ids = [...wb2["xl/workbook.xml"].matchAll(/localSheetId="(\d+)">'([^']+)'/g)].map((m) => [Number(m[1]), m[2]]);
    return eq(ids, [[0, "115.9.17"], [1, "115.9.18"], [2, "115.9.19"], [3, "115.9.21"]]);
  });
  T("R04", "只有 1 個分頁被選取，且 activeTab 指向最新一天", () => {
    const sel = Object.keys(wb1).filter((k) => /worksheets\/sheet\d+\.xml$/.test(k) && wb1[k].includes('tabSelected="1"'));
    return eq([sel.length, /activeTab="2"/.test(wb1["xl/workbook.xml"])], [1, true]);
  });
  T("R04b", "切換 active 分頁後，舊的選取標記被清掉（不會出現 2 個）", () => {
    const wb2 = XB.assembleWorkbook(base, XB.readSheets(wb1), "115.9.18");
    const sel = Object.keys(wb2).filter((k) => /worksheets\/sheet\d+\.xml$/.test(k) && wb2[k].includes('tabSelected="1"'));
    return eq([sel.length, /activeTab="0"/.test(wb2["xl/workbook.xml"])], [1, true]);
  });

  const many = (cat, n, hasData) => Array.from({ length: n }, (_, i) => ({ category: cat, name: cat + i, am: hasData(i) ? 1 : 0, pm: 0, cumulative: i }));
  T("R05", "工種 15 項（僅第 0、1、14 項有數字）：有數字的優先放入，空位依序補 2~9 項、第 10 項起不入，警告一次", () => {
    const r = XB.buildDaySheet(tsheet, mk("2026-09-20", { items: many("工種", 15, (i) => i === 0 || i === 1 || i === 14) }), O);
    const has = (n) => r.xml.includes(">" + n + "<");
    return eq([r.warnings.length, has("工種14"), has("工種0"), has("工種9"), has("工種10")], [1, true, true, true, false]);
  });
  T("R06", "材料剛好 15 項：不警告", () => eq(XB.buildDaySheet(tsheet, mk("2026-09-20", { items: many("材料", 15, () => true) }), O).warnings, []));
  T("R07", "工種 13 項且全部有數字：多出者不入，警告『有數字的項目…超過』", () => {
    const r = XB.buildDaySheet(tsheet, mk("2026-09-20", { items: many("工種", 13, () => true) }), O);
    return (r.warnings.length === 1 && /有數字的項目共 13 項/.test(r.warnings[0])) || r.warnings;
  });
  const P = (n, extra) => Array.from({ length: n }, (_, i) => Object.assign({ name: "人" + i, am: true, pm: false, amHours: 0, pmHours: 0, reason: "" }, extra ? extra(i) : {}));
  T("R08", "上午 9 人（第 8、9 人有加班）：加班者優先、其餘併入第 6 列並警告", () => {
    const r = XB.buildDaySheet(tsheet, mk("2026-09-20", { attendance: P(9, (i) => (i >= 7 ? { amHours: 1, reason: "趕" } : {})) }), O);
    return eq([r.warnings.length, r.xml.includes(">人7<"), r.xml.includes(">人8<"), /<t xml:space="preserve">人3、人4、人5、人6</.test(r.xml)], [1, true, true, true]);
  });
  T("R09", "本日施工項目 30 行：超過 13 行併入最後一行並警告", () => {
    const text = Array.from({ length: 30 }, (_, i) => "第" + (i + 1) + "行").join("\n");
    const r = XB.buildDaySheet(tsheet, mk("2026-09-20", { header: { todayWork: text } }), O);
    return eq([r.warnings.length, r.xml.includes("第13行 第14行")], [1, true]);
  });

  const ctrl = String.fromCharCode(1) + String.fromCharCode(8) + "ctrl";
  const nasty = ["<script>alert(1)</script>", 'a&b"c\'d', "=1+1", "  前後空白  ", "😀 emoji 中文", "line1\nline2", ctrl, '</t></is></c><c r="Z1"><v>1</v></c>'];
  const nastySheet = XB.buildDaySheet(tsheet, mk("2026-09-20", {
    basic: { "業主": nasty[0], "工程名稱": nasty[1] },
    header: { weather: "晴", status: "施工", remark: nasty[5], reporter: nasty[7], directorAm: nasty[2], safetyAm: nasty[3] },
    attendance: [{ name: nasty[4], am: true, pm: false, amHours: 1, pmHours: 0, reason: nasty[6] }],
    items: [{ category: "工種", name: nasty[0], am: 1, pm: 0, cumulative: 1 }],
  }), O).xml;
  const wbN = XB.assembleWorkbook(base, [{ name: "115.9.20", xml: nastySheet }]);
  const wsN = XLSX.read(await toBuf(wbN)).Sheets["115.9.20"];
  T("R10", "特殊字元原樣寫入且不破壞結構（SheetJS 讀回）", () => eq(
    [wsN.A3.v, wsN.A4.v, wsN.J34.v, wsN.J37.v, wsN.M34.v, wsN.P34.v, wsN.O53.v, wsN.A8.v],
    ["業主：" + nasty[0], "工程名稱：" + nasty[1], "=1+1", "  前後空白  ", nasty[4], "ctrl", nasty[7], nasty[0]]));
  T("R11", "注入字串沒有製造出額外儲存格（Z1 不存在）", () => wsN.Z1 === undefined || "Z1 被注入");
  T("R12", "備註多行原樣保留換行", () => eq(wsN.A46.v, "備註：" + nasty[5]));

  const buf = await toBuf(wb1);
  const x = XLSX.read(buf);
  T("R13", "SheetJS 讀得到 3 個分頁且順序正確", () => eq(x.SheetNames, ["115.9.18", "115.9.19", "115.9.21"]));
  T("R14", "SheetJS 讀回數值：115.9.19 公司工上午2/下午2/累計3、日期「日期：115」/9/19", () => {
    const s = x.Sheets["115.9.19"];
    return eq([s.A8.v, s.D8.v, s.E8.v, s.F8.v, s.J5.v, s.N5.v, s.P5.v], ["公司工", 2, 2, 3, "日期：115", 9, 19]);
  });
  T("R15", "SheetJS 合併數 174、含備註區 A46:S52 與日期 J5:L5", () => {
    const m = x.Sheets["115.9.19"]["!merges"].map((r) => XLSX.utils.encode_range(r));
    return eq([m.length, m.includes("A46:S52"), m.includes("J5:L5")], [174, true, true]);
  });
  T("R16", "檔案大小合理（3 天 < 40KB）", () => buf.length < 40000 || buf.length);

  const t0 = Date.now();
  const big = [];
  for (let i = 0; i < 365; i++) {
    const d = new Date(Date.UTC(2026, 0, 1 + i)).toISOString().slice(0, 10);
    big.push(build(d, { items: [{ category: "工種", name: "公司工", am: 1, pm: 1, cumulative: i }] }));
  }
  const bb = await toBuf(XB.assembleWorkbook(base, big));
  const ms = Date.now() - t0;
  T("R17", "365 個分頁可組裝並能被 SheetJS 讀回", () => {
    const xb = XLSX.read(bb);
    return eq([xb.SheetNames.length, xb.SheetNames[0], xb.SheetNames[364]], [365, "115.1.1", "115.12.31"]);
  });
  console.log("365 天檔案：" + (bb.length / 1024).toFixed(0) + "KB，含壓縮共耗時 " + ms + "ms");

  let f = 0;
  R.forEach((r) => {
    console.log("[" + r[1] + "] " + r[0] + " " + r[2] + (r[3] ? "\n      -> " + r[3] : ""));
    if (r[1] !== "PASS") f++;
  });
  console.log({ total: R.length, notPass: f });
})();
