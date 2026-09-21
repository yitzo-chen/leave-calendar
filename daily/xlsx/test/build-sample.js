// 以範本產生範例 xlsx：node build-sample.js <輸出路徑>
// 需要 jszip（NODE_PATH 指向含 jszip 的資料夾）。
// 範例是「資料驅動」：先給每天的紀錄，再用 computeCumulatives 算出逐日累計（當天＋之前），
// 同時輸出 <輸出路徑>.expected.json 供 verify.py 逐格核對。
const fs = require("fs");
const JSZip = require("jszip");
const XB = require("../xlsx-builder.js");
const TEMPLATE = "D:/洲際液化天然氣接收站AI/日報範本.xlsx";

async function loadTemplate() {
  const zip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  const parts = {};
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir || name.endsWith(".bin")) continue;
    parts[name] = await zip.file(name).async("string");
  }
  return parts;
}
async function writeXlsx(parts, path) {
  const zip = new JSZip();
  const order = ["[Content_Types].xml", "_rels/.rels"];
  Object.keys(parts)
    .sort((a, b) => (order.indexOf(a) + 1 || 99) - (order.indexOf(b) + 1 || 99) || a.localeCompare(b))
    .forEach((n) => zip.file(n, parts[n], { createFolders: false }));
  fs.writeFileSync(path, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// 清單（啟用中項目，依清單順序）
const LONG_MATERIAL = "樓梯防滑地磚(20*27)含黏著劑及施工損耗合計(箱)";
const CATALOG = [
  ["工種", "公司工"], ["工種", "模板工"], ["工種", "鋼筋工(工地)"], ["工種", "搭架工"],
  ["機具", "吊車"], ["機具", "壓送車"],
  ["材料", "210kg/cm²混凝土(m³)"], ["材料", "砂(m³)"], ["材料", LONG_MATERIAL],
];
const categoryByName = {};
CATALOG.forEach(([c, n]) => { categoryByName[n] = c; });

// 每天的紀錄（只列有數字的項目）
const RECORDS = [
  { date: "2026-09-18", name: "公司工", am: 3, pm: 2 }, { date: "2026-09-18", name: "模板工", am: 2, pm: 2 },
  { date: "2026-09-18", name: "吊車", am: 1, pm: 1 }, { date: "2026-09-18", name: "210kg/cm²混凝土(m³)", am: 3.25, pm: 0 },
  { date: "2026-09-18", name: LONG_MATERIAL, am: 2, pm: 0 },
  { date: "2026-09-19", name: "公司工", am: 2, pm: 2 }, { date: "2026-09-19", name: "鋼筋工(工地)", am: 1, pm: 1 },
  { date: "2026-09-19", name: "210kg/cm²混凝土(m³)", am: 2.5, pm: 0 }, { date: "2026-09-19", name: "砂(m³)", am: 4, pm: 0 },
  { date: "2026-09-20", name: "公司工", am: 4, pm: 4 }, { date: "2026-09-20", name: "模板工", am: 1, pm: 1 },
  { date: "2026-09-20", name: "吊車", am: 2, pm: 2 }, { date: "2026-09-20", name: LONG_MATERIAL, am: 1, pm: 0 },
];
const ATTENDANCE = {
  "2026-09-18": [
    { name: "陳一", am: true, pm: true, amHours: 0, pmHours: 0, reason: "" },
    { name: "林二", am: true, pm: false, amHours: 1, pmHours: 0, reason: "材料進場" },
    { name: "張三", am: false, pm: true, amHours: 0, pmHours: 2, reason: "趕澆置" },
  ],
  "2026-09-19": [{ name: "陳一", am: true, pm: true, amHours: 0, pmHours: 0, reason: "" }],
  "2026-09-20": [
    { name: "陳一", am: true, pm: true, amHours: 0, pmHours: 0, reason: "" },
    { name: "林二", am: true, pm: true, amHours: 0, pmHours: 3, reason: "配合泵送車進場時間安排夜間趕工" },
  ],
};
const HEADER = {
  weather: "雨", status: "施工",
  todayWork: "1.一樓牆面粉刷\n2.二樓鋼筋綁紮，因應混凝土澆置進度趕工，需配合泵送車進場時間安排人力調度",
  tomorrowPlan: "1.三樓模板組立", remark: "第一行備註\n第二行備註", reporter: "陳易佐",
  directorAm: "王大明", directorPm: "王大明", safetyAm: "李小華", safetyPm: "李小華",
};
const BASIC = { "公司名稱": "", "業主": "中油", "工程名稱": "洲際液化天然氣接收站", "合約金額（元）": "", "開工日期（YYYY/MM/DD）": "" };

(async () => {
  const out = process.argv[2] || "sample.xlsx";
  const prepared = XB.prepareTemplate(await loadTemplate());
  const templateSheet = prepared["xl/worksheets/sheet1.xml"];
  const cumByDate = XB.computeCumulatives(RECORDS, categoryByName);

  const sheets = [], expected = {}, warningsAll = [];
  // 故意亂序：先產生 20、再 18、再 19
  for (const date of ["2026-09-20", "2026-09-18", "2026-09-19"]) {
    const cum = XB.cumulativeAsOf(cumByDate, date);
    const today = {};
    RECORDS.filter((r) => r.date === date).forEach((r) => { today[r.name] = r; });
    const items = CATALOG.map(([category, name]) => ({
      category, name, am: (today[name] || {}).am || 0, pm: (today[name] || {}).pm || 0, cumulative: cum[name] || 0,
    }));
    const r = XB.buildDaySheet(templateSheet, { date, basic: BASIC, header: HEADER, items, attendance: ATTENDANCE[date] || [] });
    const name = XB.sheetNameForDate(date);
    sheets.push({ name, xml: r.xml });
    warningsAll.push([date, r.warnings]);
    // 預期的累計欄（工種第8列起、機具第20列起、材料第31列起，依清單順序）
    const exp = {};
    const rows = { "工種": 8, "機具": 20, "材料": 31 };
    const seen = { "工種": 0, "機具": 0, "材料": 0 };
    items.forEach((it) => {
      const row = rows[it.category] + seen[it.category]++;
      exp["A" + row] = it.name;
      exp["F" + row] = Math.round((cum[it.name] || 0) * 100) / 100;
    });
    expected[name] = exp;
  }
  await writeXlsx(XB.assembleWorkbook(prepared, sheets), out);
  fs.writeFileSync(out + ".expected.json", JSON.stringify(expected, null, 1));
  console.log("已產生", out, "警告:", JSON.stringify(warningsAll));
  console.log("逐日累計（公司工）:", ["2026-09-18", "2026-09-19", "2026-09-20"].map((d) => d + "=" + XB.cumulativeAsOf(cumByDate, d)["公司工"]).join("  "));
})().catch((e) => { console.error(e); process.exit(1); });
