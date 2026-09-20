/**
 * 日報 xlsx 產生器（純字串處理，不依賴任何 zip / XML 函式庫）。
 *
 * 目的：以「洲際 LNG 日報格式.xlsx」為範本，直接改寫 sheet XML，
 * 每天產生一個分頁（分頁名＝民國年.月.日），並組成一個工作簿。
 * 直接改 XML 才能完整保留範本的合併儲存格、框線、A4 列印設定。
 *
 * 寫成 ES5 風格，Node（測試）與 Google Apps Script（正式）共用同一份。
 * 欄位對應請看 cellmap.json。
 */
var XlsxBuilder = (function () {
  // ---------- 常數（來自 cellmap.json，範本原狀） ----------
  var CAP = { labor: 11, equipment: 10, material: 15, today: 13, tomorrow: 11, crew: 6 };
  var ROWS = {
    labor: 8, equipment: 20, material: 31, today: 8, tomorrow: 22,
    crewAm: 34, crewPm: 40,
  };
  var TEXT_LINE_UNITS = 58; // H:S 合併寬約 64 個字寬單位，保守取 58（中文字算 2）
  var CATEGORY_KEY = { "工種": "labor", "機具": "equipment", "材料": "material" };

  // ---------- 小工具 ----------
  function escapeText(s) {
    return String(s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function escapeAttr(s) {
    return escapeText(s).replace(/"/g, "&quot;");
  }
  function round2(n) { return Math.round(n * 100) / 100; }
  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function parseRef(ref) {
    var m = /^([A-Z]+)(\d+)$/.exec(ref);
    return { col: colToNum(m[1]), row: Number(m[2]) };
  }
  function parseRange(range) {
    var p = range.split(":");
    var a = parseRef(p[0]), b = parseRef(p[1] || p[0]);
    return { c1: a.col, r1: a.row, c2: b.col, r2: b.row };
  }

  /** 民國分頁名：2026-09-20 → 115.9.20 */
  function sheetNameForDate(date) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
    if (!m) throw new Error("日期格式錯誤：" + date);
    return (Number(m[1]) - 1911) + "." + Number(m[2]) + "." + Number(m[3]);
  }
  /** 分頁名 → 可排序的 yyyy-mm-dd；不是日期分頁則回傳 null */
  function dateForSheetName(name) {
    var m = /^(\d+)\.(\d+)\.(\d+)$/.exec(name);
    if (!m) return null;
    function p(n) { return ("0" + n).slice(-2); }
    return (Number(m[1]) + 1911) + "-" + p(m[2]) + "-" + p(m[3]);
  }

  /** 顯示寬度：全形字算 2、其餘算 1 */
  function charUnits(ch) {
    var c = ch.charCodeAt(0);
    var wide = (c >= 0x1100 && c <= 0x115f) || (c >= 0x2e80 && c <= 0xa4cf) ||
      (c >= 0xac00 && c <= 0xd7a3) || (c >= 0xf900 && c <= 0xfaff) ||
      (c >= 0xfe30 && c <= 0xfe6f) || (c >= 0xff00 && c <= 0xff60) ||
      (c >= 0xffe0 && c <= 0xffe6) || c === 0x25a1 || c === 0x25a0;
    return wide ? 2 : 1;
  }
  function textUnits(s) {
    var u = 0;
    for (var i = 0; i < s.length; i++) u += charUnits(s.charAt(i));
    return u;
  }
  /** 把一段文字依換行與寬度折成多行 */
  function wrapLines(text, maxUnits) {
    var out = [];
    var paras = String(text || "").replace(/\r/g, "").split("\n");
    for (var i = 0; i < paras.length; i++) {
      var line = "", units = 0, para = paras[i].replace(/\s+$/, "");
      if (para === "") { out.push(""); continue; }
      for (var j = 0; j < para.length; j++) {
        var ch = para.charAt(j), w = charUnits(ch);
        if (units + w > maxUnits) { out.push(line); line = ""; units = 0; }
        line += ch; units += w;
      }
      out.push(line);
    }
    while (out.length && out[out.length - 1] === "") out.pop();
    return out;
  }

  // ---------- sheet XML 儲存格操作 ----------
  function cellRegex(ref) {
    return new RegExp('<c r="' + ref + '"((?:\\s+[A-Za-z:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)');
  }
  /** 寫入儲存格；保留原有樣式(s)。value：數字 / 字串 / 空值(清空) */
  function setCell(xml, ref, value, styleOverride) {
    var m = cellRegex(ref).exec(xml);
    if (!m) throw new Error("範本找不到儲存格 " + ref);
    var sm = /\ss="(\d+)"/.exec(m[1]);
    var s = styleOverride != null ? styleOverride : (sm ? sm[1] : null);
    var sAttr = s !== null ? ' s="' + s + '"' : "";
    var out;
    if (value === null || value === undefined || value === "") {
      out = '<c r="' + ref + '"' + sAttr + "/>";
    } else if (typeof value === "number") {
      if (!isFinite(value)) throw new Error("非有限數字：" + ref);
      out = '<c r="' + ref + '"' + sAttr + "><v>" + String(value) + "</v></c>";
    } else {
      out = '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + escapeText(value) + "</t></is></c>";
    }
    return xml.slice(0, m.index) + out + xml.slice(m.index + m[0].length);
  }
  function styleOfCell(xml, ref) {
    var m = cellRegex(ref).exec(xml);
    if (!m) throw new Error("範本找不到儲存格 " + ref);
    var sm = /\ss="(\d+)"/.exec(m[1]);
    return sm ? sm[1] : null;
  }

  // ---------- 合併儲存格 ----------
  function replaceMerges(xml, mutate) {
    var m = /<mergeCells count="\d+">([\s\S]*?)<\/mergeCells>/.exec(xml);
    if (!m) throw new Error("範本找不到 mergeCells");
    var refs = [], re = /<mergeCell ref="([^"]+)"\/>/g, x;
    while ((x = re.exec(m[1]))) refs.push(x[1]);
    refs = mutate(refs);
    var body = "";
    for (var i = 0; i < refs.length; i++) body += '<mergeCell ref="' + refs[i] + '"/>';
    return xml.slice(0, m.index) + '<mergeCells count="' + refs.length + '">' + body + "</mergeCells>" + xml.slice(m.index + m[0].length);
  }

  // ---------- 樣式：由範本既有樣式衍生新樣式（只改指定屬性，框線等一律保留） ----------
  function parseXfs(stylesXml) {
    var cm = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    if (!cm) throw new Error("styles.xml 找不到 cellXfs");
    var xfs = cm[2].match(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
    if (Number(cm[1]) !== xfs.length) throw new Error("cellXfs 數量與 count 不符");
    return { cm: cm, xfs: xfs };
  }
  function xfFontId(stylesXml, index) {
    return Number(/fontId="(\d+)"/.exec(parseXfs(stylesXml).xfs[index])[1]);
  }
  /**
   * 複製 baseIndex 的樣式並修改：
   *   mods.fontId：換字型；mods.align：對齊屬性 {名稱: 值}，值為 null 代表移除
   * @return {xml, index} 新的 styles.xml 與新樣式編號
   */
  function cloneXf(stylesXml, baseIndex, mods) {
    var p = parseXfs(stylesXml);
    var xm = /^<xf ([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)$/.exec(p.xfs[baseIndex]);
    var attrs = xm[1].replace(/\/$/, ""), inner = xm[2] || "";
    if (mods.fontId != null) {
      attrs = attrs.replace(/fontId="\d+"/, 'fontId="' + mods.fontId + '"');
      if (!/applyFont="1"/.test(attrs)) attrs += ' applyFont="1"';
    }
    var alignAttrs = {};
    var am = /<alignment([^>]*?)\/>/.exec(inner);
    if (am) {
      var re = /([A-Za-z]+)="([^"]*)"/g, x;
      while ((x = re.exec(am[1]))) alignAttrs[x[1]] = x[2];
    }
    var k;
    for (k in (mods.align || {})) {
      if (mods.align[k] === null) delete alignAttrs[k]; else alignAttrs[k] = mods.align[k];
    }
    var alignXml = "";
    for (k in alignAttrs) alignXml += " " + k + '="' + alignAttrs[k] + '"';
    alignXml = alignXml ? "<alignment" + alignXml + "/>" : "";
    inner = am ? inner.replace(am[0], alignXml) : alignXml + inner;
    if (alignXml && !/applyAlignment="1"/.test(attrs)) attrs += ' applyAlignment="1"';
    var neu = "<xf " + attrs + ">" + inner + "</xf>";
    var newXml = stylesXml.slice(0, p.cm.index) + '<cellXfs count="' + (p.xfs.length + 1) + '">' + p.cm[2] + neu + "</cellXfs>" + stylesXml.slice(p.cm.index + p.cm[0].length);
    return { xml: newXml, index: p.xfs.length };
  }
  /** 只改儲存格的樣式編號（不動內容） */
  function restyleCell(xml, ref, styleIndex) {
    var m = cellRegex(ref).exec(xml);
    if (!m) throw new Error("範本找不到儲存格 " + ref);
    var head = m[0].replace(/^<c r="[A-Z]+\d+"/, "");
    var attrs = /^((?:\s+[A-Za-z:]+="[^"]*")*)/.exec(head)[1];
    var newAttrs = /\ss="\d+"/.test(attrs) ? attrs.replace(/\ss="\d+"/, ' s="' + styleIndex + '"') : attrs + ' s="' + styleIndex + '"';
    return xml.slice(0, m.index) + '<c r="' + ref + '"' + newAttrs + head.slice(attrs.length) + xml.slice(m.index + m[0].length);
  }

  function refsOf(col, from, to) {
    var out = [];
    for (var r = from; r <= to; r++) out.push(col + r);
    return out;
  }

  /**
   * 範本前處理（只需做一次，結果存成「已處理範本」供之後每天的分頁使用）：
   *  1. 備註方塊 A46：向左靠齊、靠上、自動換行（框線不變）
   *  2. 工種/機具/材料名稱格：字型統一成範本多數名稱格使用的字型（原本 A9、A10 等是 8pt，其餘 12pt，看起來不一致），
   *     並加上「縮小字型以符合儲存格」，名稱太長時自動縮小而不是被截斷
   *  3. 人員姓名、主任/工安、加班原因、填表人：同樣加上縮小字型以符合
   * @param parts 範本套件 {路徑: 文字}
   * @return 新的套件（不修改傳入的物件）
   */
  function prepareTemplate(parts) {
    var out = {}, k;
    for (k in parts) out[k] = parts[k];
    var styles = out["xl/styles.xml"], sheet = out["xl/worksheets/sheet1.xml"];
    var cache = {};
    function derive(baseIndex, key, mods) {
      var ck = baseIndex + "|" + key;
      if (!(ck in cache)) {
        var r = cloneXf(styles, baseIndex, mods);
        styles = r.xml;
        cache[ck] = r.index;
      }
      return cache[ck];
    }
    function apply(ref, key, mods) {
      var base = Number(styleOfCell(sheet, ref));
      sheet = restyleCell(sheet, ref, derive(base, key, mods));
    }
    var shrink = { align: { shrinkToFit: "1", wrapText: null } };

    // 名稱格的統一字型 = 名稱格中最常用的字型
    var nameRefs = refsOf("A", 8, 18).concat(refsOf("A", 20, 29), refsOf("A", 31, 45));
    var count = {}, best = null;
    nameRefs.forEach(function (ref) {
      var f = xfFontId(styles, Number(styleOfCell(sheet, ref)));
      count[f] = (count[f] || 0) + 1;
      if (best === null || count[f] > count[best]) best = f;
    });
    nameRefs.forEach(function (ref) {
      apply(ref, "name" + best, { fontId: best, align: shrink.align });
    });

    var shrinkRefs = refsOf("M", 34, 45).concat(refsOf("P", 34, 45), ["J34", "J37", "J40", "J43", "O53"]);
    shrinkRefs.forEach(function (ref) { apply(ref, "shrink", shrink); });

    apply("A46", "remark", { align: { horizontal: "left", vertical: "top", wrapText: "1", shrinkToFit: null } });

    out["xl/styles.xml"] = styles;
    out["xl/worksheets/sheet1.xml"] = sheet;
    return out;
  }

  // ---------- 逐日累計 ----------
  /**
   * 算出「每一天」的累計（= 當天 + 之前所有日期），供各日期分頁的累計欄使用。
   * 算法與後端 cumulative 相同：材料＝(上午+下午) 直接加總；工種/機具＝(上午+下午)/2 累加。
   * @param records [{date, name, am, pm}]（全部日期）
   * @param categoryByName {項目名稱: "工種"|"機具"|"材料"}
   * @return {日期: {項目名稱: 累計}}，只含有紀錄的日期；其他日期用 cumulativeAsOf 取值
   */
  function computeCumulatives(records, categoryByName) {
    var sorted = records.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    var running = {}, byDate = {}, i;
    for (i = 0; i < sorted.length; i++) {
      var r = sorted[i];
      var sum = Number(r.am || 0) + Number(r.pm || 0);
      var v = categoryByName[r.name] === "材料" ? sum : sum / 2;
      running[r.name] = (running[r.name] || 0) + v;
      var snap = {};
      for (var n in running) snap[n] = running[n];
      byDate[r.date] = snap;
    }
    return byDate;
  }
  /** 某日期的累計：取「不晚於該日的最近一天」的累計快照（該日沒有紀錄時延續前一天） */
  function cumulativeAsOf(byDate, date) {
    var best = null;
    for (var d in byDate) if (d <= date && (best === null || d > best)) best = d;
    return best === null ? {} : byDate[best];
  }

  // ---------- 挑選要放進範本的項目（容量有限） ----------
  function pickItems(list, capacity, label, warnings) {
    if (list.length <= capacity) return list.slice();
    var withData = [], without = [];
    for (var i = 0; i < list.length; i++) {
      (list[i].am || list[i].pm ? withData : without).push(i);
    }
    var chosen = withData.slice(0, capacity);
    if (withData.length > capacity) {
      warnings.push(label + "有數字的項目共 " + withData.length + " 項，超過範本容量 " + capacity + " 列，多出的項目未列入 xlsx");
    } else {
      warnings.push(label + "項目共 " + list.length + " 項，超過範本容量 " + capacity + " 列，沒有數字的部分項目未列入 xlsx");
      for (var k = 0; k < without.length && chosen.length < capacity; k++) chosen.push(without[k]);
    }
    chosen.sort(function (a, b) { return a - b; });
    return chosen.map(function (idx) { return list[idx]; });
  }

  function box(on) { return on ? "■" : "□"; }

  // ---------- 主要：由範本 sheet XML 產生某一天的 sheet XML ----------
  /**
   * @param templateSheetXml 範本分頁的 XML 字串
   * @param data {date, basic, header, items, attendance}
   *   basic: {公司名稱, 業主, 工程名稱, "合約金額（元）", "開工日期（YYYY/MM/DD）"}
   *   header: {weather,status,todayWork,tomorrowPlan,remark,reporter,directorAm,directorPm,safetyAm,safetyPm}
   *   items: [{category:"工種"|"機具"|"材料", name, am, pm, cumulative}]（依清單順序）
   *   attendance: [{name, am, pm, amHours, pmHours, reason}]
   * @param opts 保留（目前未使用）；templateSheetXml 需先經 prepareTemplate 處理
   * @return {xml, warnings}
   */
  function buildDaySheet(templateSheetXml, data, opts) {
    opts = opts || {};
    var warnings = [];
    var xml = templateSheetXml;
    var b = data.basic || {}, h = data.header || {};
    var dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(data.date);
    if (!dm) throw new Error("日期格式錯誤：" + data.date);

    // 表頭
    xml = setCell(xml, "A1", b["公司名稱"] || "");
    xml = setCell(xml, "A3", "業主：" + (b["業主"] || ""));
    xml = setCell(xml, "A4", "工程名稱：" + (b["工程名稱"] || ""));
    xml = setCell(xml, "A5", "合約金額：" + (b["合約金額（元）"] || ""));
    xml = setCell(xml, "A6", "開工日期：" + (b["開工日期（YYYY/MM/DD）"] || ""));
    xml = setCell(xml, "K5", Number(dm[1]) - 1911);
    xml = setCell(xml, "N5", Number(dm[2]));
    xml = setCell(xml, "P5", Number(dm[3]));
    xml = setCell(xml, "J6", "    天氣：  " + box(h.weather === "晴") + "晴 " + box(h.weather === "陰") + "陰   " +
      box(h.weather === "雨") + "雨   施工狀況：" + box(h.status === "施工") + "施工  " + box(h.status === "休息") + "休息");

    // 工種 / 機具 / 材料
    var groups = { labor: [], equipment: [], material: [] };
    var items = data.items || [];
    for (var i = 0; i < items.length; i++) {
      var key = CATEGORY_KEY[items[i].category];
      if (key) groups[key].push(items[i]);
    }
    var labels = { labor: "工種", equipment: "機具", material: "材料" };
    ["labor", "equipment", "material"].forEach(function (g) {
      var picked = pickItems(groups[g], CAP[g], labels[g], warnings);
      for (var r = 0; r < CAP[g]; r++) {
        var row = ROWS[g] + r, it = picked[r];
        xml = setCell(xml, "A" + row, it ? it.name : "");
        if (g === "material") {
          xml = setCell(xml, "D" + row, it && it.am ? round2(it.am) : "");
        } else {
          xml = setCell(xml, "D" + row, it && it.am ? round2(it.am) : "");
          xml = setCell(xml, "E" + row, it && it.pm ? round2(it.pm) : "");
        }
        xml = setCell(xml, "F" + row, it ? round2(it.cumulative || 0) : "");
      }
    });

    // 本日 / 明日施工項目：逐行寫入 H 欄
    function fillLines(text, startRow, cap, label) {
      var lines = wrapLines(text, TEXT_LINE_UNITS);
      if (lines.length > cap) {
        warnings.push(label + "共 " + lines.length + " 行，超過範本 " + cap + " 行，多出的內容已併入最後一行");
        lines = lines.slice(0, cap - 1).concat([lines.slice(cap - 1).join(" ")]);
      }
      for (var r = 0; r < cap; r++) xml = setCell(xml, "H" + (startRow + r), lines[r] || "");
    }
    fillLines(h.todayWork, ROWS.today, CAP.today, "本日施工項目");
    fillLines(h.tomorrowPlan, ROWS.tomorrow, CAP.tomorrow, "預計明日施工項目");

    // 人員與加班（上午、下午各 6 列）
    var att = data.attendance || [];
    function fillCrew(startRow, shift, director, safety) {
      var hoursKey = shift === "am" ? "amHours" : "pmHours";
      var people = att.filter(function (a) { return a[shift] || a[hoursKey] > 0; });
      var shown = people;
      var extraNames = "";
      if (people.length > CAP.crew) {
        var ot = people.filter(function (a) { return a[hoursKey] > 0; });
        var rest = people.filter(function (a) { return !(a[hoursKey] > 0); });
        var ordered = ot.concat(rest);
        shown = ordered.slice(0, CAP.crew);
        extraNames = ordered.slice(CAP.crew).map(function (a) { return a.name; }).join("、");
        warnings.push((shift === "am" ? "上午" : "下午") + "人員共 " + people.length + " 人，超過範本 " + CAP.crew + " 列，多出的姓名併入最後一列");
      }
      for (var r = 0; r < CAP.crew; r++) {
        var row = startRow + r, p = shown[r];
        var name = p ? p.name : "";
        if (r === CAP.crew - 1 && extraNames) name += "、" + extraNames;
        var hrs = p && p[hoursKey] > 0 ? round2(p[hoursKey]) : "";
        xml = setCell(xml, "M" + row, name);
        xml = setCell(xml, "O" + row, hrs);
        xml = setCell(xml, "P" + row, hrs && p.reason ? p.reason : "");
      }
      xml = setCell(xml, "J" + startRow, director || "");
      xml = setCell(xml, "J" + (startRow + 3), safety || "");
      xml = setCell(xml, "L" + startRow, "本工");
    }
    fillCrew(ROWS.crewAm, "am", h.directorAm, h.safetyAm);
    fillCrew(ROWS.crewPm, "pm", h.directorPm, h.safetyPm);
    xml = setCell(xml, "P33", "加班原因");

    // 備註方塊（A46:S52 併成一格）與填表人
    xml = setCell(xml, "A46", "備註：" + (h.remark || ""));
    xml = setCell(xml, "O53", h.reporter || "");

    // 合併儲存格：備註區併成單一方塊、加班原因欄 P:S 逐列合併
    xml = replaceMerges(xml, function (refs) {
      var kept = refs.filter(function (ref) {
        var r = parseRange(ref);
        return !(r.r1 >= 46 && r.r1 <= 52 && r.c1 >= 1 && r.c2 <= 19);
      });
      kept.push("A46:S52");
      for (var row = 33; row <= 45; row++) kept.push("P" + row + ":S" + row);
      return kept;
    });

    // 只留單一分頁被選取（多個 tabSelected 會讓 Excel 進入「群組」編輯）
    xml = xml.replace(/\s+tabSelected="1"/g, "");
    return { xml: xml, warnings: warnings };
  }

  // ---------- 組成工作簿（所有零件皆為 XML 文字） ----------
  function partMap(parts) { return parts; }

  /**
   * @param base 範本套件（含 styles 等共用零件）的 {路徑: 文字}
   * @param sheets [{name, xml}]，會依日期由舊到新排序
   * @param activeName 要設為開啟時選取的分頁名（預設最新一天）
   * @return {路徑: 文字}
   */
  function assembleWorkbook(base, sheets, activeName) {
    var sorted = sheets.slice().sort(function (a, b) {
      return (dateForSheetName(a.name) || a.name) < (dateForSheetName(b.name) || b.name) ? -1 : 1;
    });
    if (!sorted.length) throw new Error("至少需要一個分頁");
    var active = activeName || sorted[sorted.length - 1].name;
    var out = {};
    var keep = ["_rels/.rels", "docProps/core.xml", "xl/styles.xml", "xl/sharedStrings.xml", "xl/theme/theme1.xml"];
    keep.forEach(function (p) { if (base[p] !== undefined) out[p] = base[p]; });

    var i, sheetTags = "", relTags = "", ctTags = "", names = "", defs = "", activeIdx = 0;
    for (i = 0; i < sorted.length; i++) {
      var n = i + 1, nm = escapeAttr(sorted[i].name);
      var sx = sorted[i].xml.replace(/\s+r:id="rId\d+"(?=\s*\/>)/, "") // 拿掉指向 printerSettings 的參照
        .replace(/\s+tabSelected="1"/g, ""); // 先清掉舊的選取標記，再只標記目前要開啟的分頁
      if (sorted[i].name === active) {
        activeIdx = i;
        sx = sx.replace(/<sheetView /, '<sheetView tabSelected="1" ');
      }
      out["xl/worksheets/sheet" + n + ".xml"] = sx;
      sheetTags += '<sheet name="' + nm + '" sheetId="' + n + '" r:id="rIdSheet' + n + '"/>';
      relTags += '<Relationship Id="rIdSheet' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
      ctTags += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      names += "<vt:lpstr>" + escapeText(sorted[i].name) + "</vt:lpstr>";
      defs += '<definedName name="_xlnm.Print_Area" localSheetId="' + i + '">\'' + nm.replace(/'/g, "''") + "'!$A$1:$S$55</definedName>";
    }

    // workbook.xml
    var wb = base["xl/workbook.xml"];
    wb = wb.replace(/<sheets>[\s\S]*?<\/sheets>/, "<sheets>" + sheetTags + "</sheets>");
    wb = wb.replace(/<definedNames>[\s\S]*?<\/definedNames>/, "<definedNames>" + defs + "</definedNames>");
    wb = wb.replace(/<workbookView ([^>]*?)(\/?)>/, function (all, attrs, close) {
      attrs = attrs.replace(/\s*activeTab="\d+"/, "");
      return "<workbookView " + attrs + ' activeTab="' + activeIdx + '"' + close + ">";
    });
    out["xl/workbook.xml"] = wb;

    // workbook.xml.rels：去掉舊的 worksheet 關聯，加入新的
    var rels = base["xl/_rels/workbook.xml.rels"].replace(/<Relationship [^>]*relationships\/worksheet"[^>]*\/>/g, "");
    out["xl/_rels/workbook.xml.rels"] = rels.replace("</Relationships>", relTags + "</Relationships>");

    // [Content_Types].xml：去掉舊的 sheet 與 printerSettings 宣告，加入新的
    var ct = base["[Content_Types].xml"]
      .replace(/<Override PartName="\/xl\/worksheets\/[^>]*\/>/g, "")
      .replace(/<Default Extension="bin"[^>]*\/>/g, "");
    out["[Content_Types].xml"] = ct.replace("</Types>", ctTags + "</Types>");

    // docProps/app.xml：分頁清單要與工作簿一致
    var app = base["docProps/app.xml"];
    if (app) {
      var total = sorted.length;
      app = app.replace(/<vt:lpstr>工作表<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>\d+<\/vt:i4>/, "<vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>" + total + "</vt:i4>")
        .replace(/<vt:lpstr>具名範圍<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>\d+<\/vt:i4>/, "<vt:lpstr>具名範圍</vt:lpstr></vt:variant><vt:variant><vt:i4>" + total + "</vt:i4>");
      var titles = "";
      for (i = 0; i < sorted.length; i++) titles += "<vt:lpstr>" + escapeText(sorted[i].name) + "</vt:lpstr>";
      for (i = 0; i < sorted.length; i++) titles += "<vt:lpstr>'" + escapeText(sorted[i].name) + "'!Print_Area</vt:lpstr>";
      app = app.replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, '<TitlesOfParts><vt:vector size="' + (total * 2) + '" baseType="lpstr">' + titles + "</vt:vector></TitlesOfParts>");
      out["docProps/app.xml"] = app;
    }
    return out;
  }

  /** 讀出既有工作簿的所有分頁 → [{name, xml}]（用於覆蓋單日分頁或重寫後續分頁） */
  function readSheets(parts) {
    var wb = parts["xl/workbook.xml"], rels = parts["xl/_rels/workbook.xml.rels"];
    var out = [], re = /<sheet name="([^"]*)" sheetId="\d+" r:id="([^"]+)"\/>/g, m;
    while ((m = re.exec(wb))) {
      var rm = new RegExp('<Relationship Id="' + m[2] + '"[^>]*Target="([^"]+)"').exec(rels) ||
        new RegExp('<Relationship [^>]*Target="([^"]+)"[^>]*Id="' + m[2] + '"').exec(rels);
      if (!rm) throw new Error("找不到分頁關聯 " + m[2]);
      var name = m[1].replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"');
      out.push({ name: name, xml: parts["xl/" + rm[1]] });
    }
    return out;
  }

  return {
    CAP: CAP,
    sheetNameForDate: sheetNameForDate,
    dateForSheetName: dateForSheetName,
    wrapLines: wrapLines,
    textUnits: textUnits,
    setCell: setCell,
    styleOfCell: styleOfCell,
    prepareTemplate: prepareTemplate,
    computeCumulatives: computeCumulatives,
    cumulativeAsOf: cumulativeAsOf,
    buildDaySheet: buildDaySheet,
    assembleWorkbook: assembleWorkbook,
    readSheets: readSheets,
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = XlsxBuilder;
