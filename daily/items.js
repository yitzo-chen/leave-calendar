// 工項完成量頁：項目與預估數量來自試算表「合約工項」（日報原表 DPR 的 32 列），每日完成量存到「工項完成量」，
// 送出後後端會在雲端「施工日報彙整」xlsx 加上「工項.民國.月.日」分頁，供文書人員對照填入 DPR 清理版。
// 換算小工具（車次、鋼板樁支數、工天×產能、澆置區塊…）的偏好與區塊表只存在這個瀏覽器，不影響送出的資料。
(function () {
  "use strict";
  const cfg = window.REPORT_CONFIG || {};
  const $ = (id) => document.getElementById(id);
  const D = { blocks: (window.ITEM_BLOCKS || {}).blocks || [], blocksVer: (window.ITEM_BLOCKS || {}).ver || "0" };

  // ---------- 換算工具的本機偏好（每車容量、區塊表等） ----------
  const TOOLS_KEY = "dailyItems.tools.v1";
  let store = { tools: {} };
  try { const s = localStorage.getItem(TOOLS_KEY); if (s) store = JSON.parse(s); } catch (e) { /* 私密瀏覽等情境略過 */ }
  if (!store || typeof store !== "object") store = { tools: {} };
  if (!store.tools || typeof store.tools !== "object") store.tools = {};
  function persist() { try { localStorage.setItem(TOOLS_KEY, JSON.stringify(store)); } catch (e) { /* 略過 */ } }

  // ---------- 共用小工具 ----------
  const r2 = (x) => Math.round(x * 100) / 100;
  const fmt = (n) => Number(r2(n).toFixed(2)).toLocaleString("zh-TW", { maximumFractionDigits: 2 });
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  function num(s) { if (s == null) return 0; const t = String(s).replace(/,/g, "").trim(); if (t === "") return 0; const v = Number(t); return isFinite(v) && v >= 0 ? v : NaN; }
  function ymd(d) { const p = (n) => String(n).padStart(2, "0"); return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()); }
  function shift(s, n) { const [y, m, d] = s.split("-").map(Number); return ymd(new Date(y, m - 1, d + n)); }
  function makeClientId() { return window.crypto && crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(36).slice(2); }

  // ---------- 狀態 ----------
  let currentCaseId = "", date = ymd(new Date()), ITEMS = [], itemsSig = "", serverPrior = {}, loadedValues = {}, draft = {}, dirty = false, dayLoaded = false;
  const rows = {};
  function prior(id) { return serverPrior[id] || 0; } // 到前一天為止的累計（期初已完成＋之前各日），由後端算好
  function setDirty(v) { dirty = v; $("dirty").hidden = !v; }

  async function apiGet(action, params) {
    const url = new URL(cfg.SCRIPT_URL);
    url.searchParams.set("action", action);
    if (currentCaseId) url.searchParams.set("caseId", currentCaseId);
    Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));
    return (await fetch(url.toString())).json();
  }
  async function apiPost(payload) {
    const res = await fetch(cfg.SCRIPT_URL, { method: "POST", body: JSON.stringify(Object.assign({ caseId: currentCaseId }, payload)), headers: { "Content-Type": "text/plain;charset=utf-8" } });
    return res.json();
  }


  // ---------- 換算小工具：只算出數字帶入本日量，約定值（每車容量、係數等）記在這個瀏覽器 ----------
  // 項目 id＝日報原表列號：3 擋土壁(補強柱) 4 鋼板樁 5/6/7 開挖/回填/棄土 8/9 PC/結構混凝土 10~12 模板 13 鋼筋 30~32 保麗龍/防水
  let TOOLKIND = { '3': 'pct', '4': 'pile', '5': 'trip', '6': 'trip', '7': 'trip', '8': 'blocks', '9': 'blocks', '10': 'manp', '11': 'manp', '12': 'manp', '13': 'manp', '30': 'blocks', '31': 'blocks', '32': 'blocks' };
  const HP = 341; // SS-01 H 型鋼補強柱支數：圍幕周長 272.8 m ÷ 0.8 m 間距
  function pref(k, f, d) { const t = store.tools[k]; return t && t[f] != null ? t[f] : d; }
  function setPref(k, f, v) { (store.tools[k] = store.tools[k] || {})[f] = v; persist(); }
  function parseList(s) {
    const toks = s.split(/[\s,，、+＋;；]+/).filter(Boolean); let sum = 0;
    for (const t of toks) { const v = Number(t); if (!isFinite(v) || v < 0) return { msg: '「' + t + '」不是數字' }; sum += v; }
    return toks.length ? { sum, n: toks.length } : { msg: '請輸入數字' };
  }
  function makeTool(it, panel) {
    const kind = TOOLKIND[it.id] || 'sum', isTon = it.unit === 'ton', isConc = false;
    if (kind === 'blocks') return makeBlocksTool(it, panel);
    const fld = (f, lab) => '<label>' + lab + '<input class="s" data-f="' + f + '" inputmode="decimal" autocomplete="off"></label>';
    let h = '<div class="tf">', note = '';
    if (kind === 'trip') {
      h += fld('n', '車次') + '<span>×</span>' + fld('cap', '每車 m³') + '<span>×</span>' + fld('coef', '換算係數');
      note = '預設一台土車 7 m³（依你的約定）；要換算鬆方請改係數，留 1＝不換算。改過的值會記在這個瀏覽器。';
    } else if (kind === 'manp') {
      const reb = it.id === '13';
      h += fld('mp', reb ? '當日鋼筋工（工天）' : '當日模板工（工天）') + '<span>×</span>' + fld('pr', reb ? '每人日綁紮 ton' : '每人日施作 m²');
      note = '工天＝（上午人數＋下午人數）÷ 2；正式版會自動帶入「出工」卡片的' + (reb ? '鋼筋工' : '模板工') + '。每人日產能預設採進度表的用值（' + (reb ? '牆柱 0.8、板與基礎 1.0 t/人日' : '牆柱 16、頂版與樓板 30 m²/人日；拆模 50～100') + '），請改成你們實際的值，會記在這個瀏覽器。';
    } else if (kind === 'pct') {
      h += fld('hn', '本日打設補強柱（支）') + '<span>或</span>' + fld('pc', '本日新增完成 %');
      note = '第 3 列（合約 A-3）原為型鋼內支撐，已改為鋼板樁配合 H 型鋼補強柱（SS-01 約 ' + HP + ' 支＝272.8 m ÷ 0.8 m）。可直接填本日打設支數，自動換成 %；換算量＝% × 合約量 ' + fmt(it.qty) + ' ' + it.unit + '（合約量為初估）。以補強柱支數認定進度是我的推測，請確認。';
    } else if (kind === 'pile') {
      h += '<label>算法<select data-f="mode"><option value="len">樁長合計：支數 × 單支長</option><option value="wid">牆面延長米：支數 × 單支有效寬</option></select></label>' +
        '<label data-z>區域<select data-f="zone"><option value="SS-01">SS-01（16 m 樁）</option><option value="SS-02">SS-02（6 m 樁）</option></select></label>' +
        fld('n', '支數') + '<span>×</span><label><span data-f="plab"></span><input class="s" data-f="p" inputmode="decimal" autocomplete="off"></label>';
      note = '日報以實際施作量填。預設用「牆面延長米」：FSP-IV 有效寬 0.4 m，SS-01 牆長 272.8 m（682 支）＋SS-02 牆長 71 m（約 178 支）＝343.8 m，對合約 395 m（初估）少 51.2 m。若要改填樁長合計，切換算法：SS-01 樁長 16 m、SS-02 樁長 6 m。';
    } else {
      h += '<label style="flex:1 1 100%">多筆數字（用空格、逗號或 + 分開）<textarea data-f="list" placeholder="例：8.5, 9, 7.5"></textarea></label>';
      if (isTon) h += '<label>輸入單位<select data-f="u"><option value="kg">kg（自動 ÷1000 換成 ton）</option><option value="ton">ton</option></select></label>';
      if (isConc) h += fld('loss', '不計入 %');
      note = isConc ? '完成量建議以澆置區塊的設計體積為準，送貨單合計只用來查核；「不計入 %」可填損耗或超打比例。' : '把多筆數字加總後帶入本日量。';
    }
    h += '</div><div class="res"></div><div class="acts"><button type="button" class="ghost-btn" data-a="set">取代本日量</button><button type="button" class="ghost-btn" data-a="add">加到本日量</button></div><div class="note">' + note + '</div>';
    panel.innerHTML = h;
    const q = f => panel.querySelector('[data-f="' + f + '"]'), res = panel.querySelector('.res'), btns = panel.querySelectorAll('[data-a]');
    let val = null;
    if (kind === 'trip') { q('cap').value = pref('trip', 'cap', '7'); q('coef').value = pref('trip', 'coef', '1'); }
    if (kind === 'pile') { q('mode').value = pref('pile', 'mode', 'wid'); q('zone').value = pref('pile', 'zone', 'SS-01'); }
    if (kind === 'manp') q('pr').value = pref('manp', it.id === '13' ? 'rebar' : 'form', it.id === '13' ? '0.8' : '16');
    if (isTon && kind === 'sum') q('u').value = pref('ton', 'u', 'kg');
    if (isConc) q('loss').value = pref('conc', 'loss', '0');
    function pileSync() {
      const m = q('mode').value, z = q('zone').value;
      panel.querySelector('[data-z]').style.display = m === 'len' ? '' : 'none';
      q('plab').textContent = m === 'len' ? '單支長 m' : '單支有效寬 m';
      q('p').value = m === 'len' ? pref('pile', 'len_' + z, z === 'SS-01' ? '16' : '6') : pref('pile', 'wid', '0.4');
    }
    if (kind === 'pile') pileSync();
    function calc() {
      val = null; let msg = '', extra = '';
      const n = f => num(q(f).value), blank = f => !q(f).value.trim();
      if (kind === 'trip') {
        if (blank('n') || blank('cap')) msg = '請填車次與每車容量';
        else if (isNaN(n('n')) || isNaN(n('cap')) || isNaN(n('coef'))) msg = '有欄位不是數字';
        else { val = n('n') * n('cap') * (blank('coef') ? 1 : n('coef')); setPref('trip', 'cap', q('cap').value); setPref('trip', 'coef', q('coef').value); }
      } else if (kind === 'manp') {
        if (blank('mp') || blank('pr')) msg = '請填工天與每人日產能';
        else if (isNaN(n('mp')) || isNaN(n('pr'))) msg = '有欄位不是數字';
        else { val = n('mp') * n('pr'); setPref('manp', it.id === '13' ? 'rebar' : 'form', q('pr').value); }
      } else if (kind === 'pct') {
        if (blank('pc')) msg = '請填本日新增完成 %';
        else if (isNaN(n('pc')) || n('pc') > 100) msg = '請填 0～100 的數字';
        else { val = n('pc') / 100 * it.qty; extra = '（累計將達 ' + ((prior(it.id, date) + val) / it.qty * 100).toFixed(1) + '%）'; }
      } else if (kind === 'pile') {
        if (blank('n') || blank('p')) msg = '請填支數與' + q('plab').textContent;
        else if (isNaN(n('n')) || isNaN(n('p'))) msg = '有欄位不是數字';
        else {
          val = n('n') * n('p'); const m = q('mode').value, z = q('zone').value;
          setPref('pile', 'mode', m); setPref('pile', 'zone', z); setPref('pile', m === 'len' ? 'len_' + z : 'wid', q('p').value);
          extra = m === 'len' ? '（' + z + '）' : '';
        }
      } else {
        const r = parseList(q('list').value);
        if (r.msg) msg = r.msg;
        else {
          let v = r.sum; extra = '（' + r.n + ' 筆合計 ' + fmt(r.sum) + '）';
          if (isTon) { setPref('ton', 'u', q('u').value); if (q('u').value === 'kg') v /= 1000; }
          if (isConc) { const l = n('loss'); if (isNaN(l) || l < 0 || l >= 100) msg = '不計入 % 要在 0～99'; else { v *= 1 - l / 100; setPref('conc', 'loss', q('loss').value); } }
          if (!msg) val = v;
        }
      }
      res.classList.toggle('err', !!msg);
      res.textContent = msg ? msg : '= ' + fmt(val) + ' ' + it.unit + extra;
      btns.forEach(b => { b.disabled = val == null; });
    }
    function put(v) { draft[it.id] = String(r2(v)); setDirty(true); rows[it.id].inp.value = draft[it.id]; refresh(); }
    panel.querySelector('[data-a="set"]').onclick = () => { if (val != null) put(val); };
    panel.querySelector('[data-a="add"]').onclick = () => { if (val != null) put((num(draft[it.id]) || 0) + val); };
    panel.addEventListener('input', e => { if (e.target.dataset.f === 'mode' || e.target.dataset.f === 'zone') pileSync(); if (e.target.dataset.f === 'hn' && e.target.value.trim() && !isNaN(num(e.target.value))) q('pc').value = String(r2(num(e.target.value) / HP * 100)); calc(); });
    panel.addEventListener('change', e => { if (e.target.dataset.f === 'mode' || e.target.dataset.f === 'zone') pileSync(); if (e.target.dataset.f === 'hn' && e.target.value.trim() && !isNaN(num(e.target.value))) q('pc').value = String(r2(num(e.target.value) / HP * 100)); calc(); });
    calc();
  }
  
  // 澆置區塊（A-10 PC 墊層、A-11 結構混凝土）：勾選今天完成的區塊 → 帶入混凝土量，A-11 可一併帶入鋼筋(A-15)與模板(A-12/13/14)
  function getBlk() {
    let b = store.tools.blk;
    if (!b || b.ver !== D.blocksVer) { b = store.tools.blk = { ver: D.blocksVer, list: D.blocks.map(x => Object.assign({}, x)) }; persist(); }
    return b;
  }
  function makeBlocksTool(it, panel) {
    const B = getBlk(), ticked = {}, isVol = it.unit === 'm3';
    const mine = () => B.list.filter(b => b.item === it.id);
    panel.innerHTML =
      '<div class="blk-list"></div>' +
      '<div class="res" data-r="sum"></div>' +
      '<div class="acts"><button type="button" class="ghost-btn" data-a="set">取代本日量</button><button type="button" class="ghost-btn" data-a="add">加到本日量</button></div>' +
      '<details class="blk-more"><summary>編輯區塊表／送貨單查核</summary>' +
      '<label style="margin-top:8px">批次加入（每行：名稱 數量 [鋼筋t] [模板m²]，用空格或逗號分開）<textarea data-f="bulk" placeholder="例：底版-A區 120.5"></textarea></label>' +
      '<div class="acts"><button type="button" class="ghost-btn" data-a="badd">加入區塊表</button><button type="button" class="ghost-btn danger" data-a="brest">還原初稿</button><span class="hint bmsg" style="margin:0"></span></div>' +
      (isVol ? '<label style="margin-top:10px">送貨單查核（只比對、不帶入；多筆 m³ 用空格、逗號或 + 分開）<textarea data-f="list" placeholder="例：8.5, 9, 7.5"></textarea></label><div class="res" data-r="chk"></div>' : '') +
      '</details>' +
      '<div class="note">區塊表是初稿（詳見 澆置區塊數量表_初稿_1151107.xlsx）：混凝土依經理版（115.11.02）澆置工項分段，兩版昇層高程不同處按高度比例調整；保麗龍與防水毯依底版面積、外牆周長×高度概估，沒有縮放到合約量，所以全勾完可能與合約量不同。勾選今天完成的區塊即帶入該量；區塊表的修改記在這個瀏覽器；正式版會記錄每個區塊的完成日，避免重複計算。</div>';
    const q = f => panel.querySelector('[data-f="' + f + '"]');
    const list = panel.querySelector('.blk-list'), res = panel.querySelector('[data-r=sum]'), chk = panel.querySelector('[data-r=chk]');
    const setB = panel.querySelector('[data-a=set]'), addB = panel.querySelector('[data-a=add]'), msgEl = panel.querySelector('.bmsg'), restB = panel.querySelector('[data-a=brest]');
    let tot = { v: 0, r: 0, f: 0, n: 0 }, armed = false;
    function calc() {
      const on = mine().filter(b => ticked[b.id]);
      tot = { v: r2(on.reduce((s, b) => s + b.vol, 0)), r: r2(on.reduce((s, b) => s + (b.rebar || 0), 0)), f: Math.round(on.reduce((s, b) => s + (b.form || 0), 0)), n: on.length };
      res.textContent = '= ' + fmt(tot.v) + ' ' + it.unit + '（已勾 ' + tot.n + ' 個區塊）';
      setB.disabled = addB.disabled = tot.n === 0;
      if (!chk) return;
      const raw = q('list').value.trim();
      chk.classList.remove('err');
      if (!raw) { chk.textContent = ''; return; }
      const r = parseList(raw);
      if (r.msg) { chk.classList.add('err'); chk.textContent = r.msg; return; }
      chk.textContent = '送貨單合計 ' + fmt(r.sum) + ' m³（' + r.n + ' 筆）' + (tot.v > 0 ? '，與設計體積差 ' + fmt(r.sum - tot.v) + '（' + ((r.sum - tot.v) / tot.v * 100).toFixed(1) + '%）' : '；尚未勾區塊，無法比對');
    }
    function drawList() {
      const ms = mine();
      list.innerHTML = ms.length ? '' : '<div class="hint" style="margin:8px">區塊表是空的，請在下方批次加入或還原初稿。</div>';
      ms.forEach(b => {
        const r = document.createElement('div'); r.className = 'blk';
        r.innerHTML = '<label class="blk-l"><input type="checkbox"><span class="blk-n"></span></label><span class="blk-v"></span><button type="button" class="ghost-btn" aria-label="刪除此區塊">✕</button>';
        const nm = r.querySelector('.blk-n'); nm.textContent = (b.site ? b.site + ' ' : '') + b.name;
        if (b.pid || b.date) { const s = document.createElement('small'); s.className = 'blk-d'; s.textContent = (b.pid || '') + (b.date ? ' 預定 ' + b.date.slice(2).replace(/-/g, '/') : ''); nm.appendChild(s); }
        r.querySelector('.blk-v').textContent = fmt(b.vol) + ' ' + it.unit;
        r.title = '鋼筋 ' + (b.rebar || 0) + ' t／模板 ' + (b.form || 0) + ' m²' + (b.src ? '／' + b.src : '') + (b.note ? '／' + b.note : '');
        const cb = r.querySelector('input'); cb.checked = !!ticked[b.id];
        cb.addEventListener('change', () => { ticked[b.id] = cb.checked; calc(); });
        r.querySelector('button').addEventListener('click', () => { B.list = B.list.filter(x => x.id !== b.id); delete ticked[b.id]; persist(); drawList(); calc(); });
        list.appendChild(r);
      });
    }
    function put(add) {
      const apply = (id, v) => { draft[id] = String(r2((add ? (num(draft[id]) || 0) : 0) + v)); setDirty(true); rows[id].inp.value = draft[id]; };
      apply(it.id, tot.v);
      refresh();
    }
    panel.querySelector('[data-a=badd]').addEventListener('click', () => {
      const ta = q('bulk'), bad = []; let add = 0;
      ta.value.split(/\n+/).map(s => s.trim()).filter(Boolean).forEach(line => {
        const m = line.match(/^(.*?)[\s,，]+(\d+(?:\.\d+)?)(?:[\s,，]+(\d+(?:\.\d+)?))?(?:[\s,，]+(\d+(?:\.\d+)?))?$/);
        if (m && m[1].trim()) { B.list.push({ id: 'u' + Date.now().toString(36) + B.list.length, site: '', name: m[1].trim(), item: it.id, vol: Number(m[2]), rebar: Number(m[3] || 0), form: Number(m[4] || 0), pid: '', date: '', src: '手動加入', note: '' }); add++; } else bad.push(line);
      });
      ta.value = bad.join('\n'); persist(); drawList(); calc();
      msgEl.textContent = '已加入 ' + add + ' 個區塊' + (bad.length ? '；' + bad.length + ' 行格式不對，留在輸入框' : '');
    });
    restB.addEventListener('click', () => {
      if (!armed) { armed = true; restB.textContent = '再按一次確認還原'; restB.classList.add('armed'); setTimeout(() => { armed = false; restB.textContent = '還原初稿'; restB.classList.remove('armed'); }, 4000); return; }
      armed = false; restB.textContent = '還原初稿'; restB.classList.remove('armed');
      B.list = D.blocks.map(x => Object.assign({}, x)); Object.keys(ticked).forEach(k => delete ticked[k]); persist(); drawList(); calc();
      msgEl.textContent = '已還原成初稿（所有工項的區塊表一起還原）';
    });
    setB.addEventListener('click', () => { if (tot.n) put(false); });
    addB.addEventListener('click', () => { if (tot.n) put(true); });
    if (q('list')) q('list').addEventListener('input', calc);
    drawList(); calc();
  }

  // ---------- 畫面 ----------
  function build() {
    if (currentCaseId !== "lng") TOOLKIND = {}; // 換算工具是依洲際 LNG 案的項目序號設計的，其他案場只提供「多筆加總」
    const box = $("groups"); box.innerHTML = "";
    Object.keys(rows).forEach((k) => delete rows[k]);
    ITEMS.forEach((it) => {
      const wrap = document.createElement("div"); wrap.className = "grp solo";
      const row = document.createElement("div"); row.className = "row";
      row.innerHTML =
        '<div><div class="nm"><span class="seq">' + it.seq + "</span>" + esc(it.name) + "</div>" +
        '<div class="meta"><span>預估 <b>' + fmt(it.qty) + "</b> " + esc(it.unit) + '</span><span>累計 <b class="cum"></b></span><span class="rem"></span><span class="pct"></span><button type="button" class="tool-btn" aria-expanded="false">🧮 換算</button></div>' +
        '<div class="bar"><i></i></div></div>' +
        '<div class="qty"><input type="text" inputmode="decimal" autocomplete="off" aria-label="' + esc(it.seq + " " + it.name + " 本日完成量") + '" placeholder="0"><span class="u">' + esc(it.unit) + "</span></div>";
      const inp = row.querySelector("input");
      inp.addEventListener("input", () => { draft[it.id] = inp.value; setDirty(true); refresh(); syncSubmit(); });
      const tb = row.querySelector(".tool-btn"), panel = document.createElement("div");
      panel.className = "tool"; panel.hidden = true; row.appendChild(panel);
      tb.addEventListener("click", () => {
        const open = panel.hidden; panel.hidden = !open; tb.setAttribute("aria-expanded", String(open));
        if (open && !panel.dataset.ready) { makeTool(it, panel); panel.dataset.ready = "1"; }
      });
      wrap.appendChild(row); wrap._items = [it]; box.appendChild(wrap);
      rows[it.id] = { el: row, inp, cum: row.querySelector(".cum"), rem: row.querySelector(".rem"), pct: row.querySelector(".pct"), bar: row.querySelector(".bar i"), it };
    });
  }

  function refresh() {
    let filled = 0, overN = 0, bad = 0;
    const lines = [];
    ITEMS.forEach((it) => {
      const R = rows[it.id], raw = draft[it.id];
      let v = num(raw); const isBad = isNaN(v); if (isBad) { v = 0; bad++; }
      const cum = r2(prior(it.id) + v), over = it.qty > 0 && cum > it.qty + 1e-9, pct = it.qty ? cum / it.qty * 100 : 0;
      R.cum.textContent = fmt(cum) + " " + it.unit;
      R.rem.textContent = over ? "超出 " + fmt(cum - it.qty) : "剩 " + fmt(it.qty - cum);
      R.pct.textContent = pct.toFixed(1) + "%";
      R.bar.style.width = Math.min(100, pct) + "%";
      R.el.classList.toggle("over", over); R.el.classList.toggle("bad", isBad); R.el.classList.toggle("filled", v > 0);
      if (document.activeElement !== R.inp) R.inp.value = raw == null ? "" : raw;
      if (v > 0) { filled++; lines.push(it.name + "　本日 " + fmt(v) + " " + it.unit); }
      if (over) overN++;
    });
    $("summary").innerHTML = "<span>本日有填 <b>" + filled + "</b> 項</span>" +
      (overN ? '<span class="warn">累計超出預估 ' + overN + " 項（只提醒，不擋送出）</span>" : "<span>累計都在預估數量內</span>") +
      (bad ? '<span class="warn">' + bad + " 格不是數字</span>" : "");
    $("gen").value = lines.map((l, i) => (i + 1) + ". " + l).join("\n");
  }

  function showStatus(msg, kind) { const s = $("status"); s.textContent = msg; s.className = "status" + (kind ? " " + kind : ""); s.hidden = false; }
  function syncSubmit() {
    const noPw = $("fPassword").value.trim() === "";
    $("saveBtn").disabled = noPw || !dayLoaded;
    $("pwHint").textContent = !dayLoaded ? "資料讀取成功後才能送出" : noPw ? "請先輸入通關密碼才能送出" : "";
  }
  function syncLinks() {
    const q = "?case=" + encodeURIComponent(currentCaseId) + "&date=" + encodeURIComponent(date);
    $("backLink").href = "index.html" + q;
  }

  // ---------- 載入／送出 ----------
  let loadSeq = 0;
  async function loadDay(d) {
    const my = ++loadSeq;
    dayLoaded = false; syncSubmit();
    $("loadStatus").textContent = "讀取中…";
    try {
      const res = await apiGet("workItems", { date: d });
      if (my !== loadSeq) return;
      if (!res.ok) throw new Error(res.error || "讀取失敗");
      if (!Array.isArray(res.items)) throw new Error("後端尚未更新，找不到工項完成量功能");
      date = d; $("fDate").value = d;
      const sig = JSON.stringify(res.items.map((i) => [i.seq, i.name, i.unit, i.qty]));
      if (sig !== itemsSig) { ITEMS = res.items.map((i) => ({ id: String(i.seq), seq: i.seq, name: i.name, unit: i.unit, qty: Number(i.qty) || 0 })); itemsSig = sig; build(); }
      if (!ITEMS.length) showStatus("這個案場還沒有「合約工項」資料，請到試算表「合約工項」分頁新增。", "error");
      serverPrior = res.prior || {}; loadedValues = {}; draft = {};
      Object.keys(res.values || {}).forEach((k) => { loadedValues[k] = Number(res.values[k]); draft[k] = String(res.values[k]); });
      setDirty(false); dayLoaded = true; syncLinks(); refresh();
      $("loadStatus").textContent = Object.keys(loadedValues).length ? "已載入這天已存的資料" : "";
    } catch (err) {
      if (my !== loadSeq) return;
      $("loadStatus").textContent = ""; showStatus("讀取失敗：" + err.message, "error");
    }
    syncSubmit();
  }

  async function save() {
    const pw = $("fPassword").value.trim();
    if (!pw) { showStatus("請先輸入通關密碼", "error"); return; }
    if (!dayLoaded) { showStatus("資料還沒讀取成功，不能送出", "error"); return; }
    const items = [], clear = [], bad = [];
    ITEMS.forEach((it) => {
      const v = num(draft[it.id]);
      if (isNaN(v)) { bad.push(it.seq + " " + it.name); return; }
      if (v > 0) items.push({ seq: it.seq, qty: r2(v) });
      else if (loadedValues[it.id] > 0) clear.push(it.seq); // 載入時有值、現在被清空：視為使用者明確清除
    });
    if (bad.length) { showStatus("這些項目的完成量不是數字：" + bad.join("、"), "error"); return; }
    if (!items.length && !clear.length) { showStatus("沒有要送出的內容：本日全部空白，這天也沒有存過資料。", "error"); return; }
    $("saveBtn").disabled = true; showStatus("送出中…");
    try {
      const res = await apiPost({ action: "saveWorkItems", date, password: pw, clientId: makeClientId(), items, clearItems: clear });
      if (!res.ok) { showStatus("送出失敗：" + (res.error || "未知錯誤"), "error"); return; }
      let msg = "已存入 " + date + "：寫入 " + res.saved + " 項" + (res.cleared ? "、清除 " + res.cleared + " 項" : "") + "。";
      if (res.xlsx && res.xlsx.ok && !res.xlsx.skipped) { msg += "雲端「施工日報彙整」已新增／更新「工項」分頁。"; if (res.xlsx.url) xlsxUrl = res.xlsx.url; }
      else if (res.xlsx && !res.xlsx.ok) msg += "但雲端 xlsx 更新失敗（" + res.xlsx.error + "）。資料已存入試算表，不會遺失；之後再送出一次或請管理者處理即可。";
      showStatus(msg, res.xlsx && !res.xlsx.ok ? "error" : "");
      await loadDay(date);
    } catch (err) {
      showStatus("送出失敗：" + err.message, "error");
    } finally { syncSubmit(); }
  }

  let xlsxUrl = "";
  async function openXlsx() {
    if (xlsxUrl) { window.open(xlsxUrl, "_blank", "noopener"); return; }
    const win = window.open("", "_blank");
    try {
      const r = await apiGet("xlsxUrl");
      if (r.ok && r.url && /^https:\/\//.test(r.url)) { xlsxUrl = r.url; if (win) { win.opener = null; win.location.href = r.url; } return; }
    } catch (e) { /* 往下顯示訊息 */ }
    if (win) win.close();
    showStatus("彙整檔尚未產生：第一次送出工項完成量或日報後就會建立。");
  }

  function gotoDate(d) {
    if (dirty && !window.confirm("有尚未送出的修改，切換日期會捨棄它們。確定要切換嗎？")) { $("fDate").value = date; return; }
    $("status").hidden = true; loadDay(d);
  }

  async function init() {
    if (!cfg.SCRIPT_URL) { $("errorBanner").hidden = false; $("errorBanner").textContent = "尚未設定 config.js 的 SCRIPT_URL，暫時無法讀取/送出資料。"; return; }
    const qs = new URLSearchParams(location.search);
    const qDate = qs.get("date");
    if (qDate && /^\d{4}-\d{2}-\d{2}$/.test(qDate)) date = qDate;
    $("fDate").value = date;
    $("prev").addEventListener("click", () => gotoDate(shift(date, -1)));
    $("next").addEventListener("click", () => gotoDate(shift(date, 1)));
    $("today").addEventListener("click", () => gotoDate(ymd(new Date())));
    $("fDate").addEventListener("change", (e) => { if (e.target.value) gotoDate(e.target.value); });
    $("fPassword").addEventListener("input", syncSubmit);
    $("saveBtn").addEventListener("click", save);
    $("openXlsxBtn").addEventListener("click", openXlsx);
    $("openAll").hidden = true; $("closeAll").hidden = true;
    $("q").addEventListener("input", () => {
      const k = $("q").value.trim().toLowerCase(); let any = false;
      document.querySelectorAll("#groups > .grp").forEach((g) => {
        const it = g._items[0], hit = !k || (it.seq + " " + it.name).toLowerCase().includes(k);
        g.style.display = hit ? "" : "none"; if (hit) any = true;
      });
      $("msgEmpty").hidden = any;
    });
    $("copyBtn").addEventListener("click", () => {
      const t = $("gen").value; if (!t) { $("copyMsg").textContent = "還沒有內容可複製"; return; }
      const ok = () => { $("copyMsg").textContent = "已複製"; }, fb = () => { $("gen").select(); $("copyMsg").textContent = "已選取，請手動複製（Ctrl+C）"; };
      try { navigator.clipboard.writeText(t).then(ok, fb); } catch (e) { fb(); }
    });
    $("clearToolsBtn").addEventListener("click", () => { store.tools = {}; persist(); $("clearToolsMsg").textContent = "已清除換算工具的記憶（每車容量、區塊表修改等）；重新整理後生效。"; });
    syncSubmit();
    try {
      const cs = await apiGet("cases");
      if (!cs.ok) throw new Error(cs.error || "讀取案場失敗");
      let stored = ""; try { stored = localStorage.getItem("dailyReport.caseId") || ""; } catch (e) { /* 略過 */ }
      const want = qs.get("case") || stored, list = cs.cases || [];
      const found = list.find((c) => c.code === want) || list[0];
      if (!found) throw new Error("沒有可用的案場");
      currentCaseId = found.code;
      $("caseName").textContent = (found.name ? found.name + "（" + found.code + "）" : found.code);
      $("caseLoadStatus").textContent = "";
    } catch (err) { $("caseLoadStatus").textContent = "讀取案場失敗：" + err.message; return; }
    syncLinks();
    await loadDay(date);
  }
  init();
})();
