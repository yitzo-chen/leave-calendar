(function () {
  const cfg = window.REPORT_CONFIG || {};
  function el(id) { return document.getElementById(id); }

  let ITEMS = []; // {category, name, workCode}
  let baseline = {}; // 累計到前一天為止的值，key=項目名稱
  let BASIC = {}; // 基本資料（業主/工程名稱/合約金額等），列印用
  const rowsByName = {};
  let dayLoaded = false; // 當天資料成功載入後才允許送出，避免空白表單覆蓋真實資料
  let loadedItemNames = []; // 載入時「有數字」的項目名稱，用來算出使用者清空了哪些
  let loadedAttendanceNames = []; // 載入時已存在的出勤人員，用來算出使用者移除了誰

  function toDateInputValue(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  function addDays(dateStr, n) {
    const d = new Date(dateStr + "T00:00:00");
    d.setDate(d.getDate() + n);
    return toDateInputValue(d);
  }
  function round2(n) {
    return Math.round(n * 100) / 100;
  }
  function toRocParts(dateStr) {
    const [y, m, d] = dateStr.split("-").map(Number);
    return { year: y - 1911, month: m, day: d };
  }
  function makeClientId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  async function apiGet(action, params) {
    const url = new URL(cfg.SCRIPT_URL);
    url.searchParams.set("action", action);
    Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));
    const res = await fetch(url.toString());
    return res.json();
  }
  async function apiPost(payload) {
    const res = await fetch(cfg.SCRIPT_URL, {
      method: "POST",
      body: JSON.stringify(payload),
      headers: { "Content-Type": "text/plain;charset=utf-8" },
    });
    return res.json();
  }

  // ---------- 本工人員出勤：動態新增/刪除列 ----------
  function addAttendanceRow(data) {
    const row = document.createElement("div");
    row.className = "attendance-row";

    function makeInput(cls, type, value, placeholder) {
      const input = document.createElement("input");
      input.type = type;
      input.className = cls;
      if (placeholder) input.placeholder = placeholder;
      if (type === "checkbox") input.checked = !!value;
      else input.value = value || "";
      if (type === "number") {
        input.min = "0";
        input.step = "0.5";
        input.inputMode = "decimal";
      }
      return input;
    }

    const nameInput = makeInput("att-name", "text", data && data.name, "姓名");
    nameInput.maxLength = 20;
    nameInput.setAttribute("list", "peopleNames");
    row.appendChild(nameInput);
    row.appendChild(makeInput("att-am", "checkbox", data && data.am));
    row.appendChild(makeInput("att-pm", "checkbox", data && data.pm));
    row.appendChild(makeInput("att-am-hours", "number", data && data.amHours));
    row.appendChild(makeInput("att-pm-hours", "number", data && data.pmHours));

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "remove-btn";
    removeBtn.textContent = "✕";
    removeBtn.addEventListener("click", () => row.remove());
    row.appendChild(removeBtn);

    const reasonInput = makeInput("att-reason", "text", data && data.reason, "加班原因（有加班才填）");
    reasonInput.maxLength = 40;
    row.appendChild(reasonInput);

    el("attendanceList").appendChild(row);
  }

  function clearAttendanceRows() {
    el("attendanceList").innerHTML = "";
  }

  function collectAttendance() {
    return Array.from(el("attendanceList").querySelectorAll(".attendance-row"))
      .map((row) => ({
        name: row.querySelector(".att-name").value.trim(),
        am: row.querySelector(".att-am").checked,
        pm: row.querySelector(".att-pm").checked,
        amHours: Number(row.querySelector(".att-am-hours").value || 0),
        pmHours: Number(row.querySelector(".att-pm-hours").value || 0),
        reason: row.querySelector(".att-reason").value.trim(),
      }))
      .filter((a) => a.name);
  }

  // ---------- 內部記錄：選工（結構化）或備註（純文字）兩種列，僅後台查看，不進列印/xlsx ----------
  const PICK_TIMES = ["上午", "下午", "全天"];

  function buildOptions(values, selected) {
    return values.map((v) => `<option value="${v}"${v === selected ? " selected" : ""}>${v}</option>`).join("");
  }

  function addPickRow(data) {
    const row = document.createElement("div");
    row.className = "pick-row";
    row.innerHTML = `
      <input type="text" class="pick-category" list="pickCategoryList" maxlength="20" placeholder="類別">
      <input type="text" class="pick-content" maxlength="60" placeholder="內容說明">
      <select class="pick-time">${buildOptions(PICK_TIMES, data && data.time)}</select>
      <input type="text" class="pick-remark" maxlength="60" placeholder="備註（選填）">
      <button type="button" class="remove-btn">✕</button>
    `;
    row.querySelector(".pick-category").value = (data && data.category) || "";
    row.querySelector(".pick-content").value = (data && data.content) || "";
    row.querySelector(".pick-remark").value = (data && data.remark) || "";
    row.querySelector(".remove-btn").addEventListener("click", () => row.remove());
    el("pickList").appendChild(row);
  }

  function addNoteRow(data) {
    const row = document.createElement("div");
    row.className = "note-row";
    row.innerHTML = `<input type="text" class="note-content" maxlength="2000" placeholder="想記錄的內容"><button type="button" class="remove-btn">✕</button>`;
    row.querySelector(".note-content").value = (data && data.content) || "";
    row.querySelector(".remove-btn").addEventListener("click", () => row.remove());
    el("noteList").appendChild(row);
  }

  function clearInternalRows() {
    el("pickList").innerHTML = "";
    el("noteList").innerHTML = "";
  }

  function collectInternalNotes() {
    const picks = Array.from(el("pickList").querySelectorAll(".pick-row")).map((row) => ({
      type: "選工",
      category: row.querySelector(".pick-category").value.trim(),
      content: row.querySelector(".pick-content").value.trim(),
      time: row.querySelector(".pick-time").value,
      remark: row.querySelector(".pick-remark").value.trim(),
    })).filter((p) => p.content);
    const notes = Array.from(el("noteList").querySelectorAll(".note-row")).map((row) => ({
      type: "備註",
      content: row.querySelector(".note-content").value.trim(),
    })).filter((n) => n.content);
    return picks.concat(notes);
  }

  // ---------- 出工/機具/材料：新增項目 ----------
  // 新增成功後直接把該筆項目插進畫面，不整批重新讀取設定——
  // 一來避免 Sheets 寫入後緊接著讀取可能的短暫延遲讓新項目「看起來沒加成功」，
  // 二來避免重新整理把使用者正在其他項目上打到一半的數字清空。
  const LIST_CONTAINER = { 工種: "laborList", 機具: "equipmentList", 材料: "materialList" };

  // 訊息顯示在「＋新增項目」那一列下方，不用捲到頁面最底部才看到結果
  function showAddMsg(input, text, kind) {
    const row = input.closest(".add-item-row");
    let msg = row.nextElementSibling;
    if (!msg || !msg.classList.contains("add-msg")) {
      msg = document.createElement("p");
      msg.className = "field-hint add-msg";
      row.after(msg);
    }
    msg.textContent = text;
    msg.style.color = kind === "error" ? "var(--danger)" : "";
  }

  async function addItemPrompt(category, inputId) {
    const input = el(inputId);
    const name = input.value.trim();
    if (!name) return;
    const password = el("fPassword").value;
    if (!password.trim()) {
      showAddMsg(input, "請先到頁面最下方輸入通關密碼，才能新增項目", "error");
      return;
    }
    showAddMsg(input, "新增中…", "");
    try {
      const result = await apiPost({ action: "addItem", category, name, password });
      if (!result.ok) throw new Error(result.error || "新增失敗");
      input.value = "";
      if (!ITEMS.some((it) => it.name === name)) {
        ITEMS.push({ category, name, workCode: "" });
      }
      const container = el(LIST_CONTAINER[category]);
      container.appendChild(buildItemRow({ category, name, workCode: "" }, category === "材料"));
      showAddMsg(input, `已新增「${name}」`, "");
    } catch (err) {
      showAddMsg(input, "新增項目失敗：" + err.message, "error");
    }
  }

  // 材料是直接加總用量，工種/機具是「工天數」邏輯(上午+下午)/2 —— 要跟後端算法一致
  function itemValue(name, am, pm) {
    const item = ITEMS.find((it) => it.name === name);
    const sum = (Number(am) || 0) + (Number(pm) || 0);
    return item && item.category === "材料" ? sum : sum / 2;
  }

  function buildItemRow(item, isMaterial) {
    const row = document.createElement("div");
    row.className = "item-row" + (isMaterial ? " material" : "");

    const nameSpan = document.createElement("span");
    nameSpan.className = "item-name";
    nameSpan.textContent = item.name;
    row.appendChild(nameSpan);

    const amInput = document.createElement("input");
    amInput.type = "number";
    amInput.min = "0";
    amInput.step = isMaterial ? "any" : "0.5"; // 材料用量常有小數（如 3.25 m³），工種/機具以半天為單位
    amInput.inputMode = "decimal";
    row.appendChild(amInput);

    let pmInput = null;
    if (!isMaterial) {
      pmInput = document.createElement("input");
      pmInput.type = "number";
      pmInput.min = "0";
      pmInput.step = "0.5";
      pmInput.inputMode = "decimal";
      row.appendChild(pmInput);
    }

    const cumSpan = document.createElement("span");
    cumSpan.className = "item-cumulative";
    cumSpan.textContent = "累計 0";
    row.appendChild(cumSpan);

    function refresh() {
      const base = baseline[item.name] || 0;
      const cur = itemValue(item.name, amInput.value, pmInput ? pmInput.value : 0);
      cumSpan.textContent = "累計 " + round2(base + cur);
    }
    amInput.addEventListener("input", refresh);
    if (pmInput) pmInput.addEventListener("input", refresh);

    rowsByName[item.name] = { row, amInput, pmInput, refresh };
    return row;
  }

  function renderItemList(container, items, isMaterial) {
    container.innerHTML = "";
    items.forEach((item) => container.appendChild(buildItemRow(item, isMaterial)));
  }

  function refreshAllCumulative() {
    Object.values(rowsByName).forEach((r) => r.refresh());
  }

  function collectItems() {
    return Object.entries(rowsByName).map(([name, r]) => ({
      name,
      am: Number(r.amInput.value || 0),
      pm: r.pmInput ? Number(r.pmInput.value || 0) : 0,
    }));
  }

  function clearItemInputs() {
    Object.values(rowsByName).forEach((r) => {
      r.amInput.value = "";
      if (r.pmInput) r.pmInput.value = "";
    });
  }

  async function loadConfig() {
    const result = await apiGet("config");
    if (!result.ok) throw new Error(result.error || "讀取設定失敗");
    ITEMS = result.items;
    BASIC = result.basic || {};
    el("projectName").textContent = result.basic["工程名稱"] || "";
    if (result.basic["公司名稱"]) el("companyName").textContent = result.basic["公司名稱"];
    renderItemList(el("laborList"), ITEMS.filter((it) => it.category === "工種"), false);
    renderItemList(el("equipmentList"), ITEMS.filter((it) => it.category === "機具"), false);
    renderItemList(el("materialList"), ITEMS.filter((it) => it.category === "材料"), true);
  }

  async function fillDatalist(action, datalistId) {
    const result = await apiGet(action);
    if (!result.ok) return;
    const datalist = el(datalistId);
    datalist.innerHTML = "";
    result.names.forEach((n) => {
      const opt = document.createElement("option");
      opt.value = n;
      datalist.appendChild(opt);
    });
  }
  async function loadReporters() {
    await fillDatalist("reporters", "reporterNames");
    await fillDatalist("peopleNames", "peopleNames");
    await fillDatalist("internalCategories", "pickCategoryList");
  }

  let loadDaySeq = 0; // 快速連切日期時，只採用最後一次請求的回應

  async function loadDay(date) {
    const mySeq = ++loadDaySeq;
    dayLoaded = false;
    syncSubmitLock();
    const status = el("loadStatus");
    status.textContent = "讀取中…";
    clearItemInputs();
    clearAttendanceRows();
    clearInternalRows();
    el("fWeather").value = "晴";
    el("fStatus").value = "施工";
    el("fTodayWork").value = "";
    el("fTomorrowPlan").value = "";
    el("fRemark").value = "";
    ["fDir", "fSafe"].forEach((id) => { el(id).value = ""; });

    const [dayResult, cumResult] = await Promise.all([
      apiGet("day", { date }),
      apiGet("cumulative", { upTo: addDays(date, -1) }),
    ]);

    if (mySeq !== loadDaySeq) return;

    baseline = (cumResult.ok && cumResult.totals) || {};

    if (dayResult.ok && dayResult.header) {
      const h = dayResult.header;
      el("fWeather").value = h.weather || "晴";
      el("fStatus").value = h.status || "施工";
      el("fTodayWork").value = h.todayWork || "";
      el("fTomorrowPlan").value = h.tomorrowPlan || "";
      el("fRemark").value = h.remark || "";
      el("fDir").value = h.directorAm || h.directorPm || "";
      el("fSafe").value = h.safetyAm || h.safetyPm || "";
      if (h.reporter) el("fReporter").value = h.reporter;
      status.textContent = "已載入當天既有資料，修改後送出即覆蓋更新";
    } else {
      status.textContent = "這天還沒有資料，直接填寫送出即可";
    }

    loadedItemNames = [];
    loadedAttendanceNames = [];
    if (dayResult.ok && dayResult.records) {
      dayResult.records.forEach((r) => {
        if (Number(r.am) || Number(r.pm)) loadedItemNames.push(r.name);
        const target = rowsByName[r.name];
        if (!target) return;
        target.amInput.value = r.am || "";
        if (target.pmInput) target.pmInput.value = r.pm || "";
      });
    }

    if (dayResult.ok && dayResult.attendance) {
      dayResult.attendance.forEach((a) => {
        addAttendanceRow(a);
        loadedAttendanceNames.push(String(a.name).trim());
      });
    }

    // 內部記錄卡片刻意不預填既有資料（送出只會新增、不會覆蓋，畫面每次切換日期都是空白的新增入口）

    dayLoaded = !!dayResult.ok;
    syncSubmitLock();
    if (!dayResult.ok) status.textContent = "當天資料讀取失敗，無法送出，請重新整理頁面";

    refreshAllCumulative();
  }

  function setStatus(text, kind) {
    const box = el("formStatus");
    box.hidden = false;
    box.textContent = text;
    box.className = "form-status" + (kind ? " " + kind : "");
  }

  // 日報送出成功後，依後端回傳的 xlsx 更新狀態顯示；xlsx 失敗不影響日報已送出的事實
  function showSubmitResult(xlsx) {
    const box = el("formStatus");
    if (!xlsx) { setStatus("已送出！", "success"); return; }
    if (!xlsx.ok) {
      setStatus("日報已送出，但 xlsx 更新失敗：" + (xlsx.error || "未知錯誤") + "（稍後可再送出一次重試）", "error");
      return;
    }
    const warnings = Array.isArray(xlsx.warnings) ? xlsx.warnings : [];
    setStatus("已送出！xlsx 已更新" + (warnings.length ? "，但有警告：" + warnings.join("；") : ""), warnings.length ? "error" : "success");
    if (xlsx.url && /^https:\/\//.test(xlsx.url)) {
      xlsxUrl = xlsx.url; // 檔案若被重建，網址會變；以最新一次送出的結果為準
      const a = document.createElement("a");
      a.href = xlsx.url;
      a.target = "_blank";
      a.rel = "noopener";
      a.textContent = "開啟 xlsx";
      box.appendChild(document.createTextNode(" "));
      box.appendChild(a);
    }
  }

  // ---------- 開啟完整試算表（雲端 xlsx 彙整檔）----------
  let xlsxUrl = ""; // 目前已知的 xlsx 網址；尚未產生（還沒送出過日報）時為空

  // 向後端查詢 xlsx 網址。回傳 { url } 或 { error }；舊版後端不認得此動作時回 error
  async function fetchXlsxUrl() {
    try {
      const result = await apiGet("xlsxUrl");
      if (!result.ok) return { error: result.error || "讀取失敗" };
      if (typeof result.url !== "string") return { error: "後端尚未支援此功能" };
      if (result.url && !/^https:\/\//.test(result.url)) return { error: "網址格式不正確" };
      return { url: result.url };
    } catch (err) {
      return { error: "網路連線失敗" };
    }
  }

  function flashLoadStatus(text) {
    const status = el("loadStatus");
    status.textContent = text;
    setTimeout(() => { if (status.textContent === text) status.textContent = ""; }, 5000);
  }

  async function openXlsx() {
    if (xlsxUrl) { window.open(xlsxUrl, "_blank", "noopener"); return; }
    // 沒有快取網址時要先問後端；先開空白分頁再導向，避免非同步之後才 window.open 被瀏覽器擋掉
    const win = window.open("", "_blank");
    const result = await fetchXlsxUrl();
    if (result.url) {
      xlsxUrl = result.url;
      if (win) { win.opener = null; win.location.href = result.url; } else window.open(result.url, "_blank", "noopener");
      return;
    }
    if (win) win.close();
    flashLoadStatus(result.error ? "無法開啟試算表：" + result.error : "試算表尚未產生，第一次送出日報後就會建立");
  }

  function syncSubmitLock() {
    const noPassword = el("fPassword").value.trim() === "";
    el("submitBtn").disabled = noPassword || !dayLoaded;
    el("passwordHint").textContent = !dayLoaded
      ? "當天資料讀取完成後才能送出"
      : noPassword ? "請先輸入密碼才能送出" : "";
  }

  // ---------- 日報列印（比照休假月曆「假單輸出」機制：填進隱藏範本 → window.print()） ----------
  function buildItemRows(category, isMaterial) {
    return ITEMS.filter((it) => it.category === category)
      .map((item) => {
        const r = rowsByName[item.name];
        const am = r ? r.amInput.value || "" : "";
        const pm = r && r.pmInput ? r.pmInput.value || "" : "";
        const cum = round2((baseline[item.name] || 0) + itemValue(item.name, am, pm));
        return isMaterial
          ? `<tr><td>${esc(item.name)}</td><td>${am}</td><td>${cum}</td></tr>`
          : `<tr><td>${esc(item.name)}</td><td>${am}</td><td>${pm}</td><td>${cum}</td></tr>`;
      })
      .join("");
  }

  // 本工人員出勤 → 列印範本「人員／加班／加班原因」區。
  // 一個時段一個區塊（至少6列，對應範例每時段6格）：左側 主任/工安 姓名各佔一半列數，
  // 「本工」為固定文字；右側每人一列，加班時數與原因印在同一列。
  // 有登記該時段加班時數的人，即使沒勾該時段出勤也會列入。
  function esc(v) {
    return String(v).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  function buildPeopleRows(attendance) {
    const shifts = [
      { label: "上午", key: "am", hoursKey: "amHours", director: esc(el("fDir").value.trim()), safety: esc(el("fSafe").value.trim()) },
      { label: "下午", key: "pm", hoursKey: "pmHours", director: esc(el("fDir").value.trim()), safety: esc(el("fSafe").value.trim()) },
    ];
    return shifts.map((sh) => {
      const people = attendance.filter((a) => a[sh.key] || a[sh.hoursKey] > 0);
      let total = Math.max(6, people.length);
      if (total % 2) total += 1;
      const half = total / 2;
      let html = "";
      for (let i = 0; i < total; i++) {
        const p = people[i];
        const hours = p && p[sh.hoursKey] > 0 ? p[sh.hoursKey] : "";
        html += "<tr>";
        if (i === 0) html += `<td class="pr-people-shift" rowspan="${total}">${sh.label}</td>`;
        if (i === 0) html += `<td class="pr-people-label" rowspan="${half}">主任</td><td rowspan="${half}">${sh.director}</td>`;
        if (i === half) html += `<td class="pr-people-label" rowspan="${half}">工安</td><td rowspan="${half}">${sh.safety}</td>`;
        if (i === 0) html += `<td class="pr-people-label" rowspan="${total}">本工</td>`;
        html += `<td>${p ? esc(p.name) : ""}</td><td>${hours}</td><td>${hours ? esc(p.reason) : ""}</td></tr>`;
      }
      return html;
    }).join("");
  }

  function buildPrintReport() {
    const date = el("fDate").value;
    el("prCompany").textContent = BASIC["公司名稱"] || "";
    el("prOwner").textContent = BASIC["業主"] || "";
    el("prProject").textContent = BASIC["工程名稱"] || "";
    el("prContract").textContent = BASIC["合約金額（元）"] || "";
    const roc = toRocParts(date);
    el("prYear").textContent = roc.year;
    el("prMonth").textContent = roc.month;
    el("prDay").textContent = roc.day;
    el("prStartDate").textContent = BASIC["開工日期（YYYY/MM/DD）"] || "";

    const weather = el("fWeather").value;
    el("prChkSun").textContent = weather === "晴" ? "☑" : "□";
    el("prChkCloud").textContent = weather === "陰" ? "☑" : "□";
    el("prChkRain").textContent = weather === "雨" ? "☑" : "□";
    const status = el("fStatus").value;
    el("prChkWork").textContent = status === "施工" ? "☑" : "□";
    el("prChkRest").textContent = status === "休息" ? "☑" : "□";

    el("prLaborRows").innerHTML = buildItemRows("工種", false);
    el("prEquipmentRows").innerHTML = buildItemRows("機具", false);
    el("prMaterialRows").innerHTML = buildItemRows("材料", true);

    el("prPeopleRows").innerHTML = buildPeopleRows(collectAttendance());

    el("prTodayWork").textContent = el("fTodayWork").value.trim();
    el("prTomorrowPlan").textContent = el("fTomorrowPlan").value.trim();

    el("prRemark").textContent = el("fRemark").value.trim();
    el("prReporter").textContent = el("fReporter").value.trim() || "－";
  }

  document.addEventListener("DOMContentLoaded", async () => {
    const dateInput = el("fDate");
    const today = toDateInputValue(new Date());
    dateInput.value = today;

    if (!cfg.SCRIPT_URL) {
      el("errorBanner").hidden = false;
      el("errorBanner").textContent = "尚未設定 config.js 的 SCRIPT_URL，暫時無法讀取/送出資料。";
      el("submitBtn").disabled = true;
      return;
    }

    el("fPassword").addEventListener("input", syncSubmitLock);
    syncSubmitLock();

    dateInput.addEventListener("change", () => loadDay(dateInput.value));
    el("prevDayBtn").addEventListener("click", () => {
      dateInput.value = addDays(dateInput.value, -1);
      loadDay(dateInput.value);
    });
    el("nextDayBtn").addEventListener("click", () => {
      dateInput.value = addDays(dateInput.value, 1);
      loadDay(dateInput.value);
    });
    el("todayBtn").addEventListener("click", () => {
      dateInput.value = toDateInputValue(new Date());
      loadDay(dateInput.value);
    });
    el("laborAddBtn").addEventListener("click", () => addItemPrompt("工種", "laborNewName"));
    el("equipmentAddBtn").addEventListener("click", () => addItemPrompt("機具", "equipmentNewName"));
    el("materialAddBtn").addEventListener("click", () => addItemPrompt("材料", "materialNewName"));
    el("addPickBtn").addEventListener("click", () => addPickRow());
    el("addNoteBtn").addEventListener("click", () => addNoteRow());

    // 內部記錄卡片預設隱藏，只在電腦按 Ctrl+Alt+I 才切換顯示（僅供工地主任查看，手機不支援組合鍵）
    document.addEventListener("keydown", (e) => {
      if (e.ctrlKey && e.altKey && e.key.toLowerCase() === "i") {
        e.preventDefault();
        const sec = el("internalSection");
        sec.hidden = !sec.hidden;
        if (!sec.hidden) sec.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
    el("addAttendanceBtn").addEventListener("click", () => addAttendanceRow());
    el("openXlsxBtn").addEventListener("click", openXlsx);
    fetchXlsxUrl().then((r) => { if (r.url) xlsxUrl = r.url; }); // 先取好網址，點擊時可直接開啟
    el("printReportBtn").addEventListener("click", () => {
      buildPrintReport();
      document.body.classList.add("printing-report");
      setTimeout(() => window.print(), 50);
    });
    window.addEventListener("afterprint", () => {
      document.body.classList.remove("printing-report");
    });

    el("reportForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      const date = dateInput.value;
      const reporter = el("fReporter").value.trim();
      if (!date || !reporter) {
        setStatus("日期與填表人為必填", "error");
        return;
      }
      const submitBtn = el("submitBtn");
      submitBtn.disabled = true;
      setStatus("送出中…", "");

      const payload = {
        date,
        password: el("fPassword").value,
        clientId: makeClientId(),
        header: {
          weather: el("fWeather").value,
          status: el("fStatus").value,
          todayWork: el("fTodayWork").value.trim(),
          tomorrowPlan: el("fTomorrowPlan").value.trim(),
          remark: el("fRemark").value.trim(),
          // 表單只填一次；後端 日報頭 仍分上午/下午兩欄，這裡兩欄寫同一個值，列印時上下午區塊各印一次
          directorAm: el("fDir").value.trim(),
          directorPm: el("fDir").value.trim(),
          safetyAm: el("fSafe").value.trim(),
          safetyPm: el("fSafe").value.trim(),
          reporter,
        },
        items: collectItems(),
        attendance: collectAttendance(),
        internalNotes: collectInternalNotes(),
        // 載入時有數字、現在兩格都清空的項目；載入時存在、現在已不在名單裡的人員 → 後端才會刪除
        clearItems: collectItems().filter((it) => !it.am && !it.pm && loadedItemNames.includes(it.name)).map((it) => it.name),
        removeAttendance: loadedAttendanceNames.filter((n) => !collectAttendance().some((a) => a.name === n)),
      };

      try {
        const result = await apiPost(payload);
        if (!result.ok) throw new Error(result.error || "送出失敗");
        showSubmitResult(result.xlsx);
        el("internalSection").hidden = true; // 送出成功後自動收合，平常畫面上預設看不到，只有按 Ctrl+Alt+I 當下才會打開
        await loadDay(date);
      } catch (err) {
        setStatus("送出失敗：" + err.message + "（若一直失敗，稍等半分鐘再試一次）", "error");
      } finally {
        syncSubmitLock();
      }
    });

    try {
      await loadConfig();
      await loadReporters().catch(() => {}); // 下拉選單載入失敗不影響當天資料
      await loadDay(today);
    } catch (err) {
      el("errorBanner").hidden = false;
      el("errorBanner").textContent = "讀取失敗：" + err.message;
    }
  });
})();
