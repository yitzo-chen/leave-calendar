(function () {
  const cfg = window.REPORT_CONFIG || {};
  function el(id) { return document.getElementById(id); }

  let ITEMS = []; // {category, name, workCode}
  let baseline = {}; // 累計到前一天為止的值，key=項目名稱
  const rowsByName = {};

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

  // 材料是直接加總用量，工種/機具是「工天數」邏輯(上午+下午)/2 —— 要跟後端算法一致
  function itemValue(name, am, pm) {
    const item = ITEMS.find((it) => it.name === name);
    const sum = (Number(am) || 0) + (Number(pm) || 0);
    return item && item.category === "材料" ? sum : sum / 2;
  }

  function renderItemList(container, items, isMaterial) {
    container.innerHTML = "";
    items.forEach((item) => {
      const row = document.createElement("div");
      row.className = "item-row" + (isMaterial ? " material" : "");

      const nameSpan = document.createElement("span");
      nameSpan.className = "item-name";
      nameSpan.textContent = item.name;
      row.appendChild(nameSpan);

      const amInput = document.createElement("input");
      amInput.type = "number";
      amInput.min = "0";
      amInput.step = "0.5";
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

      container.appendChild(row);
      rowsByName[item.name] = { row, amInput, pmInput, refresh };
    });
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
    el("projectName").textContent = result.basic["工程名稱"] || "";
    if (result.basic["公司名稱"]) el("companyName").textContent = result.basic["公司名稱"];
    renderItemList(el("laborList"), ITEMS.filter((it) => it.category === "工種"), false);
    renderItemList(el("equipmentList"), ITEMS.filter((it) => it.category === "機具"), false);
    renderItemList(el("materialList"), ITEMS.filter((it) => it.category === "材料"), true);
  }

  async function loadReporters() {
    const result = await apiGet("reporters");
    if (!result.ok) return;
    const datalist = el("reporterNames");
    datalist.innerHTML = "";
    result.names.forEach((n) => {
      const opt = document.createElement("option");
      opt.value = n;
      datalist.appendChild(opt);
    });
  }

  async function loadDay(date) {
    const status = el("loadStatus");
    status.textContent = "讀取中…";
    clearItemInputs();
    el("fWeather").value = "晴";
    el("fStatus").value = "施工";
    el("fTodayWork").value = "";
    el("fTomorrowPlan").value = "";
    el("fOvertimeAM").value = "";
    el("fOvertimePM").value = "";
    el("fRemark").value = "";

    const [dayResult, cumResult] = await Promise.all([
      apiGet("day", { date }),
      apiGet("cumulative", { upTo: addDays(date, -1) }),
    ]);

    baseline = (cumResult.ok && cumResult.totals) || {};

    if (dayResult.ok && dayResult.header) {
      const h = dayResult.header;
      el("fWeather").value = h.weather || "晴";
      el("fStatus").value = h.status || "施工";
      el("fTodayWork").value = h.todayWork || "";
      el("fTomorrowPlan").value = h.tomorrowPlan || "";
      el("fOvertimeAM").value = h.overtimeAM || "";
      el("fOvertimePM").value = h.overtimePM || "";
      el("fRemark").value = h.remark || "";
      if (h.reporter) el("fReporter").value = h.reporter;
      status.textContent = "已載入當天既有資料，修改後送出即覆蓋更新";
    } else {
      status.textContent = "這天還沒有資料，直接填寫送出即可";
    }

    if (dayResult.ok && dayResult.records) {
      dayResult.records.forEach((r) => {
        const target = rowsByName[r.name];
        if (!target) return;
        target.amInput.value = r.am || "";
        if (target.pmInput) target.pmInput.value = r.pm || "";
      });
    }

    refreshAllCumulative();
  }

  function setStatus(text, kind) {
    const box = el("formStatus");
    box.hidden = false;
    box.textContent = text;
    box.className = "form-status" + (kind ? " " + kind : "");
  }

  function syncSubmitLock() {
    el("submitBtn").disabled = el("fPassword").value.trim() === "";
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

    try {
      await loadConfig();
      await loadReporters();
      await loadDay(today);
    } catch (err) {
      el("errorBanner").hidden = false;
      el("errorBanner").textContent = "讀取失敗：" + err.message;
    }

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
          overtimeAM: el("fOvertimeAM").value.trim(),
          overtimePM: el("fOvertimePM").value.trim(),
          remark: el("fRemark").value.trim(),
          reporter,
        },
        items: collectItems(),
      };

      try {
        const result = await apiPost(payload);
        if (!result.ok) throw new Error(result.error || "送出失敗");
        setStatus("已送出！", "success");
        await loadDay(date);
      } catch (err) {
        setStatus("送出失敗：" + err.message + "（若一直失敗，稍等半分鐘再試一次）", "error");
      } finally {
        syncSubmitLock();
      }
    });
  });
})();
