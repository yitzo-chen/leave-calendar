(function () {
  const cfg = window.LEAVE_CONFIG || {};
  const COLORS = ["#C8630A", "#2E7D6B", "#3E5C9A", "#9A4C8F", "#B08900",
                  "#5B7A3A", "#A14E4E", "#4A7A9A", "#7A5A9A", "#8A6A3A"];

  const state = {
    year: new Date().getFullYear(),
    month: new Date().getMonth() + 1, // 1-12
    employees: [],   // [{name, color}]
    records: [],     // [{name, start:Date, end:Date, type}]
    activeEmployee: null,
    demo: false,
  };

  const el = (id) => document.getElementById(id);

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    el("companyName").textContent = cfg.COMPANY_NAME || "員工休假月曆";
    document.title = cfg.COMPANY_NAME || "員工休假月曆";
    el("prevMonth").addEventListener("click", () => shiftMonth(-1));
    el("nextMonth").addEventListener("click", () => shiftMonth(1));
    el("todayBtn").addEventListener("click", goToday);
    el("printBtn").addEventListener("click", () => window.print());
    el("downloadBtn").addEventListener("click", downloadCSV);
    setupLeaveForm();

    if (cfg.SHEET_ID) {
      el("sheetLink").href = `https://docs.google.com/spreadsheets/d/${cfg.SHEET_ID}/edit`;
      await loadFromSheet();
    } else {
      loadDemoData();
    }

    render();
  }

  async function loadFromSheet() {
    try {
      const [roster, log] = await Promise.all([
        fetchSheetCsv(cfg.SHEET_ID, cfg.ROSTER_SHEET_NAME),
        fetchSheetCsv(cfg.SHEET_ID, cfg.LOG_SHEET_NAME),
      ]);
      state.employees = parseRoster(roster);
      state.records = parseLog(log);
      el("errorBanner").hidden = true;
      return true;
    } catch (err) {
      showError("讀取 Google 試算表失敗：" + err.message + "（請確認試算表已設定「知道連結的使用者」可檢視，且分頁名稱與 config.js 相符）");
      loadDemoData();
      return false;
    }
  }

  function loadDemoData() {
    state.demo = true;
    el("demoBanner").hidden = false;
    el("sheetLink").style.display = "none";
    const y = state.year, m = state.month;
    state.employees = [
      { name: "王小明", color: COLORS[0] },
      { name: "陳大文", color: COLORS[1] },
      { name: "林美玲", color: COLORS[2] },
    ];
    state.records = [
      { name: "王小明", start: new Date(y, m - 1, 20), end: new Date(y, m - 1, 22), type: "全天" },
      { name: "陳大文", start: new Date(y, m - 1, 5), end: new Date(y, m - 1, 5), type: "上午半天" },
      { name: "林美玲", start: new Date(y, m - 1, 15), end: new Date(y, m - 1, 16), type: "全天" },
    ];
  }

  function showError(msg) {
    const b = el("errorBanner");
    b.hidden = false;
    b.textContent = msg;
  }

  // ---------- CSV ----------
  async function fetchSheetCsv(sheetId, sheetName) {
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const text = await res.text();
    if (text.trim().startsWith("<")) throw new Error("試算表尚未公開或分頁名稱不符");
    return parseCsv(text);
  }

  function parseCsv(text) {
    const rows = [];
    let row = [], field = "", inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else field += c;
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === ",") {
        row.push(field); field = "";
      } else if (c === "\n") {
        row.push(field); rows.push(row); row = []; field = "";
      } else if (c === "\r") {
        // skip
      } else {
        field += c;
      }
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows.filter((r) => r.some((v) => v !== ""));
  }

  function parseRoster(rows) {
    return rows.slice(1)
      .map((r) => (r[0] || "").trim())
      .filter(Boolean)
      .map((name, i) => ({ name, color: COLORS[i % COLORS.length] }));
  }

  function parseLog(rows) {
    return rows.slice(1)
      .filter((r) => (r[0] || "").trim())
      .map((r) => {
        const name = r[0].trim();
        const start = parseDate(r[1]);
        const end = parseDate(r[2]) || start;
        const type = (r[3] || "全天").trim();
        return start ? { name, start, end, type } : null;
      })
      .filter(Boolean);
  }

  function parseDate(v) {
    if (!v) return null;
    v = v.trim();
    // Google gviz 常把日期輸出成 "Date(2026,8,20)"（月份 0-based）
    const m = v.match(/^Date\((\d+),(\d+),(\d+)/);
    if (m) return new Date(+m[1], +m[2], +m[3]);
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }

  // ---------- calendar math ----------
  function shiftMonth(delta) {
    let m = state.month + delta, y = state.year;
    if (m < 1) { m = 12; y--; } else if (m > 12) { m = 1; y++; }
    state.month = m; state.year = y;
    render();
  }

  function goToday() {
    const now = new Date();
    state.year = now.getFullYear();
    state.month = now.getMonth() + 1;
    render();
  }

  function sameDay(a, b) {
    return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }

  function statusFor(name, date) {
    let full = false, am = false, pm = false;
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    for (const r of state.records) {
      if (r.name !== name) continue;
      const d0 = new Date(r.start.getFullYear(), r.start.getMonth(), r.start.getDate());
      const d1 = new Date(r.end.getFullYear(), r.end.getMonth(), r.end.getDate());
      if (d >= d0 && d <= d1) {
        if (r.type === "全天") full = true;
        else if (r.type === "上午半天") am = true;
        else if (r.type === "下午半天") pm = true;
      }
    }
    return { full, am, pm };
  }

  // ---------- render ----------
  function render() {
    el("monthLabel").textContent = `${state.year} 年 ${state.month} 月`;
    renderRoster();
    renderGrid();
  }

  function renderRoster() {
    const ul = el("employeeList");
    ul.innerHTML = "";
    el("empCount").textContent = state.employees.length;
    const datalist = el("empNames");
    datalist.innerHTML = state.employees.map((e) => `<option value="${escapeHtml(e.name)}">`).join("");
    state.employees.forEach((emp) => {
      const days = monthlyDays(emp.name);
      const li = document.createElement("li");
      li.className = "employee-item" + (state.activeEmployee === emp.name ? " active" : "");
      li.innerHTML =
        `<i class="dot" style="background:${emp.color}"></i>` +
        `<span>${escapeHtml(emp.name)}</span>` +
        `<span class="days">${days ? days + "天" : ""}</span>`;
      li.addEventListener("click", () => {
        state.activeEmployee = state.activeEmployee === emp.name ? null : emp.name;
        render();
      });
      ul.appendChild(li);
    });
  }

  function monthlyDays(name) {
    let total = 0;
    const daysInMonth = new Date(state.year, state.month, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
      const s = statusFor(name, new Date(state.year, state.month - 1, d));
      if (s.full) total += 1;
      else if (s.am || s.pm) total += 0.5;
    }
    return total ? String(total).replace(/\.0$/, "") : "";
  }

  function renderGrid() {
    const grid = el("calendarGrid");
    grid.innerHTML = "";
    const first = new Date(state.year, state.month - 1, 1);
    const startOffset = first.getDay(); // 0=Sun
    const daysInMonth = new Date(state.year, state.month, 0).getDate();
    const today = new Date();

    for (let i = 0; i < startOffset; i++) {
      const cell = document.createElement("div");
      cell.className = "day-cell empty";
      grid.appendChild(cell);
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(state.year, state.month - 1, d);
      const weekday = date.getDay();
      const cell = document.createElement("div");
      cell.className =
        "day-cell" +
        (weekday === 0 || weekday === 6 ? " weekend" : "") +
        (sameDay(date, today) ? " today" : "");

      const num = document.createElement("div");
      num.className = "day-num";
      num.textContent = d;
      cell.appendChild(num);

      const stack = document.createElement("div");
      stack.className = "chip-stack";
      state.employees.forEach((emp) => {
        const s = statusFor(emp.name, date);
        if (!s.full && !s.am && !s.pm) return;
        const dimmed = state.activeEmployee && state.activeEmployee !== emp.name;
        if (s.full) stack.appendChild(makeChip(emp, "全", dimmed, false));
        if (s.am) stack.appendChild(makeChip(emp, "上", dimmed, true));
        if (s.pm) stack.appendChild(makeChip(emp, "下", dimmed, true));
      });
      cell.appendChild(stack);
      grid.appendChild(cell);
    }
  }

  function makeChip(emp, label, dimmed, half) {
    const span = document.createElement("span");
    span.className = "chip" + (half ? " half" : "") + (dimmed ? " dimmed" : "");
    span.innerHTML =
      `<i class="dot" style="background:${emp.color}"></i>` +
      `${escapeHtml(emp.name)}${half ? "(" + label + ")" : ""}`;
    return span;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // ---------- 請假申請表單 ----------
  function setupLeaveForm() {
    const dialog = el("leaveDialog");
    const form = el("leaveForm");
    const addBtn = el("addLeaveBtn");
    const status = el("formStatus");

    if (!cfg.SCRIPT_URL) {
      addBtn.disabled = true;
      addBtn.title = "尚未設定 config.js 的 SCRIPT_URL，暫時無法從網頁送出請假申請";
    }

    addBtn.addEventListener("click", () => {
      form.reset();
      status.hidden = true;
      status.className = "form-status";
      const todayStr = toDateInputValue(new Date());
      el("fStart").value = todayStr;
      el("fEnd").value = todayStr;
      dialog.showModal();
      el("fName").focus();
    });

    el("cancelLeaveBtn").addEventListener("click", () => dialog.close());

    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      const name = el("fName").value.trim();
      const start = el("fStart").value;
      const end = el("fEnd").value;
      const type = el("fType").value;
      const note = el("fNote").value.trim();

      if (!name || !start || !end) {
        setStatus(status, "姓名、開始日期、結束日期為必填", "error");
        return;
      }
      if (end < start) {
        setStatus(status, "結束日期不能早於開始日期", "error");
        return;
      }

      const submitBtn = el("submitLeaveBtn");
      submitBtn.disabled = true;
      setStatus(status, "送出中…", "");

      try {
        const result = await submitLeaveWithRetry({ name, start, end, type, note }, status);
        if (!result.ok) throw new Error(result.error || "送出失敗");
        setStatus(status, "已送出！月曆更新中…", "success");
        await loadFromSheet();
        render();
        setTimeout(() => dialog.close(), 900);
      } catch (err) {
        setStatus(status, "送出失敗：" + err.message + "（若一直失敗，稍等半分鐘再試一次）", "error");
      } finally {
        submitBtn.disabled = false;
      }
    });
  }

  function setStatus(el, text, kind) {
    el.hidden = false;
    el.textContent = text;
    el.className = "form-status" + (kind ? " " + kind : "");
  }

  function toDateInputValue(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Apps Script 網頁應用程式閒置一段時間後，第一次呼叫常需要 5~20 秒「冷啟動」，
  // 有時還會回傳暫時性的錯誤（如 404）。這裡自動重試幾次，避免使用者誤以為系統壞了。
  async function submitLeaveWithRetry(payload, statusEl) {
    const delays = [0, 3000, 6000];
    let lastErr;
    for (let i = 0; i < delays.length; i++) {
      if (i > 0) {
        setStatus(statusEl, `送出中…第一次連線較慢，重試中（${i + 1}/${delays.length}）`, "");
        await sleep(delays[i]);
      }
      try {
        return await submitLeave(payload);
      } catch (err) {
        lastErr = err;
      }
    }
    throw lastErr;
  }

  async function submitLeave(payload) {
    const res = await fetch(cfg.SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }

  // ---------- download ----------
  function downloadCSV() {
    const daysInMonth = new Date(state.year, state.month, 0).getDate();
    const rows = [["姓名", "日期", "假別"]];
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(state.year, state.month - 1, d);
      const dateStr = `${state.year}-${String(state.month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      state.employees.forEach((emp) => {
        const s = statusFor(emp.name, date);
        if (s.full) rows.push([emp.name, dateStr, "全天"]);
        if (s.am) rows.push([emp.name, dateStr, "上午半天"]);
        if (s.pm) rows.push([emp.name, dateStr, "下午半天"]);
      });
    }
    const csv = rows.map((r) => r.map(csvEscape).join(",")).join("\r\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `休假紀錄_${state.year}${String(state.month).padStart(2, "0")}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function csvEscape(v) {
    v = String(v);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }
})();
