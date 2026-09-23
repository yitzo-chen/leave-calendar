/**
 * 本機模擬 Apps Script 後端，供部署到正式 Google 帳號前先在 localhost 預覽/測試。
 * 用法：node local-dev-server.js [port，預設 8766]（需 xlsx(SheetJS) 於 NODE_PATH，同 xlsx/test/ 下的測試）
 *
 * 行為：從一份「全新試算表」開始（只預填「基本資料」＝跟你正式試算表現在的狀態一致：
 * 工程名稱＝洲際液化天然氣接收站，其餘業主/合約金額/開工日期/公司名稱都還是空的），
 * 啟動時跑一次真正的 setupSheets()，讓它自己遷移出「案場設定」（只會長出 lng 這一列，
 * 案場資料都還沒填，跟你說的「先留空位、以後再填」一致）與其他分頁 —— 這跟之後部署到你
 * 正式帳號時，第一次跑 setupSheets() 會發生的事完全一樣。之後所有 GET/POST 都跑真正的
 * google-apps-script.gs / xlsx-drive.gs 邏輯（同一份程式碼，不是簡化版）。
 *
 * 只是本機預覽用：狀態只存在記憶體，關掉伺服器就消失；不會動到你正式的 Google 帳號。
 * 前端不用改設定，開發時用瀏覽器主控台把 fetch 導到這個伺服器即可（見下方輸出的操作指令）。
 */
const http = require("http");
const fs = require("fs");
const { makeGlueEnv } = require("./xlsx/test/fake-gas.js");

const TEMPLATE_SRC = "D:/洲際液化天然氣接收站AI/日報範本.xlsx";
const TEMPLATE_BYTES = fs.readFileSync(TEMPLATE_SRC);

// 全新試算表：只有「基本資料」預先填好，其餘分頁都還沒建立 —— 跟你正式試算表現在的狀態一致
const INITIAL_SHEETS = {
  "基本資料": [["業主", ""], ["工程名稱", "洲際液化天然氣接收站"], ["合約金額（元）", ""], ["開工日期（YYYY/MM/DD）", ""], ["公司名稱", ""]],
};

const env = makeGlueEnv({ sheetsData: INITIAL_SHEETS, driveFiles: { "日報範本.xlsx": TEMPLATE_BYTES } });
const data = env.state.sheetsData, cache = {};
const maxRows = {};

function sheet(n) {
  return {
    getMaxRows: () => (n in maxRows ? maxRows[n] : 1000),
    insertRowsAfter: (after, cnt) => { maxRows[n] = (n in maxRows ? maxRows[n] : 1000) + cnt; },
    getLastRow: () => { let l = 0; data[n].forEach((r, i) => { if (r.some((v) => v !== "" && v != null)) l = i + 1; }); return l; },
    getDataRange: () => ({ getValues: () => JSON.parse(JSON.stringify(data[n])) }),
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        getValue() { const row = data[n][r - 1] || []; return row[c - 1] === undefined ? "" : row[c - 1]; },
        setValue(v) { while (data[n].length < r) data[n].push([]); data[n][r - 1][c - 1] = v; },
        getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = data[n][r - 1 + i] || []; const x = []; for (let j = 0; j < nc; j++) x.push(row[c - 1 + j] === undefined ? "" : row[c - 1 + j]); o.push(x); } return o; },
        setValues(vs) { if (r + nr - 1 > (n in maxRows ? maxRows[n] : 1000)) throw new Error("範圍超出分頁大小"); for (let i = 0; i < nr; i++) { while (data[n].length < r + i) data[n].push([]); for (let j = 0; j < nc; j++) data[n][r - 1 + i][c - 1 + j] = vs[i][j]; } },
        setDataValidation() { return this; },
        setNumberFormat() { return this; },
      };
    },
    deleteRow: (r) => { data[n].splice(r - 1, 1); },
    appendRow: (arr) => { data[n].push(arr.slice()); },
  };
}
env.ctx.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getId: () => "SS1",
    getSheetByName: (n) => (data[n] ? sheet(n) : null),
    insertSheet: (n) => { data[n] = []; return sheet(n); },
  }),
  newDataValidation: () => ({ requireValueInList: () => ({ build: () => ({}) }) }),
};
env.ctx.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
env.ctx.CacheService = { getScriptCache: () => ({ get: (k) => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
env.ctx.ContentService = { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: "JSON" } };

env.call("setupSheets()");
console.log("[local-dev-server] setupSheets() 完成，分頁：", Object.keys(data).join("、"));
console.log("[local-dev-server] 案場設定：", JSON.stringify(data["案場設定"]));

const PORT = Number(process.argv[2]) || 8766;
http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "*");
  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }
  const url = new URL(req.url, "http://localhost");
  if (req.method === "GET" && url.pathname === "/__dump") {
    // 純本機除錯用：把目前整份「假試算表」的所有分頁內容印出來，不是正式 doGet 的一部分
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.end(JSON.stringify(data, null, 1));
    return;
  }
  if (req.method === "GET") {
    const params = {};
    url.searchParams.forEach((v, k) => { params[k] = v; });
    const out = env.call("doGet")({ parameter: params });
    res.setHeader("Content-Type", "application/json");
    res.end(out.s);
    return;
  }
  if (req.method === "POST") {
    let body = "";
    req.on("data", (c) => { body += c; });
    req.on("end", () => {
      const out = env.call("doPost")({ postData: { contents: body } });
      res.setHeader("Content-Type", "application/json");
      res.end(out.s);
    });
    return;
  }
  res.writeHead(404); res.end();
}).listen(PORT, () => {
  console.log("[local-dev-server] 監聽 http://localhost:" + PORT);
  console.log("[local-dev-server] 前端(另開 http.server)開著時，在瀏覽器主控台貼這段就能連上本機後端：");
  console.log('  const f=window.fetch; window.fetch=(u,o)=>f("http://localhost:' + PORT + '"+new URL(u,location.href).search,o);');
});
