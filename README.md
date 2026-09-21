# 工地服務入口

洲際液化天然氣接收站的內部網頁入口，主頁有兩個入口：

| 入口 | 路徑 | 說明 |
|---|---|---|
| 員工請假 | `leave/` | 休假月曆、請假申請、假單輸出（原本的 `leave-calendar`）。設定方式見 [leave/README.md](leave/README.md) |
| 日報填寫 | `daily/` | 每日施工日報填寫、列印，資料寫入 Google 試算表並同步產生 xlsx |

## 檔案結構
- `index.html`、`style.css` — 主頁（左右分割的兩個入口，手機上下對半）
- `leave/` — 休假月曆（`index.html`、`app.js`、`style.css`、`config.js`、`google-apps-script.gs`）
- `daily/` — 日報填寫（`index.html`、`app.js`、`style.css`、`config.js`）

## 注意
- 各頁只用相對路徑，內頁的「‹ 主頁」連結為 `../`。
- `daily/` 只放網頁本身；日報的 Apps Script 後端原始碼與測試在本機 `D:\daily-report`（尚無遠端倉庫），
  修改日報前端時需要把該處的 `index.html`、`app.js`、`style.css`、`config.js` 同步複製到這裡的 `daily/`。
- 舊網址 `…/leave-calendar/` 現在是主頁；月曆移到 `…/leave-calendar/leave/`。
