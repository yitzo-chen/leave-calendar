# 施工日報填寫系統

試行對象：lng案（洲際液化天然氣接收站，利用尚未開工的部分安全測試，不影響 SS-01/SS-02 正式資料）。
架構跟 `../leave/` 一樣：Google 試算表當資料庫、Apps Script 當後端、靜態網頁當前端。

**目前狀態：前端、後端（Apps Script 第 11 版）、xlsx 雲端存檔皆已完成並上線，網址為 `…/leave-calendar/daily/`。**

## 在本倉庫中的位置與維護方式

日報原始碼已從 `D:\daily-report` 併入本倉庫的 `daily/`（保留原本的 commit 歷史），這裡是**唯一一份**原始碼，不要再另存副本。

- 網頁（會被員工看到）：`index.html`、`app.js`、`style.css`、`config.js`
- 後端（要手動貼進 Apps Script 編輯器並部署，不會因為 push 而生效）：`google-apps-script.gs`、`xlsx/xlsx-builder.js`（Apps Script 內檔名 `xlsx-builder.gs`）、`xlsx/xlsx-drive.gs`
- 測試：`xlsx/test/`，在該資料夾內執行，例如 `node validation.test.js`（需要 `jszip`、`xlsx` 套件與範本 `日報範本.xlsx`）
- 因為 GitHub Pages 會提供倉庫內所有檔案，後端與測試檔也可從網址讀到；內容不含密碼（密碼存在 Apps Script 的指令碼屬性）。
- 修改前端後，把 `index.html` 內 `?v=` 的版本號遞增，避免瀏覽器快取舊檔。

## 資料放在哪裡

已經在 Google Drive「日報網頁填寫資料庫」資料夾建立試算表：
https://docs.google.com/spreadsheets/d/1PaWIt-FQPy3bfNhoNQ_Jc66Vdfi8UoR4SDZy2_9S8OI/edit

`config.js` 的 `SHEET_ID` 已經預填好，不用再改。

## 設定步驟

1. 打開上面那份試算表
2. 選單「擴充功能」→「Apps Script」→ 把 `google-apps-script.gs` 全部內容貼上覆蓋 `Code.gs` → 存檔
3. 函式選單選 `setupSheets` → 按「執行」→ 第一次會跳 Google 授權畫面，允許即可
   - 這會自動建立 4 個分頁：`基本資料`、`工種機具材料清單`、`日報記錄`、`日報頭`
   - `工種機具材料清單` 會自動帶入現行 Excel 範本的預設項目（工種8項/機具6項/材料9項），`啟用中`欄位控制要不要顯示在表單，`工項編號`欄位先留空，之後接成本管理系統再填
   - 之後要新增/停用項目，直接在這個分頁改，不用重跑腳本
4. 右上角「部署」→「新增部署作業」→ 齒輪選「網頁應用程式」
   - 執行身分：**我**
   - 具有存取權的使用者：**只有我**（維持私人測試，之後要開放現場填寫再改成「所有人」）
5. 部署完成拿到的網址貼到 `config.js` 的 `SCRIPT_URL`

## 幫送出加密碼（跟休假月曆同一套機制，建議先設定再測試）

Apps Script 編輯器 → 左側齒輪「專案設定」→「指令碼屬性」→ 新增屬性 `SUBMIT_PASSWORD`，值填你要的密碼。
沒設定的話送出不檢查密碼（測試期間可以先不設，正式使用前記得補上）。

## 後端 API（目前只能用 curl 測試，前端還沒寫）

- `GET ?action=config` — 讀取啟用中的工種/機具/材料清單 + 基本資料
- `GET ?action=day&date=2026-09-18` — 讀取某天已存在的日報頭+明細（給表單預填用）
- `GET ?action=cumulative&upTo=2026-09-18` — 讀取每個項目到某天為止的累計工天數
- `POST` — 送出/修正一天的日報（同一個日期再送一次會覆蓋舊資料，upsert）

```json
{
  "date": "2026-09-18",
  "password": "...",
  "clientId": "任意隨機字串，同一次送出重試要帶同一個",
  "header": {
    "weather": "晴", "status": "施工", "todayWork": "...", "tomorrowPlan": "...",
    "overtimeAM": "", "overtimePM": "", "remark": "", "reporter": "陳易佐"
  },
  "items": [
    { "name": "公司工", "am": 5, "pm": 5 },
    { "name": "挖土機(50型)", "am": 1, "pm": 1 }
  ]
}
```

## 還沒做的部分

- 前端網頁（`index.html`/`app.js`/`style.css`）——手機優先、分區塊表單、送出前預填當天已有資料
- 假單輸出那種正式格式列印功能
- 跟休假月曆合併成主入口網頁（使用者未來規劃，這次不處理）
