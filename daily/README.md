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

## 工項完成量、主任／工安姓名選單（2026-11-07，尚未部署時只存在本機原始碼）

- **日報頁**（`index.html`）：日期列新增「📋 工項完成量」連結，帶著目前的案場與日期切換到 `items.html`；主任、工安欄新增「填過的名字」下拉清單與一鍵點選（資料來自日報頭的歷史，依最近使用排序，新名字送出後自動出現）。
- **工項完成量頁**（`items.html`＋`items.js`＋`items.css`＋`items-blocks.js`）：項目與預估數量＝試算表「合約工項」分頁（照日報原表 DPR 的 32 列，預估數量為目前合約總量，累計超過只提醒）；每列有「🧮 換算」小工具（車次、鋼板樁支數、工天×產能、澆置區塊…，偏好只存瀏覽器）。送出（需通關密碼）後寫入「工項完成量」分頁。
- **雲端日報彙整 xlsx**：送出後後端只在「施工日報彙整_<案場>.xlsx」新增／更新「工項.民國.月.日」分頁（欄位對應 DPR：預估數量、前日完成、本日完成、累計），供文書人員對照填入清理版。
- **新分頁**：`合約工項`（綠色，可人工修改；序號是鍵，建立後勿改）、`工項完成量`（灰色，網頁寫入勿手改）。`setupSheets()` 只在分頁不存在／空白時建立並放入 lng 的 32 筆種子資料，不動任何既有資料。
- **資料保護（不得覆蓋既有日報彙整）**：①`xlsxUpdateItems_` 只替換「工項.」分頁，原樣保留所有日報分頁，組檔後若有日報分頁遺失即中止不寫入；②第一次寫入前自動在同資料夾存一份「…_工項功能上線前備份.xlsx」（每案場一次）；③範本樣式與雲端檔不同時拒絕寫入；④工項資料只存在資料庫，全量重建 xlsx（「更新全部日報」）時工項分頁會由資料庫重新產生，不會消失；⑤送出只動被明確列出的序號（`items`／`clearItems`），沒列出的既有資料不動。測試：`xlsx/test/items.test.js`（28 項，含日報分頁 XML 逐字比對、日報資料表內容不變）。
- **部署順序（先備份）**：①複製一份試算表（Drive）與目前的 xlsx 彙整檔當備份 ②把 `google-apps-script.gs`、`xlsx/xlsx-builder.js`、`xlsx/xlsx-drive.gs` 貼進 Apps Script 並存檔 ③執行 `setupSheets()`（只會新增兩個分頁）④部署新版本 ⑤確認 GET `workItems`／`leaderNames` 正常後，才 push 前端（GitHub Pages）。前端先上、後端還沒上時，新頁面會顯示「後端尚未更新」，日報頁其餘功能不受影響。
- 本機測試：`node local-dev-server.js 8777`（假後端，跑真正的 .gs 邏輯）＋把前端複製到暫存資料夾、`config.js` 的 `SCRIPT_URL` 指向 `http://localhost:8777`；`/__xlsx` 可下載假 Drive 內的 xlsx 用 Excel 開啟檢查。**不要直接改 repo 內的 `config.js` 來測試。**

## 日報 xlsx：本工人員欄 8 列（2026-11-10 已部署：Apps Script 第21版、Drive 範本已換、更新全部日報13個分頁；尚未 commit）

- 範本人員區「本工」上午／下午由各 6 列改為各 8 列；「預計明日施工項目」由 11 列縮為 7 列（22~28），第 29 列起為人員區標題、上午 30~37、下午 38~45，請款表（46~52）與所有列高不變。
- 範本轉換腳本：`D:\AI\日報系統\port_template_本工8列.py`；builder 的 `CAP.crew=8`、`CAP.tomorrow=7`、`ROWS.crewAm=30`、`ROWS.crewPm=38`。
- **部署順序**（範本與程式必須同步，否則會寫壞檔）：①貼 `xlsx-builder.js` 進 Apps Script ②把新範本 `日報範本.xlsx` 覆蓋 Drive 同名檔（檔案 ID 不變，作法見 `D:\AI\日報系統\進度紀錄.md` 10/07 13:50）③選單「更新全部日報」。回復：舊範本備份在 `D:\AI\日報系統\備份\日報範本_本工8列前_20261110.xlsx`。
- 網頁列印版（`buildPeopleRows`）是另一套版面（每時段至少 6 列、不夠自動加），這次沒動。

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
