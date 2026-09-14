# 員工休假月曆

公司內部休假總覽網頁：即時讀取 Google 試算表資料，顯示月曆檢視、可列印、可下載當月 CSV。
**填假本身還是在 Google 試算表操作**，這個網頁只負責把資料變好看、方便列印/下載。

## 檔案結構
- `index.html` — 頁面結構
- `style.css` — 樣式（含深色模式、列印樣式）
- `app.js` — 讀取試算表資料、月曆運算、列印/下載功能
- `config.js` — **唯一需要手動編輯的檔案**

## 設定步驟

1. 把「員工休假管理表.xlsx」上傳到 Google 雲端硬碟，用 Google 試算表開啟
2. 右上角「共用」→「一般存取權」改成「知道連結的使用者」→ 權限選「檢視者」以上即可
   （員工要能直接編輯的話選「編輯者」；這個網頁本身只讀取，不會寫入）
3. 複製網址列的試算表 ID：`https://docs.google.com/spreadsheets/d/【這一串】/edit`
4. 打開 `config.js`，把 ID 貼到 `SHEET_ID: ""` 中間
5. 確認 `ROSTER_SHEET_NAME` / `LOG_SHEET_NAME` 跟試算表分頁名稱一致（預設「員工名單」「請假紀錄」）

## 發布到 GitHub Pages

```bash
git init
git add .
git commit -m "初版：員工休假月曆網頁"
git branch -M main
git remote add origin https://github.com/<your-account>/leave-calendar.git
git push -u origin main
```

推上去之後，到 GitHub 該 repo 的 Settings → Pages，Source 選 `main` branch、`/ (root)`，
存檔後幾分鐘網址就會是 `https://<your-account>.github.io/leave-calendar/`。

## 注意事項
- 這是**獨立的新 repo**，跟你其他私人 repo 完全無關，不會互相影響
- 這個 repo 一旦啟用 GitHub Pages 就是公開網頁，任何知道網址的人都能看到員工姓名與休假日期，
  請自行評估是否要放公司內部敏感資訊（例如請假事由備註），必要時可以把「備註」欄位排除在網頁顯示外
- 若試算表分頁被改名、或分享權限被改回「限制」，網頁會抓不到資料並顯示錯誤訊息、自動改顯示示範資料
