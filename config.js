// 請把 Google 試算表網址中 /d/ 與 /edit 中間那一串英數字貼到下面
// 例如網址是 https://docs.google.com/spreadsheets/d/1AbC-xyz123/edit
// 那 SHEET_ID 就是 "1AbC-xyz123"
// 留空時網頁會顯示示範資料。
//
// SCRIPT_URL：照 README「開放網頁直接填假」步驟部署 Apps Script 後拿到的網址
// （長得像 https://script.google.com/macros/s/xxxx/exec）。留空的話「我要請假」
// 按鈕會停用，畫面仍可正常瀏覽/列印/下載，只是不能從網頁送出新的請假申請。
window.LEAVE_CONFIG = {
  SHEET_ID: "1418hcwCosUFt-4uWfYEpRcK1pDFFt_-Tvg2_0Ny6Wk8",
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbyPqlZ0GahYHzLKXKdfgSJcNbEVtP41mZ5ZeOSEVjORvqCce35XgkHE84GJeapucquc/exec",
  ROSTER_SHEET_NAME: "員工名單",
  LOG_SHEET_NAME: "請假紀錄",
  COMPANY_NAME: "員工休假月曆",
};
