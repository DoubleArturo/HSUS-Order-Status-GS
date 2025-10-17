// /**
//  * 這是一個簡易觸發器 (Simple Trigger)，會在使用者編輯試算表時自動執行。
//  * 它的功能是檢查編輯是否發生在目標工作表的標題列，如果是，才呼叫真正的同步函式。
//  * @param {Object} e - 由 Google Sheets 自動傳入的事件物件。
//  */
// function onEditHeader(e) {
//   // --- 設定區 ---
//   const mainSheetName = "Order Shipping Mgt. Table"; // 1. 主要工作表名稱
//   const headerRows = 2; // 2. 要監聽的標題列數 (前2列)
//   // --- 設定結束 ---

//   // 從事件物件 e 中取得被編輯的範圍、工作表、以及列數
//   const range = e.range;
//   const editedSheet = range.getSheet();
//   const editedRow = range.getRow();

//   // 判斷條件：如果編輯的不是目標工作表，或者編輯的列數大於我們設定的標題列數，就直接結束
//   if (editedSheet.getName() !== mainSheetName || editedRow > headerRows) {
//     return;
//   }

//   // 如果通過所有檢查，才執行核心的同步函式
//   syncHeaders();
//   SpreadsheetApp.getActiveSpreadsheet().toast('標題已同步！'); // 顯示一個短暫的成功訊息
// }


// /**
//  * 核心同步函式：複製主表的標題到快取表。
//  */
// function syncHeaders() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const mainSheetName = "Order Shipping Mgt. Table";
//   const cacheSheetName = "Order Shipping Mgt. Table | User Edits";

//   const mainSheet = ss.getSheetByName(mainSheetName);
//   const cacheSheet = ss.getSheetByName(cacheSheetName);

//   if (!mainSheet || !cacheSheet) return;

//   const headerRows = 2; // 標題列數
//   const headerCols = mainSheet.getLastColumn();

//   // 取得主表標題列的值 (純文字)
//   const headerValues = mainSheet.getRange(1, 1, headerRows, headerCols).getDisplayValues();

//   // 將值寫入緩存表（不帶公式）
//   cacheSheet.getRange(1, 1, headerRows, headerCols).setValues(headerValues);
// }