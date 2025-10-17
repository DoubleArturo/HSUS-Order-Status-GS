// /**
//  * onEdit 總管函式。
//  * 根據被編輯的工作表名稱，呼叫對應的處理函式。
//  * @param {Object} e - 由 Google Sheets 自動傳入的事件物件。
//  */
// function onEditLog(e) {
//   try {
//     if (!e || !e.range) return;
//     const sheetName = e.range.getSheet().getName();

//     // 根據工作表名稱，決定要執行哪個腳本
//     if (sheetName === "Order Shipping Mgt. Table") {
//       logOrderShippingEdit(e); // 呼叫處理第一個表的函式
//     } else if (sheetName === "Operation | Pending Order Dashboard") {
//       logOperationDashboardEdit(e); // 呼叫處理第二個表的函式
//     }

//   } catch (err) {
//     console.error(`onEdit Error: ${err.message}\nStack: ${err.stack}`);
//   }
// }