/**
 * 處理 "Operation | Pending Order Dashboard" 工作表的編輯紀錄。
 * @param {Object} e - 從 onEdit 總管函式傳入的事件物件。
 */
function logOperationDashboardEdit(e) {
  // --- 設定區 ---
  const LOG_SHEET_NAME = "Operation Dashboard | Edit History";
  const HEADER_ROW = 2;       // 標題所在的列數，請根據您的表格調整
  const ID_COLUMN = 8;        // 用來識別資料的欄位 (H欄 = P/O)，請根據您的表格調整
  // 要監聽的特定欄位編號列表
  const COLS_TO_WATCH = [1, 2, 3, 28, 30, 32]; // A,B,C, AB,AD,AF
  // --- 設定結束 ---

  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // 如果編輯的不是標題列之後，或不在我們監聽的欄位列表中，則結束
  if (row <= HEADER_ROW || !COLS_TO_WATCH.includes(col)) {
    return;
  }

  const oldValue = e.oldValue !== undefined ? String(e.oldValue) : "";
  const newValue = e.value !== undefined ? String(e.value) : "";
  if (oldValue === newValue) return;

  const timestamp = new Date();
  const user = e.user.getEmail();
  const identifier = sheet.getRange(row, ID_COLUMN).getValue();
  const fieldName = sheet.getRange(HEADER_ROW, col).getValue();
  const action = (oldValue === "") ? "Created" : "Updated";

  // 在主表上更新 AH:AJ 欄位
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const createdTimeCol = headers.indexOf("Created Time") + 1;
  const updatedTimeCol = headers.indexOf("Last Updated Time") + 1;
  const historyCol = headers.indexOf("Latest Update Record") + 1;
  const timezone = e.source.getSpreadsheetTimeZone();

  if (createdTimeCol > 0) {
    const createdTimeCell = sheet.getRange(row, createdTimeCol);
    if (createdTimeCell.getValue() === "") {
      createdTimeCell.setValue(timestamp).setNumberFormat("yyyy/m/d AM/PM hh:mm:ss");
    }
  }
  if (updatedTimeCol > 0) {
    sheet.getRange(row, updatedTimeCol).setValue(timestamp).setNumberFormat("yyyy/m/d");
  }
  if (historyCol > 0) {
    const formattedDate = Utilities.formatDate(timestamp, timezone, "yyyy/MM/dd");
    const historyString = `${formattedDate} ${action} "${fieldName}"`;
    sheet.getRange(row, historyCol).setValue(historyString);
  }

  // 將詳細紀錄寫入 Log 工作表
  const ss = e.source;
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    const logHeaders = ["Timestamp", "P/O", "Row", "Field", "Action", "New Value", "Old Value", "User"];
    logSheet.appendRow(logHeaders);
    logSheet.getRange("A1:H1").setFontWeight("bold");
    logSheet.setColumnWidth(1, 160);
  }
  logSheet.appendRow([timestamp, identifier, row, fieldName, action, newValue, oldValue, user]);
}