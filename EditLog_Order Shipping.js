/**
 * 處理 "Order Shipping Mgt. Table" 工作表的編輯紀錄。
 * @param {Object} e - 從 onEdit 總管函式傳入的事件物件。
 */
function logOrderShippingEdit(e) {
  // --- 設定區 ---
  const TARGET_SHEET_NAME = "Order Shipping Mgt. Table";
  const LOG_SHEET_NAME = "Order Shipping|Edit History";
  const HEADER_ROW = 2; // 請確認此表的標題列數是否為 2
  const ID_COLUMN = 2; // B欄 = P/O
  const START_COL = 1; // 開始監聽的欄位 (A欄)
  const END_COL = 29;  // 結束監聽的欄位 (AC欄)
  // --- 設定結束 ---

  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  if (sheet.getName() !== TARGET_SHEET_NAME || row <= HEADER_ROW || col < START_COL || col > END_COL) {
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