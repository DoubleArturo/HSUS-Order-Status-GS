/**
 * 重新計算並應用 P/O 分組的交替背景色。
 * 此函式將由時間觸發器或手動選單呼叫。
 */
function recolorPoGroups() {
  // --- 設定區 ---
  const sheetName = "Order Shipping Mgt. Table";
  const dataStartRow = 4; // 資料開始的列數
  const poCol = 2;        // PO# 所在的欄位 (B欄)
  const color1 = "#FFFFFF"; // 顏色一 (白色)
  const color2 = "#F2F2F2"; // 顏色二 (淡灰色)
  // --- 設定結束 ---

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    console.error(`工作表 "${sheetName}" 不存在，腳本已停止。`);
    SpreadsheetApp.getUi().alert(`錯誤：找不到名為 "${sheetName}" 的工作表！`);
    return;
  }

  const lastRow = sheet.getLastRow();
  
  // 如果資料列數少於起始列，先清除舊顏色後再停止
  if (lastRow < dataStartRow) {
    console.log(`工作表中沒有足夠的資料列可供著色，將清除舊格式。`);
    const rangeToClear = sheet.getRange(dataStartRow, 1, sheet.getMaxRows() - dataStartRow + 1, sheet.getMaxColumns());
    rangeToClear.clearFormat();
    return;
  }
  
  // 清除舊的背景色，確保每次都是重新計算
  sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getMaxColumns()).clearFormat();

  const range = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getMaxColumns());
  const values = range.getValues();
  const backgrounds = [];

  let lastPO = null;
  // *** 這裡是唯一的修改處 ***
  // 將初始顏色設為 color1，這樣第一組資料就會被正確地上色為 color2。
  let currentColor = color1; 

  for (let i = 0; i < values.length; i++) {
    const currentPO = values[i][poCol - 1];
    const rowBackgrounds = [];

    // 當 PO# 變更時 (且不是空白)，切換顏色
    if (currentPO !== "" && currentPO !== lastPO) {
      currentColor = (currentColor === color1) ? color2 : color1;
      lastPO = currentPO;
    }
    
    const finalColor = (currentPO === "") ? color1 : currentColor;

    for (let j = 0; j < sheet.getMaxColumns(); j++) {
      rowBackgrounds.push(finalColor);
    }
    backgrounds.push(rowBackgrounds);
  }

  range.setBackgrounds(backgrounds);
  console.log(`著色完成。總共處理了 ${values.length} 列。`);
}
