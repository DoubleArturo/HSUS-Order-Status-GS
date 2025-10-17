/**
 * 歸檔並刪除指定 PO 資料的【安全版本 v3 - 支援 W 欄 ARRAYFORMULA】。
 * 此版本會保護標題列，並在寫回資料時清空指定由 ARRAYFORMULA 控制的欄位，
 * 確保公式可以正常向下擴展。
 */
function archiveProcessedPOs_Safe() {
  const sourceSheetName = 'Dealer PO | Raw Data';
  const archiveSheetName = 'Dealer PO | Archive';
  const statusColumnIndex = 25; // Col Y
  const headerRows = 1; // 您的標題只有 1 列

  // --- 【關鍵修正：將 W 欄 (23) 加入 ARRAYFORMULA 保護名單】---
  const arrayFormulaColumns = [5, 17, 23]; // E, Q, W 欄

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const archiveSheet = ss.getSheetByName(archiveSheetName);

  if (!sourceSheet || !archiveSheet) {
    Logger.log('錯誤：找不到來源或歸檔工作表。');
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow <= headerRows) {
    Logger.log('來源工作表中沒有需要處理的資料。');
    return;
  }

  // 1. 分開讀取標題和資料
  const headers = sourceSheet.getRange(1, 1, headerRows, lastCol).getValues();
  const dataRange = sourceSheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol);
  const sourceDataRows = dataRange.getValues();

  const dataToArchive = [];
  const keptRows = [];

  // 2. 處理資料分類
  for (let i = 0; i < sourceDataRows.length; i++) {
    const row = sourceDataRows[i];
    const status = row[statusColumnIndex - 1]; 

    if (status === 'Change' || status === 'Voided' || status === 'Revised') {
      dataToArchive.push(row);
    } else {
      keptRows.push(row);
    }
  }

  if (dataToArchive.length > 0) {
    const archiveLastRow = archiveSheet.getLastRow();
    if (archiveLastRow === 0 && headers.length > 0) {
      archiveSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
    }
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);

    // 在寫回資料前，遍歷所有要保留的資料列 (keptRows)
    // 並將所有 ARRAYFORMULA 欄位的值設為空字串
    const arrayFormulaIndices = arrayFormulaColumns.map(col => col - 1); // 轉換為 0-based 索引

    keptRows.forEach(row => {
      arrayFormulaIndices.forEach(index => {
        if (index < row.length) { // 確保索引存在
          row[index] = ''; // 將該欄位清空
        }
      });
    });

    // 只清除資料區域，不觸碰標題列
    dataRange.clearContent();
    
    // 如果還有需要保留的資料，則把它們寫回去
    if (keptRows.length > 0) {
      sourceSheet.getRange(headerRows + 1, 1, keptRows.length, keptRows[0].length).setValues(keptRows);
    }
    
    Logger.log(`成功歸檔 ${dataToArchive.length} 筆資料，保留 ${keptRows.length} 筆資料。`);
  } else {
    Logger.log('沒有找到需要歸檔的資料。');
  }
}