/**
 * SheetService.js
 * * 抽象化資料存取層。所有與 Google Sheets 交互的讀寫操作都應通過此服務進行。
 * 核心功能：
 * 1. 根據欄位名稱 (Header Name) 查找欄位索引 (Column Index)。
 * 2. 獲取數據時，將每行數據轉換為 Key-Value Object 格式 (例如: { 'P/O': '12345', 'QTY': 2 })。
 * 3. 優化讀取性能 (使用 getValues() 一次性讀取)。
 */

/**
 * 獲取工作表及其標頭 (Headers) 的快取版本。
 * @param {string} sheetName - 要獲取的工作表名稱 (從 Config.SHEET_NAMES 獲取)。
 * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: string[]}} - 包含工作表物件和標頭陣列的物件。
 */
function getSheetAndHeaders(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(Config.DOCUMENT_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // 錯誤處理：如果工作表不存在，拋出錯誤
    throw new Error(`SheetService Error: Cannot find sheet named "${sheetName}". Please check Config.js.`);
  }

  // 假設標頭總是在第一行 (Row 1)
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  return { sheet, headers };
}

/**
 * 從指定工作表中讀取所有資料，並轉換為物件陣列。
 * (解決問題 1：雖然仍在讀取 Sheets，但我們使用更高效的 getValues() 一次性讀取。)
 * @param {string} sheetName - 要讀取的工作表名稱 (e.g., Config.SHEET_NAMES.DEALER_PO_RAW)。
 * @param {number} headerRowIndex - 標頭所在的行數 (預設為 1)。
 * @returns {Array<Object>} - 包含每個記錄物件的陣列。
 */
function readAllRecords(sheetName, headerRowIndex = 1) {
  try {
    const { sheet, headers } = getSheetAndHeaders(sheetName);
    const lastRow = sheet.getLastRow();

    if (lastRow <= headerRowIndex) {
      Logger.log(`Sheet "${sheetName}" is empty or only contains headers.`);
      return []; // 如果沒有資料，返回空陣列
    }

    // 從標頭行下一行開始讀取到最後一行
    const dataRange = sheet.getRange(headerRowIndex + 1, 1, lastRow - headerRowIndex, headers.length);
    const values = dataRange.getValues();
    const records = [];

    // 將二維陣列的每一行 (Row) 轉換為以標頭 (Header) 為鍵的物件 (Object)
    values.forEach((row, rowIndex) => {
      const record = {};
      row.forEach((value, colIndex) => {
        const header = headers[colIndex];
        if (header) {
          // 清除標頭的空格和換行符，確保鍵值一致
          const cleanHeader = header.trim().replace(/\n/g, ' ').replace(/\r/g, '');
          record[cleanHeader] = value;
        }
      });
      // 在物件中添加一個隱藏的欄位，用於記錄此行在 Sheet 中的實際行數，
      // 對於後續的寫入和更新非常重要！
      record._rowNumber = rowIndex + headerRowIndex + 1;
      records.push(record);
    });

    return records;
  } catch (e) {
    Logger.log(`Failed to read records from ${sheetName}: ${e.message}`);
    return [];
  }
}

/**
 * 根據特定的鍵和值，在工作表中找到第一個匹配的記錄並返回。
 * @param {string} sheetName - 工作表名稱。
 * @param {string} keyName - 要搜尋的欄位名稱 (e.g., 'P/O')。
 * @param {*} keyValue - 要匹配的欄位值。
 * @returns {Object|null} - 匹配的記錄物件，或 null。
 */
function findRecordByKey(sheetName, keyName, keyValue) {
  const records = readAllRecords(sheetName);
  // 使用 Array.find() 進行記憶體內搜尋，效率優於 Sheets 內的單元格循環。
  return records.find(record => record[keyName] === keyValue) || null;
}


/**
 * 根據記錄的 _rowNumber 更新 Sheets 中的單行資料。
 * @param {string} sheetName - 工作表名稱。
 * @param {Object} record - 包含要更新值的物件，必須包含 _rowNumber。
 * @returns {boolean} - 更新是否成功。
 */
function updateRecord(sheetName, record) {
  if (!record._rowNumber) {
    Logger.log("Update Error: Record must contain the internal '_rowNumber' key.");
    return false;
  }

  try {
    const { sheet, headers } = getSheetAndHeaders(sheetName);
    const rowNumber = record._rowNumber;
    const valuesToUpdate = [];

    // 依據標頭順序，構建要寫入的單行陣列
    headers.forEach(header => {
      const cleanHeader = header.trim().replace(/\n/g, ' ').replace(/\r/g, '');

      // 獲取物件中匹配標頭的值。如果記錄物件中沒有這個鍵，則使用空值或保留原值 (這裡我們選擇更新欄位有提供的資料)
      // 為了安全起見，這裡應該只更新 record 內明確提供的欄位。
      // 對於更新，我們必須拉取原行資料，否則沒有提供的欄位會被設為 undefined。
      // 由於 App Script 沒有內建合併功能，我們使用更安全的 "只寫入提供的資料" 邏輯。

      // 1. 獲取原始行數據
      const originalRowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
      const newRowValues = [...originalRowValues]; // 複製原始值

      // 2. 遍歷標頭，用新值覆蓋舊值
      headers.forEach((header, index) => {
        const cleanHeader = header.trim().replace(/\n/g, ' ').replace(/\r/g, '');
        if (record.hasOwnProperty(cleanHeader)) {
          newRowValues[index] = record[cleanHeader];
        }
      });

      // 3. 寫入單行資料
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([newRowValues]);
    });

    return true;
  } catch (e) {
    Logger.log(`Failed to update record in ${sheetName} at row ${record._rowNumber}: ${e.message}`);
    return false;
  }
}


/**
 * 將單一記錄 (Key-Value Object) 附加到工作表底部。
 * @param {string} sheetName - 工作表名稱。
 * @param {Object} record - 要寫入的記錄物件 (e.g., { 'P/O': '12345', 'QTY': 2 })。
 * @returns {boolean} - 寫入是否成功。
 */
function appendRecord(sheetName, record) {
  try {
    const { sheet, headers } = getSheetAndHeaders(sheetName);
    const newRowValues = [];

    // 依據標頭順序，構建要寫入的單行陣列
    headers.forEach(header => {
      const cleanHeader = header.trim().replace(/\n/g, ' ').replace(/\r/g, '');
      // 根據標頭名稱查找 record 物件中對應的值
      newRowValues.push(record.hasOwnProperty(cleanHeader) ? record[cleanHeader] : '');
    });

    sheet.appendRow(newRowValues);
    return true;
  } catch (e) {
    Logger.log(`Failed to append record to ${sheetName}: ${e.message}`);
    return false;
  }
}


// --- 範例使用 (可刪除或保留作為測試) ---

/**
 * 演示如何使用 SheetService 讀取和更新 Dealer PO Raw Data。
 */
function demoSheetServiceUsage() {
  const PO_RAW_SHEET = Config.SHEET_NAMES.DEALER_PO_RAW;
  const PO_KEY = Config.PRIMARY_KEYS.PO_NUMBER;
  const targetPO = '735-2578'; // 假設一個 PO 號碼

  Logger.log(`[Demo] 1. Reading all records from: ${PO_RAW_SHEET}`);
  const allRecords = readAllRecords(PO_RAW_SHEET);
  Logger.log(`Total records read: ${allRecords.length}`);

  Logger.log(`[Demo] 2. Finding record by PO #: ${targetPO}`);
  const targetRecord = findRecordByKey(PO_RAW_SHEET, PO_KEY, targetPO);

  if (targetRecord) {
    Logger.log(`Found record at Sheet Row: ${targetRecord._rowNumber}`);
    Logger.log(`Original P/O Total: ${targetRecord['P/O - Total']}`);

    // 範例：更新記錄中的兩個欄位 (使用欄位名稱)
    const updateData = {
      _rowNumber: targetRecord._rowNumber, // 必須包含這個內部鍵
      'P/O - Total': 9999.00, // 使用欄位名稱
      'Note': 'Updated by SheetService demo at ' + new Date().toISOString() // 使用欄位名稱
    };

    Logger.log('[Demo] 3. Attempting to update record...');
    const success = updateRecord(PO_RAW_SHEET, updateData);

    if (success) {
      Logger.log('[Demo] Update successful. Check your Google Sheet.');
    } else {
      Logger.log('[Demo] Update failed.');
    }

  } else {
    Logger.log(`Record with PO ${targetPO} not found.`);
  }
}
