// /**
//  * 此函式為所有需要處理的原始數據工作表（Raw Data sheets）
//  * 呼叫 assignUuidToNewRows 函式。
//  * 您可以將此函式設定為一個定時或 On change 的觸發器。
//  */
// function processAllRawDataSheets() {
//   // 處理 "Dealer PO | Raw Data" 工作表，在 W 欄插入 UUID。
//   // W 是第 23 欄。同時檢查 D 欄 (第 4 欄) 是否有資料。
//   assignUuidToNewRows('Dealer PO | Raw Data', 23, 4);

//   // 處理 "Direct Quote | Raw Data" 工作表，在 V 欄插入 UUID。
//   // V 是第 22 欄。同時檢查 F 欄 (第 6 欄) 是否有資料。
//   assignUuidToNewRows('Direct Quote | Raw Data', 22, 6);
// }

// /**
//  * 此函式用於檢查指定工作表中的資料列，
//  * 如果指定的 UUID 欄位為空，且另一指定欄位有資料，
//  * 則自動產生並插入一個新的 UUID。
//  *
//  * @param {string} sheetName - 要處理的工作表名稱。
//  * @param {number} uuidColumnIndex - UUID 欄位的索引號 (A=1, B=2, C=3...)。
//  * @param {number} conditionColumnIndex - 判斷資料是否為空的欄位索引號。
//  */
// function assignUuidToNewRows(sheetName, uuidColumnIndex, conditionColumnIndex) {
//   // 取得目前作用中的試算表。
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   // 根據名稱取得指定的工作表。
//   const sheet = ss.getSheetByName(sheetName);

//   // 如果找不到工作表，則拋出錯誤。
//   if (!sheet) {
//     Logger.log(`錯誤: 找不到工作表名稱為 "${sheetName}"。`);
//     return;
//   }

//   // 取得工作表的最後一列。
//   const lastRow = sheet.getLastRow();
//   // 從第二列開始處理，以避免標頭。
//   const startRow = 2;

//   // 檢查工作表是否沒有資料。
//   if (lastRow < startRow) {
//     Logger.log(`"${sheetName}" 工作表中沒有新資料可處理。`);
//     return;
//   }

//   // 取得所有資料。
//   const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
//   const values = dataRange.getValues();

//   // 遍歷每一列資料。
//   for (let i = 0; i < values.length; i++) {
//     const row = values[i];
//     // 取得 UUID 欄位的值。
//     const uuidValue = row[uuidColumnIndex - 1]; // 陣列索引是從 0 開始，所以要減 1。
//     // 取得判斷資料是否存在的欄位值。
//     const conditionValue = row[conditionColumnIndex - 1];

//     // 如果 UUID 欄位為空，且判斷欄位有資料，則生成並寫入新的 UUID。
//     if ((!uuidValue || String(uuidValue).trim() === '') && (conditionValue && String(conditionValue).trim() !== '')) {
//       // 取得要寫入 UUID 的單元格範圍。
//       const cell = sheet.getRange(i + startRow, uuidColumnIndex);
//       // 生成一個新的 UUID，包含時間戳記。
//       const newUuid = generateTimestampedId();
//       // 將新的 UUID 寫入單元格。
//       cell.setValue(newUuid);
//       Logger.log(`在新資料列 ${i + startRow} 中插入帶有時間戳記的 ID: ${newUuid}`);
//     }
//   }
//   Logger.log(`完成為 "${sheetName}" 工作表中的新資料列插入帶有時間戳記的 ID。`);
// }

// /**
//  * 產生一個包含時間戳記和隨機字串的唯一 ID。
//  * 格式為：[Unix時間戳記]_[隨機字串]。
//  * @returns {string} 帶有時間戳記的唯一 ID。
//  */
// function generateTimestampedId() {
//   const timestamp = new Date().getTime(); // 取得當前的毫秒時間戳記。
//   const randomPart = Utilities.getUuid().substring(0, 8); // 取得 UUID 的前 8 個字元作為隨機字串。
//   return `${timestamp}_${randomPart}`;
// }
