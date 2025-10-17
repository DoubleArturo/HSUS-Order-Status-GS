/**
 * @fileoverview Backend server-side script for the BOL Entry Tool.
 * Handles creating/updating multiple BOL records and managing shipment fulfillment status.
 * [VERSION 9 - Split Caching Optimization]
 */

// --- 常數定義區 ---
const BOL_SHEET_NAME = 'BOL_DB';
const PLANNING_SHEET_NAME = 'Shipment_Planning_DB';
// [V9 修改] 使用分離的快取鍵名
const CACHE_KEY_PENDING = 'pendingBolData';
const CACHE_KEY_FULFILLED = 'fulfilledBolData';


/**
 * Opens the sidebar interface for the BOL Entry Tool.
 */
function openBolEntryTool() {
  const html = HtmlService.createTemplateFromFile('BolEntryTool')
    .evaluate()
    .setTitle('BOL Entry Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * [已修改 V9] 獲取初始資料，採用分離式快取以提升效能。
 * @returns {object} An object containing both pending and fulfilled lists.
 */
function getInitialBolData() {
  Logger.log('--- Starting getInitialBolData() ---'); // 檢查點 A: 函式開始執行

  try {
    const cache = CacheService.getScriptCache();
    
    // 檢查點 B: 檢查快取
    const cachedPending = cache.get(CACHE_KEY_PENDING);
    const cachedFulfilled = cache.get(CACHE_KEY_FULFILLED);
    Logger.log('CHECKPOINT B: Cache checked. Pending status: ' + (cachedPending != null) + ', Fulfilled status: ' + (cachedFulfilled != null));

    // 如果兩個快取都存在，直接從快取回傳資料
    if (cachedPending != null && cachedFulfilled != null) {
      Logger.log('CHECKPOINT C: Returning data from cache.');
      return {
        success: true,
        pendingList: JSON.parse(cachedPending),
        fulfilledList: JSON.parse(cachedFulfilled)
      };
    }
    
    // 檢查點 D: 讀取試算表資料
    Logger.log('CHECKPOINT D: Cache miss. Reading from spreadsheet.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
    if (!planningSheet) throw new Error(`Sheet '${PLANNING_SHEET_NAME}' not found.`);
    
    const lastRow = planningSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('INFO: Sheet is empty, returning empty lists.');
      return { success: true, pendingList: [], fulfilledList: [] };
    }

    // 仍使用 getRange('A2:G' + lastRow) 以確保讀取完整資料範圍
    const dataRange = planningSheet.getRange('A2:G' + lastRow);
    const planningData = dataRange.getValues();
    
    const pendingList = [];
    const fulfilledList = [];

    // 檢查點 E: 處理資料
    Logger.log('CHECKPOINT E: Processing ' + planningData.length + ' rows.');
    
    planningData.forEach(row => {
      const timestamp = row[0]; // Column A
      const key = row[2];       // Column C
      const status = row[6];    // Column G

      if (key) {
        if (status === 'Fulfilled') {
          fulfilledList.push({ key: key, timestamp: timestamp });
        } else {
          pendingList.push(key);
        }
      }
    });
    
    fulfilledList.sort((a, b) => {
      const dateA = a.timestamp ? new Date(a.timestamp) : new Date(0);
      const dateB = b.timestamp ? new Date(b.timestamp) : new Date(0);
      return dateB - dateA;
    });

    const uniquePendingList = [...new Set(pendingList)].sort();

    // 檢查點 F: 寫入快取
    Logger.log('CHECKPOINT F: Writing results back to cache.');
    cache.put(CACHE_KEY_PENDING, JSON.stringify(uniquePendingList), 300);
    cache.put(CACHE_KEY_FULFILLED, JSON.stringify(fulfilledList), 300);

    Logger.log('CHECKPOINT G: Successfully returning fresh data.');
    return { 
      success: true, 
      pendingList: uniquePendingList,
      fulfilledList: fulfilledList
    };
  } catch (e) {
    // 檢查點 H: 錯誤捕獲
    Logger.log(`ERROR H: getInitialBolData Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}


/**
 * [已修改 V8] 獲取已存在的 BOL 數據，並修正欄位索引錯誤。
 */
function getExistingBolData(poSkuKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bolSheet = ss.getSheetByName(BOL_SHEET_NAME);
    if (!bolSheet) throw new Error(`Sheet '${BOL_SHEET_NAME}' not found.`);
    const lastRow = bolSheet.getLastRow();
    if (lastRow < 2) return { success: true, bols: [], actShipDate: null, isFulfilled: false };
    
    const data = bolSheet.getRange('A2:F' + lastRow).getValues();
    const existingBols = [];
    let actShipDate = null;

    data.forEach(row => {
      if (row[1] === poSkuKey) { // Column B is poSkuKey
        if (!actShipDate && row[4] instanceof Date) {
          actShipDate = Utilities.formatDate(row[4], Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        existingBols.push({
          bolNumber: row[0],
          shippedQty: row[2],
          shippingFee: row[3],
          signed: row[5]
        });
      }
    });

    const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
    const planningData = planningSheet.getRange('C2:G' + planningSheet.getLastRow()).getValues();
    // [修正] 讀取範圍從 C 欄開始，所以 key 在索引 0，status 在索引 4
    const isFulfilled = planningData.some(row => row[0] === poSkuKey && row[4] === 'Fulfilled'); 

    return { success: true, bols: existingBols, actShipDate: actShipDate, isFulfilled: isFulfilled };
  } catch (e) {
    Logger.log(`getExistingBolData Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * [已修改 V9] 儲存 BOL 數據，並在成功後清除相關的分離式快取。
 * @param {object} data The data object from the frontend form.
 * @returns {object} A result object.
 */
function saveBolData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bolSheet = ss.getSheetByName(BOL_SHEET_NAME);
    if (!bolSheet) throw new Error(`Sheet '${BOL_SHEET_NAME}' not found`);
    
    const poSkuKey = data.poSkuKey;
    if (!poSkuKey) throw new Error("PO|SKU Key is missing.");
    
    const dataRange = bolSheet.getDataRange();
    const allData = dataRange.getValues();
    const rowsToDelete = [];
    allData.forEach((row, index) => {
      if (index > 0 && row[1] === poSkuKey) {
        rowsToDelete.push(index + 1);
      }
    });
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      bolSheet.deleteRow(rowsToDelete[i]);
    }

    const actShipDate = new Date(data.actShipDate);
    const newStatus = data.isFulfilled ? 'Fulfilled' : '';
    const timestamp = new Date();
    const newRows = [];

    data.bols.forEach(bol => {
      const shippedQty = parseInt(bol.shippedQty, 10);
      const shippingFee = parseFloat(bol.shippingFee);
      if (bol.bolNumber && shippedQty > 0) {
        newRows.push([
          bol.bolNumber, poSkuKey, shippedQty,
          isNaN(shippingFee) ? 0 : shippingFee,
          actShipDate, bol.signed,
          newStatus, timestamp
        ]);
      }
    });

    if (newRows.length > 0) {
      bolSheet.getRange(bolSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    updateFulfillmentStatus(poSkuKey, newStatus, timestamp);
    
    SpreadsheetApp.flush();

    // [V9 修改] 成功儲存後，清除相關的快取
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_PENDING);
    cache.remove(CACHE_KEY_FULFILLED);

    return { success: true, message: `BOL records saved successfully for '${poSkuKey}'!` };
  } catch (e) {
    Logger.log(`saveBolData Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 更新 Planning 表的狀態 (G欄) 和時間戳記 (A欄)。
 * @param {string} keyToUpdate The PO|SKU key to find and update.
 * @param {string} status The new status to set.
 * @param {Date} timestamp The timestamp of the change.
 */
function updateFulfillmentStatus(keyToUpdate, status, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
  if (!planningSheet) return;
  
  const lastRow = planningSheet.getLastRow();
  if (lastRow < 2) return;
  
  const keys = planningSheet.getRange('C2:C' + lastRow).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === keyToUpdate) {
      const targetRow = i + 2;
      planningSheet.getRange(targetRow, 1).setValue(timestamp); 
      planningSheet.getRange(targetRow, 7).setValue(status);
      break; 
    }
  }
}
