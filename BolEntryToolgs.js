/**
 * @fileoverview Backend server-side script for the BOL Entry Tool.
 * [VERSION 10.5 - STABLE & SERIALIZATION-SAFE]
 * - ⚡️ REVERT: This version reverts to the stable 'getValues()' method as requested.
 * (This is the method used before the 'Query' version).
 * - ⚡️ FIX: This version also PERMANENTLY fixes the "Stuck on Loading" issue
 * by converting all 'Date' objects to ISO strings *before* sorting and returning.
 * This guarantees the payload is 100% safe for google.script.run.
 */

// --- 常數定義區 ---
const BOL_SHEET_NAME = 'BOL_DB';
const PLANNING_SHEET_NAME = 'Shipment_Planning_DB';
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
 * [已修改 V10.7] 終極序列化修復。
 * - ⚡️ FIX: 在 try...catch 內部強制 JSON.stringify()，
 * 並回傳一個 100% 安全的字串，以規避 google.script.run 的序列化器問題。
 */
function getInitialBolData() {
  try {
    Logger.log('--- Starting getInitialBolData() [V10.7 - Forced Stringify] ---'); 
    
    const cache = CacheService.getScriptCache();
    // [V10.7 修正] 我們不再從快取中回傳 JSON 物件，
    // 而是回傳快取中儲存的「字串」。
    const cachedPending = cache.get(CACHE_KEY_PENDING);
    const cachedFulfilled = cache.get(CACHE_KEY_FULFILLED);

    if (cachedPending != null && cachedFulfilled != null) {
      Logger.log('[CHECKPOINT C] Returning RAW STRING data from cache.');
      const payload = {
        success: true,
        pendingList: JSON.parse(cachedPending), // 先解析以符合結構
        fulfilledList: JSON.parse(cachedFulfilled)
      };
      // 再次序列化為字串並回傳
      return JSON.stringify(payload);
    }
    
    Logger.log('[CHECKPOINT D] Cache miss. Reading from spreadsheet using getValues().');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
    if (!planningSheet) throw new Error(`Sheet '${PLANNING_SHEET_NAME}' not found.`);
    
    const lastRow = planningSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('INFO: Sheet is empty, returning empty lists.');
      cache.put(CACHE_KEY_PENDING, '[]', 300);
      cache.put(CACHE_KEY_FULFILLED, '[]', 300);
      // 確保回傳結構一致（一個已序列化的字串）
      return JSON.stringify({ success: true, pendingList: [], fulfilledList: [] });
    }

    const dataRange = planningSheet.getRange('A2:G' + lastRow);
    const planningData = dataRange.getValues();
    
    const pendingList = [];
    const fulfilledList = []; 

    Logger.log('[CHECKPOINT E] Processing ' + planningData.length + ' rows.');
    
    planningData.forEach(row => {
      const timestamp = row[0]; // Column A
      // [V10.7 修正] 確保 key 絕對是字串，即使是數字 0
      const key = (row[2] === null || row[2] === undefined) ? '' : String(row[2]); // Column C
      const status = row[6];    // Column G

      if (key) { // 確保 key 不是空字串
        if (status === 'Fulfilled') {
          // [V10.6 邏輯保留]
          const validTimestampString = (timestamp instanceof Date && !isNaN(timestamp)) 
                ? timestamp.toISOString() 
                : new Date(0).toISOString(); 
          fulfilledList.push({ key: key, timestamp: validTimestampString });
        } else {
          pendingList.push(key);
        }
      }
    });
    
    Logger.log('[CHECKPOINT E.1] Sorting lists...');
    fulfilledList.sort((a, b) => b.timestamp.localeCompare(a.timestamp)); 
    
    const uniquePendingList = [...new Set(pendingList)].sort();
    const serializableFulfilledList = fulfilledList.map(item => ({ key: item.key }));

    Logger.log('[CHECKPOINT F] Writing results back to cache.');
    // 快取儲存的仍然是字串化的陣列
    cache.put(CACHE_KEY_PENDING, JSON.stringify(uniquePendingList), 300);
    cache.put(CACHE_KEY_FULFILLED, JSON.stringify(serializableFulfilledList), 300);

    const payload = { 
      success: true, 
      pendingList: uniquePendingList,
      fulfilledList: serializableFulfilledList
    };

    // --- 🚀 [V10.7 關鍵修正] ---
    // 不直接回傳物件，而是回傳序列化後的「字串」。
    // 這將強制在 try...catch 內執行序列化。
    Logger.log('[CHECKPOINT G] Payload constructed. Forcing serialization NOW...');
    const stringPayload = JSON.stringify(payload);
    Logger.log('[CHECKPOINT G.1] Serialization successful. Returning safe string.');
    return stringPayload; 
    // --- [修正結束] ---

  } catch (e) {
    // [V10.7 關鍵修正] 
    // 如果 V10.7 的 JSON.stringify(payload) 失敗，錯誤「必定」會在這裡被捕獲！
    Logger.log(`[ERROR H] getInitialBolData FAILED (Serialization Error?): ${e.message}\n${e.stack}`);
    // 回傳一個字串化的錯誤物件
    return JSON.stringify({ success: false, message: `getInitialBolData Error: ${e.message}` }); 
  }
}

// ... (getExistingBolData, saveBolData, updateFulfillmentStatus 函數保持不變) ...

/**
 * 獲取已存在的 BOL 數據。
 * (此函數保持不變)
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
      if (row[1] === poSkuKey) { 
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
    const isFulfilled = planningData.some(row => row[0] === poSkuKey && row[4] === 'Fulfilled'); 

    return { success: true, bols: existingBols, actShipDate: actShipDate, isFulfilled: isFulfilled };
  } catch (e) {
    Logger.log(`getExistingBolData Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 儲存 BOL 數據。
 * (此函數保持不變)
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

    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_PENDING);
    cache.remove(CACHE_KEY_FULFILLED);

    return { success: true, message: `Successfully saved for '${poSkuKey}'.` };
  } catch (e) {
    Logger.log(`saveBolData Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 更新 Planning 表的狀態。
 * (此函數保持不變)
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

/**
 * [方案 1 - 手動執行] 
 * 強制清除 BOL Entry Tool 的快取
 */
function clearBolCache_Manual() {
  const cache = CacheService.getScriptCache();
  cache.remove('pendingBolData');
  cache.remove('fulfilledBolData');
  SpreadsheetApp.getUi().alert(
    'BOL Tool Cache Cleared!', 
    'The cache for the BOL Entry Tool has been successfully cleared. Please close and re-open the tool.', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  Logger.log('BOL Tool cache (pendingBolData, fulfilledBolData) has been manually cleared.');
}
