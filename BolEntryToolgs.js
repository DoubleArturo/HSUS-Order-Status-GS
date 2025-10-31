/**
 * @fileoverview Backend server-side script for the BOL Entry Tool.
 * [VERSION 12.6 - Model Name Safe]
 * - ⚡️ FEATURE: getInitialBolData() now joins with Price Book to return PO|SKU (Model) format.
 * - ⚡️ FIX: Ensures 100% serialization safety with JSON.stringify().
 */

// --- 常數定義區 ---
const BOL_SHEET_NAME = 'BOL_DB';
const PLANNING_SHEET_NAME = 'Shipment_Planning_DB';
const CACHE_KEY_PENDING = 'pendingBolData';
const CACHE_KEY_FULFILLED = 'fulfilledBolData';

// --- 🚀 [V12.6 新增] 常數 (用於 Model 查詢) ---
const PRICE_BOOK_SHEET_NAME_BOL = 'New HSUS Order Status - HSUS Price Book(QBO)'; 
const PB_SKU_COL = 7;
const PB_MODEL_NAME_COL = 22;


/**
 * Opens the sidebar interface for the BOL Entry Tool.
 */
function openBolEntryTool() {
  const html = HtmlService.createTemplateFromFile('BolEntryTool')
    .evaluate()
    .setTitle('BOL Entry Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- 🚀 [V12.6 新增] 輔助函式：獲取 SKU 到 Model Name 的對照表 ---
/**
 * Reads the HSUS Price Book(QBO) sheet to create a map of SKU# to Model Name.
 * @returns {Map<string, string>} A map where key is SKU# and value is Model Name.
 */
function getSkuModelMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priceBookSheet = ss.getSheetByName(PRICE_BOOK_SHEET_NAME_BOL); 
  if (!priceBookSheet) {
    Logger.log("Price Book sheet not found, proceeding without Model Names.");
    return new Map();
  }
  const lastRow = priceBookSheet.getLastRow();
  if (lastRow < 2) return new Map();

  const data = priceBookSheet.getRange(2, PB_SKU_COL, lastRow - 1, PB_MODEL_NAME_COL - PB_SKU_COL + 1).getValues();
  const skuModelMap = new Map();
  const modelNameRelativeIndex = PB_MODEL_NAME_COL - PB_SKU_COL;

  data.forEach(row => {
    const sku = String(row[0]).trim(); // SKU# (相對索引 0)
    const modelName = String(row[modelNameRelativeIndex]).trim(); // Sales Description 
    
    if (sku) {
      // 淨化 Model Name
      const cleanModelName = modelName.replace(/Finished Goods:|^450\w+|Standard/g, '').trim().replace(/_/g, ' ');
      skuModelMap.set(sku, cleanModelName || sku); // 預設回傳 SKU
    }
  });
  return skuModelMap;
}


/**
 * [已修改 V12.6] 獲取初始資料，包含 Model Name 並保證 100% 序列化安全。
 * @returns {string} 一個 JSON 字串，包含 { success: boolean, pendingList: object[], fulfilledList: object[] }
 */
function getInitialBolData() {
  try {
    Logger.log('--- Starting getInitialBolData() [V12.6 - Model Name Safe] ---'); 
    
    const cache = CacheService.getScriptCache();
    const cachedPending = cache.get(CACHE_KEY_PENDING);
    const cachedFulfilled = cache.get(CACHE_KEY_FULFILLED);

    if (cachedPending != null && cachedFulfilled != null) {
      Logger.log('[CHECKPOINT C] Returning RAW STRING data from cache.');
      const payload = {
        success: true,
        pendingList: JSON.parse(cachedPending),
        fulfilledList: JSON.parse(cachedFulfilled)
      };
      return JSON.stringify(payload); // 再次序列化為字串並回傳
    }
    
    Logger.log('[CHECKPOINT D] Cache miss. Reading data...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
    if (!planningSheet) throw new Error(`Sheet '${PLANNING_SHEET_NAME}' not found.`);
    
    // --- 🚀 [V12.6 關鍵修改] 獲取 SKU -> Model 對照表 ---
    const skuModelMap = getSkuModelMap();
    Logger.log('[CHECKPOINT D.1] SKU Model Map created.');

    const lastRow = planningSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('INFO: Sheet is empty, returning empty lists.');
      cache.put(CACHE_KEY_PENDING, '[]', 300);
      cache.put(CACHE_KEY_FULFILLED, '[]', 300);
      return JSON.stringify({ success: true, pendingList: [], fulfilledList: [] });
    }

    const dataRange = planningSheet.getRange('A2:G' + lastRow);
    const planningData = dataRange.getValues();
    
    const pendingMap = new Map(); // 用 Map 處理 Pending 的去重
    const fulfilledList = []; // This will hold {key, display, timestamp}

    Logger.log('[CHECKPOINT E] Processing ' + planningData.length + ' rows.');
    
    planningData.forEach(row => {
      const timestamp = row[0]; // Column A
      const key = (row[2] === null || row[2] === undefined) ? '' : String(row[2]); // Column C
      const status = row[6];    // Column G

      if (key) { 
        // --- 🚀 [V12.6 關鍵修改] 產生 display 名稱 ---
        const sku = key.split('|')[1] || '';
        const modelName = skuModelMap.get(sku) || sku; // 找不到 Model 時使用 SKU
        const display = `${key} (${modelName})`;
        // --- [修改結束] ---

        if (status === 'Fulfilled') {
          const validTimestampString = (timestamp instanceof Date && !isNaN(timestamp)) 
                ? timestamp.toISOString() 
                : new Date(0).toISOString(); 
          fulfilledList.push({ key: key, display: display, timestamp: validTimestampString });
        } else {
          // 如果 Map 中沒有這個 key，才新增 (去重)
          if (!pendingMap.has(key)) {
            pendingMap.set(key, { key: key, display: display });
          }
        }
      }
    });
    
    Logger.log('[CHECKPOINT E.1] Sorting lists...');
    fulfilledList.sort((a, b) => b.timestamp.localeCompare(a.timestamp)); // 依時間戳降序
    const uniquePendingList = [...pendingMap.values()].sort((a, b) => a.display.localeCompare(b.display)); // 依顯示名稱升序
    
    // 移除 timestamp，只保留 key 和 display
    const serializableFulfilledList = fulfilledList.map(item => ({ key: item.key, display: item.display }));

    Logger.log('[CHECKPOINT F] Writing results back to cache.');
    cache.put(CACHE_KEY_PENDING, JSON.stringify(uniquePendingList), 300);
    cache.put(CACHE_KEY_FULFILLED, JSON.stringify(serializableFulfilledList), 300);

    const payload = { 
      success: true, 
      pendingList: uniquePendingList,
      fulfilledList: serializableFulfilledList
    };

    Logger.log('[CHECKPOINT G] Payload constructed. Forcing serialization NOW...');
    const stringPayload = JSON.stringify(payload);
    Logger.log('[CHECKPOINT G.1] Serialization successful. Returning safe string.');
    return stringPayload; 

  } catch (e) {
    Logger.log(`[ERROR H] getInitialBolData FAILED: ${e.message}\n${e.stack}`);
    return JSON.stringify({ success: false, message: `getInitialBolData Error: ${e.message}` }); 
  }
}

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
