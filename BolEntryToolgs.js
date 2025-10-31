/**
 * @fileoverview Backend server-side script for the BOL Entry Tool.
 * [VERSION 10.1 - SERIALIZATION FIX]
 * - Implements Google Query Language for high-speed data retrieval (方案 3).
 * - Fixes faulty cache logic (方案 2).
 * - Includes manual cache clearing tool (方案 1).
 * - ⚡️ FIX: Removes non-serializable gviz Date object from the 'fulfilledList' return
 * payload, which was causing google.script.run to fail silently.
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
 * [已修改 V10.1] 獲取初始資料，修正序列化問題。
 * @returns {object} An object containing both pending and fulfilled lists.
 */
function getInitialBolData() {
  Logger.log('--- Starting getInitialBolData() [V10.1 - Serialization Fix] ---'); 

  try {
    const cache = CacheService.getScriptCache();
    const cachedPending = cache.get(CACHE_KEY_PENDING);
    const cachedFulfilled = cache.get(CACHE_KEY_FULFILLED);
    Logger.log('CHECKPOINT B: Cache checked. Pending: ' + (cachedPending != null) + ', Fulfilled: ' + (cachedFulfilled != null));

    if (cachedPending != null && cachedPending !== "[]" && cachedFulfilled != null && cachedFulfilled !== "[]") {
      Logger.log('CHECKPOINT C: Returning data from cache.');
      return {
        success: true,
        pendingList: JSON.parse(cachedPending),
        fulfilledList: JSON.parse(cachedFulfilled)
      };
    }
    
    Logger.log('CHECKPOINT D: Cache miss. Running high-speed Query.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME);
    if (!planningSheet) throw new Error(`Sheet '${PLANNING_SHEET_NAME}' not found.`);

    // 查詢 1: 獲取所有 "未完成" (G 欄 != 'Fulfilled' 或為空) 的、唯一的 PO_SKU_Key (C 欄)，並排序
    const pendingQuery = "SELECT C WHERE C IS NOT NULL AND (G != 'Fulfilled' OR G IS NULL) GROUP BY C ORDER BY C ASC";
    const pendingRows = _runQuery(planningSheet, pendingQuery);
    const uniquePendingList = pendingRows.map(row => row[0]); // [key1, key2]

    // 查詢 2: 獲取所有 "已完成" (G 欄 = 'Fulfilled') 的 PO_SKU_Key (C 欄) 和最新的時間戳 (A 欄)
    const fulfilledQuery = "SELECT C, MAX(A) WHERE C IS NOT NULL AND G = 'Fulfilled' GROUP BY C ORDER BY MAX(A) DESC";
    const fulfilledRows = _runQuery(planningSheet, fulfilledQuery);
    
    // --- 🚀 [V10.1 關鍵修正] ---
    // 我們只回傳 `key`。
    // `row[1]` (時間戳) 是一個 gviz Date 物件，它會導致 google.script.run 序列化失敗。
    // 前端 HTML 只需要 `item.key`，排序已由 Query 完成。
    const fulfilledList = fulfilledRows.map(row => {
      return { 
        key: row[0] // PO_SKU_Key
        // 我們刻意不回傳 row[1] (時間戳)
      }; 
    });
    // --- [修正結束] ---

    Logger.log('CHECKPOINT F: Writing query results back to cache.');
    cache.put(CACHE_KEY_PENDING, JSON.stringify(uniquePendingList), 300); // 快取 5 分鐘
    cache.put(CACHE_KEY_FULFILLED, JSON.stringify(fulfilledList), 300);

    Logger.log('CHECKPOINT G: Successfully returning fresh data from Query.');
    return { 
      success: true, 
      pendingList: uniquePendingList,
      fulfilledList: fulfilledList
    };
  } catch (e) {
    Logger.log(`ERROR H: getInitialBolData Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * [方案 3 輔助函式] 執行 Google Visualization API 查詢。
 * (此函數保持不變)
 */
function _runQuery(sheet, query) {
  const sheetId = sheet.getParent().getId();
  const sheetGid = sheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${sheetGid}&tq=${encodeURIComponent(query)}`;
  
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  
  const text = response.getContentText();
  const jsonText = text.replace(/google.visualization.Query.setResponse\((.*)\);/, '$1');
  const json = JSON.parse(jsonText);
  
  if (json.status === 'error') {
    throw new Error(`Query failed: ${json.errors.map(e => e.detailed_message).join(', ')}`);
  }
  
  return json.table.rows.map(row => {
    return row.c.map(cell => (cell ? cell.v : null)); // .v 是原始值
  });
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

    return { success: true, message: `BOL records saved successfully for '${poSkuKey}'!` };
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
 * [方案 1 - 手動執行] 強制清除 BOL Entry Tool 的快取
 * (此函數保持不變，供您使用)
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
