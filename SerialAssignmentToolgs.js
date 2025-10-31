/**
 * @fileoverview Backend server-side script for the Serial Assignment Tool.
 * [VERSION 12.5 - Model Name Safe]
 * - ⚡️ FEATURE: getPoSkuLists() now joins with Price Book to return PO|SKU (Model) format.
 * - ⚡️ FIX: Ensures 100% serialization safety with JSON.stringify().
 */

// --- 🎯 常數定義區 ---
const BOL_DB_SHEET_NAME = 'BOL_DB';
const SERIAL_RAW_DATA_SHEET_NAME = 'Serial # | Raw Data';
const SERIAL_DB_SHEET_NAME = 'Serial #_DB'; 
const ORDER_MGT_SHEET_NAME = 'Order Shipping Mgt. Table';
const PRICE_BOOK_SHEET_NAME_Serial = 'New HSUS Order Status - HSUS Price Book(QBO)'; 

// --- Column Definitions ---
// Serial # | Raw Data
const RAW_SKU_COL = 2;        // B: SKU
const RAW_SERIAL_COL = 4;     // D: Serial #
const RAW_POSKU_KEY_COL = 8;  // H: PO_SKU_Key (Helper)
const RAW_INBOUND_COL = 13;   // M: Inbound Date

// Serial #_DB
const DB_SERIAL_COL = 1;      // A: Serial #
const DB_POSKU_KEY_COL = 2;   // B: PO_SKU_Key
const DB_BOL_COL = 3;         // C: BOL #
const DB_COMPLETE_COL = 4;    // D: Complete
const DB_USER_COL = 5;        // E: Assigned User
const DB_TIMESTAMP_COL = 6;   // F: Assigned Timestamp

// HSUS Price Book(QBO)
const PB_SKU_COL = 7;
const PB_MODEL_NAME_COL = 22;

const COMPLETE_STATUS_TEXT = 'Complete Assigned';

/**
 * Opens the Serial Assignment Tool sidebar.
 */
function openSerialAssignmentTool() {
  const html = HtmlService.createTemplateFromFile('SerialAssignmentTool')
    .evaluate()
    .setTitle('Serial Assignment Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- 輔助函式：獲取 SKU 到 Model Name 的對照表 ---

/**
 * Reads the HSUS Price Book(QBO) sheet to create a map of SKU# to Model Name.
 * @returns {Map<string, string>} A map where key is SKU# and value is Model Name.
 */
function getSkuModelMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priceBookSheet = ss.getSheetByName(PRICE_BOOK_SHEET_NAME_Serial); 
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
      const cleanModelName = modelName.replace(/Finished Goods:|^450\w+|Standard/g, '').trim().replace(/_/g, ' ');
      skuModelMap.set(sku, cleanModelName || sku); // 預設回傳 SKU
    }
  });

  return skuModelMap;
}


// --- 核心函式：獲取 PO/SKU 列表 (包含 Model Name 和狀態) ---

/**
 * [V12.5 修復套用]
 * Reads from BOL_DB and Serial #_DB to get pending/finished lists,
 * returning objects with original key and formatted name.
 * @returns {string} 一個 JSON 字串，包含 { pending, finished }
 */
function getPoSkuLists() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const serialDbSheet = ss.getSheetByName(SERIAL_DB_SHEET_NAME);
    const bolSheet = ss.getSheetByName(BOL_DB_SHEET_NAME);
    
    // --- 🚀 [V12.5 關鍵] 獲取 SKU -> Model 對照表 ---
    const skuModelMap = getSkuModelMap(); 

    if (!serialDbSheet || !bolSheet) throw new Error("Required sheet not found: Serial #_DB or BOL_DB");

    // 輔助函式：取得格式化名稱
    const getFormattedKey = (poSkuKey) => {
      const parts = poSkuKey.split('|');
      // 確保 parts[parts.length - 1] 是 SKU
      const sku = parts.length > 1 ? parts[parts.length - 1] : ''; 
      const modelName = skuModelMap.get(sku) || sku; // 找不到 Model 時使用 SKU
      return `${poSkuKey} (${modelName})`;
    };
    
    const dbLastRow = serialDbSheet.getLastRow();
    const poSkuStatusMap = new Map(); // Key: PO_SKU_Key, Value: {timestamp, isComplete}
    
    if (dbLastRow >= 2) {
      const dbData = serialDbSheet.getRange(2, DB_POSKU_KEY_COL, dbLastRow - 1, DB_TIMESTAMP_COL - DB_POSKU_KEY_COL + 1).getValues();
      dbData.forEach(row => {
        const poSkuKey = row[0];
        const status = row[DB_COMPLETE_COL - DB_POSKU_KEY_COL];
        const timestamp = row[DB_TIMESTAMP_COL - DB_POSKU_KEY_COL];
        const isComplete = status === COMPLETE_STATUS_TEXT;
        
        const currentEntry = poSkuStatusMap.get(poSkuKey);
        if (!currentEntry || (timestamp instanceof Date && timestamp > currentEntry.timestamp)) {
             poSkuStatusMap.set(poSkuKey, {timestamp, isComplete});
        }
      });
    }
    
    const bolLastRow = bolSheet.getLastRow();
    const allBolPoSkus = new Set();
    if (bolLastRow >= 2) {
      bolSheet.getRange('B2:B' + bolLastRow).getValues().flat().filter(String).forEach(poSku => allBolPoSkus.add(poSku));
    }
    
    const combinedList = [...allBolPoSkus].map(poSkuKey => {
        const statusEntry = poSkuStatusMap.get(poSkuKey);
        const isComplete = statusEntry ? statusEntry.isComplete : false;
        
        const safeTimestamp = (statusEntry && statusEntry.timestamp instanceof Date && !isNaN(statusEntry.timestamp)) 
                                ? statusEntry.timestamp.toISOString() 
                                : new Date(0).toISOString(); 

        return {
            key: poSkuKey,
            // --- 🚀 [V12.5 關鍵] 使用格式化名稱 ---
            display: getFormattedKey(poSkuKey),
            isComplete: isComplete,
            timestamp: safeTimestamp 
        };
    });
    
    const pending = combinedList
        .filter(item => !item.isComplete)
        .sort((a, b) => a.display.localeCompare(b.display)); // 按 Display Name 字母排序

    const finished = combinedList
        .filter(item => item.isComplete)
        .sort((a, b) => b.timestamp.localeCompare(a.timestamp)); // 按 ISO 字串降序排序

    // --- 🚀 [V10.7 關鍵修正] ---
    // 強制序列化為 100% 安全的字串
    const payload = { pending, finished };
    return JSON.stringify(payload);
    // --- [修正結束] ---

  } catch (e) {
    Logger.log(`getPoSkuLists Error: ${e.message}`);
    // --- 🚀 [V10.7 關鍵修正] ---
    // 錯誤也必須序列化為字串
    return JSON.stringify({ success: false, message: e.message }); 
  }
}


// --- 核心函式：狀態處理 (Check Box觸發) ---

/**
 * 根據傳入的 PO|SKU 鍵和狀態，更新 Serial #_DB 中所有相關記錄的完成狀態。
 * @param {string} poSkuKey 要更新的 PO|SKU 鍵。
 * @param {boolean} isComplete 欲設定的新狀態 (true: Complete Assigned; false: 清空狀態)。
 * @returns {Object} 成功或失敗訊息。
 */
function updateAssignmentCompletionStatus(poSkuKey, isComplete) {
  try {
    if (!poSkuKey) throw new Error("PO|SKU key is required for status update.");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERIAL_DB_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SERIAL_DB_SHEET_NAME}" not found.`);

    const statusText = isComplete ? COMPLETE_STATUS_TEXT : ''; // 決定要寫入的狀態文本
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, message: "No assignments found in DB." };

    const dataRange = sheet.getRange(2, DB_POSKU_KEY_COL, lastRow - 1, DB_TIMESTAMP_COL - DB_POSKU_KEY_COL + 1);
    const values = dataRange.getValues();
    let updatedCount = 0;

    const completeColRelative = DB_COMPLETE_COL - DB_POSKU_KEY_COL;
    const userColRelative = DB_USER_COL - DB_POSKU_KEY_COL;
    const timestampColRelative = DB_TIMESTAMP_COL - DB_POSKU_KEY_COL;
    
    values.forEach(row => {
      if (row[0] === poSkuKey) {
        row[completeColRelative] = statusText;
        if (isComplete) {
          row[userColRelative] = userEmail;
          row[timestampColRelative] = timestamp;
        } else {
          row[userColRelative] = '';
          row[timestampColRelative] = '';
        }
        updatedCount++;
      }
    });

    if (updatedCount > 0) {
      dataRange.setValues(values);
    }
    
    const action = isComplete ? "標記為完成" : "重新開啟";
    return { success: true, message: `"${poSkuKey}" 已成功 ${action}，共更新 ${updatedCount} 筆記錄。` };
  } catch (e) {
    Logger.log(`updateAssignmentCompletionStatus Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}


// --- 原有函式 (保持完整性) ---
// ... (getSerialsForEditing, getAssignedSerialsForPoSku, assignSerials, 等函數保持不變) ...
// (我將折疊這些函數以節省篇幅，但它們與您上一版本中的程式碼相同)

/**
 * MODIFIED: Gets serial numbers suitable for editing an assignment based on SKU and Inbound Date.
 */
function getSerialsForEditing(sku, poSkuKey) {
  try {
    if (!sku || !poSkuKey) return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName(SERIAL_RAW_DATA_SHEET_NAME);
    const dbSheet = ss.getSheetByName(SERIAL_DB_SHEET_NAME);
    if (!rawSheet || !dbSheet) throw new Error("Required sheets not found.");
    const rawLastRow = rawSheet.getLastRow();
    const allSkuSerials = new Set();
    if (rawLastRow >= 2) {
      const rawData = rawSheet.getRange(2, RAW_SKU_COL, rawLastRow - 1, RAW_INBOUND_COL - RAW_SKU_COL + 1).getValues();
      rawData.forEach(row => {
        const rowSku = row[0];
        const serial = row[RAW_SERIAL_COL - RAW_SKU_COL];
        const inboundDate = row[RAW_INBOUND_COL - RAW_SKU_COL];
        if (rowSku === sku && inboundDate) { 
          allSkuSerials.add(serial);
        }
      });
    }
    const dbLastRow = dbSheet.getLastRow();
    const usedSerialsMap = new Map();
    if (dbLastRow >= 2) {
      const dbData = dbSheet.getRange(2, DB_SERIAL_COL, dbLastRow - 1, DB_POSKU_KEY_COL).getValues();
      dbData.forEach(row => {
        usedSerialsMap.set(row[0], row[1]); 
      });
    }
    const availableSerials = [...allSkuSerials].filter(serial => 
      !usedSerialsMap.has(serial) || usedSerialsMap.get(serial) === poSkuKey
    );
    return availableSerials.sort();
  } catch (e) {
    Logger.log(`getSerialsForEditing Error: ${e.message}`);
    // [V12.5 修復] 確保拋出錯誤
    throw new Error(e.message);
  }
}

/**
 * Gets currently assigned serials for a PO|SKU from Serial #_DB.
 */
function getAssignedSerialsForPoSku(poSkuKey) {
  try {
    if (!poSkuKey) return {};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SERIAL_DB_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SERIAL_DB_SHEET_NAME}" not found.`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const data = sheet.getRange(2, DB_SERIAL_COL, lastRow - 1, DB_BOL_COL).getValues();
    const assignments = {};
    data.forEach(row => {
      const serial = row[0];
      const rowPoSkuKey = row[DB_POSKU_KEY_COL - 1];
      const bol = row[DB_BOL_COL - 1];
      if (rowPoSkuKey === poSkuKey && bol && serial) {
        if (!assignments[bol]) {
          assignments[bol] = [];
        }
        assignments[bol].push(serial);
      }
    });
    return assignments;
  } catch (e) {
    Logger.log(`getAssignedSerialsForPoSku Error: ${e.message}`);
    // [V12.5 修復] 確保拋出錯誤
    throw new Error(e.message);
  }
}


/**
 * REWRITTEN: Handles dual-write and preserves "Complete" status when editing.
 */
function assignSerials(assignmentData) {
  try {
    const { poSkuKey, assignments } = assignmentData;
    if (!poSkuKey) throw new Error("No PO|SKU key provided.");

    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName(SERIAL_DB_SHEET_NAME);
    const rawSheet = ss.getSheetByName(SERIAL_RAW_DATA_SHEET_NAME);
    if (!dbSheet || !rawSheet) throw new Error("Required sheets not found.");

    let allSerialsToAssign = new Set(Object.values(assignments).flat());

    const dbLastRow = dbSheet.getLastRow();
    let isAlreadyComplete = false;
    if (dbLastRow >= 2) {
      const dbData = dbSheet.getRange(2, DB_POSKU_KEY_COL, dbLastRow - 1, DB_COMPLETE_COL - DB_POSKU_KEY_COL + 1).getValues();
      for (const row of dbData) {
        if (row[0] === poSkuKey && row[DB_COMPLETE_COL - DB_POSKU_KEY_COL] === COMPLETE_STATUS_TEXT) {
          isAlreadyComplete = true;
          break;
        }
      }
    }
    const rawLastRow = rawSheet.getLastRow();
    if (rawLastRow >= 2) {
      const serialColValues = rawSheet.getRange(2, RAW_SERIAL_COL, rawLastRow - 1, 1).getValues().flat();
      const helperKeyColRange = rawSheet.getRange(2, RAW_POSKU_KEY_COL, rawLastRow - 1, 1);
      const helperKeyValues = helperKeyColRange.getValues();
      const serialToRawIndex = new Map(serialColValues.map((serial, i) => [serial, i]));
      const currentlyAssignedSerials = new Set();
      if (dbLastRow >= 2) {
        const poSkuKeysInDb = dbSheet.getRange(2, DB_POSKU_KEY_COL, dbLastRow - 1, 1).getValues();
        const serialsInDb = dbSheet.getRange(2, DB_SERIAL_COL, dbLastRow - 1, 1).getValues();
        poSkuKeysInDb.forEach((keyArr, index) => {
          if (keyArr[0] === poSkuKey) {
            currentlyAssignedSerials.add(serialsInDb[index][0]);
          }
        });
      }
      currentlyAssignedSerials.forEach(serial => {
        if (!allSerialsToAssign.has(serial) && serialToRawIndex.has(serial)) {
          const rowIndex = serialToRawIndex.get(serial);
          if (helperKeyValues[rowIndex][0] === poSkuKey) {
            helperKeyValues[rowIndex][0] = '';
          }
        }
      });
      allSerialsToAssign.forEach(serial => {
        if (serialToRawIndex.has(serial)) {
          const rowIndex = serialToRawIndex.get(serial);
          helperKeyValues[rowIndex][0] = poSkuKey;
        }
      });
      helperKeyColRange.setValues(helperKeyValues);
    }
    if (dbLastRow >= 2) {
      const poSkuKeys = dbSheet.getRange(2, DB_POSKU_KEY_COL, dbLastRow - 1, 1).getValues();
      const rowsToDelete = [];
      for (let i = poSkuKeys.length - 1; i >= 0; i--) {
        if (poSkuKeys[i][0] === poSkuKey) {
          rowsToDelete.push(i + 2);
        }
      }
      rowsToDelete.forEach(rowNum => dbSheet.deleteRow(rowNum));
    }
    const newDbRows = [];
    const completeStatus = isAlreadyComplete ? COMPLETE_STATUS_TEXT : '';
    for (const bolNumber in assignments) {
      assignments[bolNumber].forEach(serial => {
        newDbRows.push([serial, poSkuKey, bolNumber, completeStatus, userEmail, timestamp]);
      });
    }
    if (newDbRows.length > 0) {
      dbSheet.getRange(dbSheet.getLastRow() + 1, 1, newDbRows.length, newDbRows[0].length).setValues(newDbRows);
    }
    updateOrderMgtSerials(poSkuKey, [...allSerialsToAssign]);
    return { success: true, message: "Serial numbers updated successfully!" };
  } catch (e) {
    Logger.log(`assignSerials Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * Helper function to update the concatenated serials in Order Shipping Mgt. Table.
 */
function updateOrderMgtSerials(poSkuKey, serials) {
    const orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDER_MGT_SHEET_NAME);
    if (!orderSheet) return;
    const lastRow = orderSheet.getLastRow();
    if (lastRow < 2) return;
    const ORDER_KEY_COL = 18; // R
    const SERIALS_WRITE_COL = 28; // AB
    const orderKeys = orderSheet.getRange(2, ORDER_KEY_COL, lastRow - 1, 1).getValues();
    const targetRowIndex = orderKeys.findIndex(row => row[0] === poSkuKey);
    if (targetRowIndex !== -1) {
        const serialsCell = orderSheet.getRange(targetRowIndex + 2, SERIALS_WRITE_COL); 
        serialsCell.setValue(serials.join(', '));
    }
}

function getBolsForPoSku(poSkuKey) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOL_DB_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${BOL_DB_SHEET_NAME}" not found.`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange('A2:C' + lastRow).getValues();
    const results = [];
    data.forEach(row => {
      if (row[1] === poSkuKey) {
        results.push({ bolNumber: row[0], shippedQty: row[2] });
      }
    });
    return results;
  } catch (e) {
    Logger.log(`getBolsForPoSku Error: ${e.message}`);
    // [V12.5 修復] 確保拋出錯誤
    throw new Error(e.message);
  }
}

function getSerialStatus(serialNumber) {
  try {
    if (!serialNumber) return { status: 'Error', message: 'Serial number cannot be empty.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName(SERIAL_RAW_DATA_SHEET_NAME);
    const dbSheet = ss.getSheetByName(SERIAL_DB_SHEET_NAME);
    if (!rawSheet || !dbSheet) throw new Error("Required sheets not found.");
    const rawLastRow = rawSheet.getLastRow();
    let inboundDate = '';
    if (rawLastRow >= 2) {
      const rawData = rawSheet.getRange('D2:M' + rawLastRow).getValues(); 
      const rawRow = rawData.find(r => r[0] === serialNumber);
      if (rawRow) inboundDate = rawRow[9]; 
    }
    if (inboundDate === '') return { status: 'Non-Inbound' };
    const dbLastRow = dbSheet.getLastRow();
    if (dbLastRow >= 2) {
      const dbData = dbSheet.getRange(2, 1, dbLastRow - 1, DB_TIMESTAMP_COL).getValues();
      const dbRow = dbData.find(r => r[0] === serialNumber);
      if (dbRow) {
        const timestamp = dbRow[DB_TIMESTAMP_COL - 1];
        let formattedDate = '';
        if (timestamp instanceof Date) {
          formattedDate = `${timestamp.getFullYear()}/${('0' + (timestamp.getMonth() + 1)).slice(-2)}/${('0' + timestamp.getDate()).slice(-2)}`;
        }
        return {
          status: 'Used',
          poQuote: dbRow[DB_POSKU_KEY_COL - 1] || 'N/A',
          bol: dbRow[DB_BOL_COL - 1] || 'N/A',
          date: formattedDate || 'N/A'
        };
      }
    }
    return { status: 'Available' };
  } catch (e) {
    Logger.log(`getSerialStatus Error: ${e.message}`);
    return { status: 'Error', message: e.message };
  }
}
