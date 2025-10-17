/**
 * @fileoverview Backend script for the consolidated PO Management Tool.
 * [VERSION 2.1 - FULLY POPULATED CODE]
 * Handles New, Update, and Void PO operations.
 * New PO uploads feature immediate processing for the first file, with subsequent files queued.
 * Also provides initial data for UI dropdowns.
 */

// --- CONFIGURATION ---
const UPLOAD_FOLDER_ID = "1MUAdxzAk9uaqPKP5h47NYW9ltNrCQQZv";
const RAW_PO_SHEET_NAME = 'Dealer PO | Raw Data';
const CUSTOMERS_SHEET_NAME = 'Customers(QBO)';
const PO_NUMBER_COLUMN = 4; // Column D contains PO Numbers
const PO_STATUS_COLUMN = 25; // Column Y contains the Status
const PO_CHANGE_TIME_COLUMN = 27; // Column AA for change timestamp
const TEMP_UPLOAD_FOLDER_ID = "1HDri2xgl9UACHpSDXpNqT9VqgrmrSYbd";

// --- QUEUE CONFIGURATION ---
const NEW_PO_QUEUE_KEY = 'NEW_PO_UPLOAD_QUEUE';
const NEW_PO_TRIGGER_HANDLER = 'processNewPoUploadQueue';
const DELAY_BETWEEN_UPLOADS_MS = 30 * 1000; // 30 seconds delay

/**
 * Opens the main PO Management Tool dialog.
 */
function openNewReviseVoidPO() {
  const html = HtmlService.createHtmlOutputFromFile('NewReviseVoidPO')
      .setWidth(480)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'New/Revise/Void PO Tool');
}

/**
 * [Backend Function] 
 * Fetches all initial data needed for the PO Mgt. Tool UI in one efficient call.
 */
function getPoMgtInitialData() {
  try {
    return {
      buyerNames: getBuyerNames_(),
      existingPoNumbers: getExistingPoNumbers_(),
      activePoNumbers: getActivePoNumbers_()
    };
  } catch (e) {
    Logger.log(`Error in getPoMgtInitialData: ${e.message}`);
    throw new Error("Failed to load initial data for dropdowns: " + e.message);
  }
}

// --- DATA FETCHING HELPER FUNCTIONS (Internal) ---

function getBuyerNames_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOMERS_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet '${CUSTOMERS_SHEET_NAME}' not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange("B2:B" + lastRow);
  const buyerNames = range.getValues().flat().filter(name => name);
  return [...new Set(buyerNames)].sort();
}

function getExistingPoNumbers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW_PO_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet '${RAW_PO_SHEET_NAME}' not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, PO_NUMBER_COLUMN, lastRow - 1, 1);
  const poNumbers = range.getValues().flat().filter(po => po);
  return [...new Set(poNumbers)].sort();
}

function getActivePoNumbers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW_PO_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet '${RAW_PO_SHEET_NAME}' not found.`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, PO_NUMBER_COLUMN, lastRow - 1, PO_STATUS_COLUMN - PO_NUMBER_COLUMN + 1);
  const values = range.getValues();
  const statusColumnIndex = PO_STATUS_COLUMN - PO_NUMBER_COLUMN;
  const activePOs = new Set();
  values.forEach(row => {
    const poNumber = row[0];
    const status = row[statusColumnIndex];
    if (poNumber && status !== 'Voided') {
      activePOs.add(poNumber);
    }
  });
  return [...activePOs].sort();
}

/**
 * [REVISED ARCHITECTURE]
 * Handles new PO uploads by immediately saving the file to a temporary location
 * and queuing only the file ID for processing.
 */
function processNewPoUpload(fileContent, buyerName, poNumber) {
  try {
    if (!buyerName || !poNumber) throw new Error("Buyer Name and PO Number are required.");
    
    // --- 核心邏輯變更 (1/2): 先儲存檔案 ---
    const decodedBlob = decodeBase64_(fileContent);
    const tempFolder = DriveApp.getFolderById(TEMP_UPLOAD_FOLDER_ID);
    
    // 儲存一個帶有時間戳的暫存檔，避免檔名衝突
    const tempFileName = `temp_${new Date().getTime()}_${poNumber}.pdf`;
    const tempFile = tempFolder.createFile(decodedBlob.setName(tempFileName));
    const tempFileId = tempFile.getId();
    
    console.log(`File temporarily saved with ID: ${tempFileId}`);

    // --- 核心邏輯變更 (2/2): 只將輕量級的指令放入佇列 ---
    const finalFileName = `${buyerName.trim()}_${poNumber.trim()}.pdf`;
    const task = { 
      tempFileId: tempFileId, // 不再儲存 fileContent
      finalFileName: finalFileName, 
      submittedAt: new Date().toISOString() 
    };
    
    addTaskToNewPoQueue_(task); // addTaskToNewPoQueue_ 函式本身不需要修改
    
    return { success: true, fileName: finalFileName };
    
  } catch (e) {
    Logger.log(`Failed to queue new PO: ${e.message}`);
    throw new Error(`Failed to queue new PO: ${e.message}`);
  }
}


function processPoUpdate(fileContent, poNumber) {
  try {
    if (!poNumber) throw new Error("An existing PO Number must be selected.");
    const decodedBlob = decodeBase64_(fileContent);
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const files = folder.searchFiles(`title contains '_${poNumber.trim()}.pdf'`);
    let baseFileName = `unknown-buyer_${poNumber.trim()}`;
    if (files.hasNext()) {
      const originalFile = files.next();
      baseFileName = originalFile.getName().replace('.pdf', '').split('_updated_')[0];
    }
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const updatedFileName = `${baseFileName}_updated_${today}.pdf`;
    folder.createFile(decodedBlob.setName(updatedFileName));
    revisePoStatus_(poNumber);

    // --- MODIFICATION START ---
    // Automatically run the archive function after a successful revision.
    console.log(`PO #${poNumber} has been revised. Triggering archive function...`);
    archiveProcessedPOs_Safe(); 
    // --- MODIFICATION END ---

    return `Success: Updated PO PDF '${updatedFileName}' has been uploaded.`;
  } catch (e) {
    Logger.log(`Update failed: ${e.message}`);
    throw new Error(`Update failed: ${e.message}`);
  }
}

/**
 * Voids a PO by updating its status in the sheet.
 * [CORRECTED VERSION to handle data type mismatch and spaces]
 */
function voidPo(poNumber) {
  try {
    if (!poNumber) throw new Error("No PO number selected.");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW_PO_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${RAW_PO_SHEET_NAME}' not found.`);
    
    // 範圍從 D2 開始，到工作表的最後一列的 Y 欄
    const range = sheet.getRange(2, PO_NUMBER_COLUMN, sheet.getLastRow() - 1, PO_STATUS_COLUMN - PO_NUMBER_COLUMN + 1);
    const values = range.getValues();
    let voidCount = 0;
    
    // 將從 UI 傳入的 poNumber 預先處理一次
    const targetPoNumber = String(poNumber).trim();
    // 計算狀態欄在我們讀取的範圍中的相對索引
    const statusColumnIndex = PO_STATUS_COLUMN - PO_NUMBER_COLUMN;

    for (let i = 0; i < values.length; i++) {
      // 將工作表中的值也進行處理，再進行比對
      const currentPoNumber = String(values[i][0]).trim();
      
      // --- 核心修正：將兩邊都轉為文字再比對，解決類型不匹配問題 ---
      if (currentPoNumber === targetPoNumber) {
        // 在讀取的二維陣列中直接修改狀態
        values[i][statusColumnIndex] = 'Voided';
        voidCount++;
      }
    }

    if (voidCount > 0) {
      // 將修改後的整個陣列一次性寫回工作表，效率最高
      range.setValues(values);
      
      console.log(`PO #${poNumber} has been voided. Triggering archive function...`);
      archiveProcessedPOs_Safe();
      
      return { success: true, message: `Successfully voided PO #${poNumber} (${voidCount} rows affected).` };
    } else {
      throw new Error(`PO #${poNumber} not found or already voided.`);
    }
  } catch (e) {
    Logger.log(e);
    throw new Error(e.message);
  }
}

// --- QUEUE AND TRIGGER MANAGEMENT SYSTEM ---

/**
 * Adds a new task to the queue. If the queue was empty,
 * it immediately starts processing the first task.
 * @param {object} task The task object to add.
 */
function addTaskToNewPoQueue_(task) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const queueJson = scriptProperties.getProperty(NEW_PO_QUEUE_KEY) || '[]';
  const queue = JSON.parse(queueJson);
  const wasQueueEmpty = queue.length === 0;
  queue.push(task);
  scriptProperties.setProperty(NEW_PO_QUEUE_KEY, JSON.stringify(queue));
  if (wasQueueEmpty) {
    Logger.log('Queue was empty. Starting processing immediately.');
    processNewPoUploadQueue();
  } else {
    Logger.log('A task is already being processed. The new task will be handled in sequence.');
  }
}

function manageNewPoQueueTrigger_() {
  const existingTriggers = ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === NEW_PO_TRIGGER_HANDLER);
  if (existingTriggers.length === 0) {
    ScriptApp.newTrigger(NEW_PO_TRIGGER_HANDLER).timeBased().after(DELAY_BETWEEN_UPLOADS_MS).create();
    Logger.log(`Scheduling next task in ${DELAY_BETWEEN_UPLOADS_MS / 1000} seconds.`);
  }
}

/**
 * [REVISED ARCHITECTURE] 
 * Processes tasks from the new PO upload queue.
 * It now retrieves the file using the ID from the queue and renames/moves it.
 */
function processNewPoUploadQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    console.warn('Could not acquire lock, another queue process is likely running.');
    return;
  }

  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === NEW_PO_TRIGGER_HANDLER) {
        ScriptApp.deleteTrigger(trigger);
        console.log('Deleted an existing trigger to prevent duplicate runs.');
      }
    });

    const scriptProperties = PropertiesService.getScriptProperties();
    const queueJson = scriptProperties.getProperty(NEW_PO_QUEUE_KEY);

    if (!queueJson || queueJson === '[]') {
      scriptProperties.deleteProperty(NEW_PO_QUEUE_KEY);
      return;
    }

    const queue = JSON.parse(queueJson);
    const task = queue.shift();

    console.log(`Processing task for final file: ${task.finalFileName}. Tasks remaining: ${queue.length}`);

    // --- 核心邏輯變更: 從 Drive 取回檔案並處理 ---
    try {
      const finalFolder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
      const tempFile = DriveApp.getFileById(task.tempFileId);
      
      // 將暫存檔案移動並重新命名到最終位置
      tempFile.moveTo(finalFolder); 
      tempFile.setName(task.finalFileName);
      
      console.log(`✅ File ${task.finalFileName} processed and moved successfully.`);
    } catch (fileError) {
      console.error(`Failed to process file with ID ${task.tempFileId}. Error: ${fileError.message}`);
      // 考慮將錯誤的 task 存到另一個地方以便手動處理
    }
    // --- 核心邏輯變更 END ---

    if (queue.length > 0) {
      scriptProperties.setProperty(NEW_PO_QUEUE_KEY, JSON.stringify(queue));
      console.log('Tasks remain in the queue. Scheduling next run.');
      manageNewPoQueueTrigger_();
    } else {
      scriptProperties.deleteProperty(NEW_PO_QUEUE_KEY);
      console.log('🎉 All tasks processed. Queue is empty and the property has been deleted.');
    }

  } catch (e) {
    console.error(`Error in processNewPoUploadQueue: ${e.message}. Stack: ${e.stack}`);
  } finally {
    lock.releaseLock();
  }
}

// --- HELPER FUNCTIONS ---

function decodeBase64_(base64String) {
  const base64Data = base64String.split(',')[1];
  return Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, "temp.pdf");
}

/**
 * [REVISED FUNCTION]
 * Finds all active rows for a given PO number and updates their status to 'Revised'.
 * This version ensures data type consistency for comparison and only updates non-voided/non-revised rows.
 * @param {string|number} poNumber The PO number to revise.
 */
function revisePoStatus_(poNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW_PO_SHEET_NAME);
  if (!sheet) {
    console.error(`Sheet '${RAW_PO_SHEET_NAME}' not found. Cannot revise status.`);
    return; // Exit if the sheet doesn't exist
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const targetPoString = String(poNumber).trim(); // Convert the target PO to a trimmed string once.

  let revisedCount = 0;

  // Loop through all rows, starting from the second row (index 1) to skip the header.
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Get current data and convert to string for safe comparison
    const currentPoNumber = String(row[PO_NUMBER_COLUMN - 1]).trim(); // Column D
    const currentStatus = String(row[PO_STATUS_COLUMN - 1]).trim();   // Column Y

    // --- KEY LOGIC CHANGE ---
    // The condition now checks for three things:
    // 1. Does the PO number match? (String vs String comparison)
    // 2. Is the current status NOT 'Revised' already?
    // 3. Is the current status NOT 'Voided'?
    if (currentPoNumber === targetPoString && currentStatus !== 'Revised' && currentStatus !== 'Voided') {
      
      const rowIndex = i + 1; // getRange is 1-based, array is 0-based
      
      // Update the status in Column Y
      sheet.getRange(rowIndex, PO_STATUS_COLUMN).setValue('Revised');
      
      // Update the timestamp in Column AA
      sheet.getRange(rowIndex, PO_CHANGE_TIME_COLUMN).setValue(new Date());
      
      revisedCount++;
      console.log(`Row ${rowIndex} with PO #${targetPoString} has been marked as 'Revised'.`);
    }
  }

  if (revisedCount > 0) {
    SpreadsheetApp.flush(); // Apply all pending changes at once.
    console.log(`Successfully revised ${revisedCount} row(s) for PO #${targetPoString}.`);
  } else {
    console.warn(`No active rows found to revise for PO #${targetPoString}.`);
  }
}