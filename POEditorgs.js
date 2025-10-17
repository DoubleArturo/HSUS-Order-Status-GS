/**
 * @fileoverview Backend script for the PO Data Correction Tool.
 * VERSION 9 - Final stable version: Fixes templateRow ReferenceError, robust queue system,
 * and includes ALL necessary menu stubs (openPOEditor, onOpen, etc.) for full stability.
 * ⚡️ ALIGNMENT FIX: savePoCorrections_AsyncWrapper renamed to savePoCorrections_AppendOnly.
 */

// --- GLOBAL CONSTANTS ---
const PO_RAW_DATA_SHEET = 'Dealer PO | Raw Data';
const PRICE_BOOK_SHEET_Editor = 'HSUS Price Book';
const QUEUE_SHEET = 'PO Processing Queue'; 

// --- Column mapping for 'Dealer PO | Raw Data' (1-based index) ---
const PO_COL = {
  PO_RECEIVED_DATE: 1, BUYER_NAME: 2, RSM: 3, PO_NUMBER: 4, MODEL_FROM_SHEET: 5,
  PO_TOTAL: 8, PAYMENT_TERM: 9, COMPANY: 11, P_O_LINE_ITEMS: 13, P_O_UNIT_PRICE: 14, 
  P_O_QTY: 15, FILE_URL: 16, SKU: 17, SHIP_TO_CONTACT: 20, SHIP_TO_PHONE: 21, 
  HELPER_KEY: 23, STATUS: 25, CHANGE_NOTE: 26, TIMESTAMP: 27, SPIFF: 30,
  STREET_ADDRESS: 34, CITY: 35, STATE: 36, ZIPCODE: 37
};

// --- HELPER FUNCTIONS (Queue Management) ---

/**
 * Gets or creates the queue sheet for status tracking.
 */
function getOrCreateQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_SHEET, ss.getSheets().length);
    sheet.getRange(1, 1, 1, 6).setValues([
      ['PO Number', 'Status', 'Submitted By', 'Submitted Time', 'Start Time', 'Completion Message']
    ]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

/**
 * Creates a single time-based trigger for background processing.
 */
function _createProcessTrigger_() {
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  const existing = triggers.some(t => t.getHandlerFunction() === '_processPoCorrectionTrigger');
  
  if (!existing) {
      ScriptApp.newTrigger('_processPoCorrectionTrigger')
          .timeBased()
          .at(new Date(new Date().getTime() + 5000)) // 5 seconds delay
          .create();
  }
}

function _deleteSelfTrigger() {
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === '_processPoCorrectionTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// --- UI AND DATA FETCH FUNCTIONS ---

/**
 * Opens the modal dialog interface for the PO Data Correction Tool.
 * This function is aligned with the user's onOpen menu call.
 */
function openPOEditor() {
  const html = HtmlService.createTemplateFromFile('POEditor')
      .evaluate()
      .setTitle('PO Data Correction Tool')
      .setWidth(800)
      .setHeight(750); 
  SpreadsheetApp.getUi().showModalDialog(html, 'PO Data Correction Tool');
}

/**
 * Fetches data for the UI. (Reads from proc_shipping_management and Price Book)
 */
function getCorrectionData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const procSheet = ss.getSheetByName('proc_shipping_management');
    const priceBookSheet = ss.getSheetByName('HSUS Price Book'); 

    if (!procSheet || !priceBookSheet) {
      throw new Error("Could not find required sheets: 'proc_shipping_management' or 'HSUS Price Book'");
    }
    
    // NOTE: PROC_COL mapping depends on proc_shipping_management structure
    const PROC_COL = {
      DATE: 1, PO_NUMBER: 2, BUYER_NAME: 3, COMPANY: 19, STREET: 4, CITY: 5, STATE: 6, ZIP: 7, 
      CONTACT: 8, PHONE: 9, MODEL: 10, QTY: 11, UNIT_PRICE: 12, ORIG_SKU: 14, STATUS: 15, 
      RSM: 16, PAYMENT_TERM: 17, FILE_URL: 18, DERIVED_SKU: 14, SPIFF: 20 
    };

    const procRange = procSheet.getRange(2, 1, procSheet.getLastRow() - 1, procSheet.getLastColumn()).getValues();
    const poMap = new Map();

    procRange.forEach((row, index) => {
      const poNumber = row[PROC_COL.PO_NUMBER - 1];
      const status = row[PROC_COL.STATUS - 1];
      
      if (!poNumber || status !== '') { return; }

      if (!poMap.has(poNumber)) {
        const rawDate = row[PROC_COL.DATE - 1];
        const formattedDate = rawDate instanceof Date ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        
        poMap.set(poNumber, {
          poNumber: poNumber, pdfUrl: row[PROC_COL.FILE_URL - 1], poReceivedDate: formattedDate,
          rsm: row[PROC_COL.RSM - 1], paymentTerm: row[PROC_COL.PAYMENT_TERM - 1],
          shipToContact: row[PROC_COL.CONTACT - 1], shipToPhone: row[PROC_COL.PHONE - 1],
          streetAddress: row[PROC_COL.STREET - 1] || '', city: row[PROC_COL.CITY - 1] || '',
          state: row[PROC_COL.STATE - 1] || '', zipcode: row[PROC_COL.ZIP - 1] || '',
          buyerName: row[PROC_COL.BUYER_NAME - 1] || '', company: row[PROC_COL.COMPANY - 1] || '', spiff: '',
          items: []
        });
      }

      poMap.get(poNumber).items.push({
        rowNumber: index + 2, model: row[PROC_COL.MODEL - 1] || '',
        sku: row[PROC_COL.DERIVED_SKU - 1] || row[PROC_COL.ORIG_SKU - 1] || '', 
        qty: row[PROC_COL.QTY - 1] || '', unitPrice: row[PROC_COL.UNIT_PRICE - 1] || ''
      });
    });

    const priceBookRange = priceBookSheet.getRange('A3:D' + priceBookSheet.getLastRow()).getValues();
    const modelToSkuMap = {};
    const modelToPriceMap = {};
    let modelNames = []; 

    priceBookRange.forEach(row => {
      const lookupName = row[0]; const sku = row[1]; const price = row[3];
      if (lookupName && sku) {
        if (!modelToSkuMap[lookupName]) { modelNames.push(lookupName); }
        modelToSkuMap[lookupName] = sku;
        modelToPriceMap[lookupName] = price;
      }
    });
    
    let finalModelNames = [...new Set(modelNames)].sort();
    const standardModels = finalModelNames.filter(name => name.includes('Standard'));
    const nonStandardModels = finalModelNames.filter(name => !name.includes('Standard'));
    finalModelNames = [...nonStandardModels.sort(), ...standardModels.sort()];
    
    return {
      success: true, allPOs: Array.from(poMap.values()), modelNames: finalModelNames, 
      modelToSkuMap: modelToSkuMap, modelToPriceMap: modelToPriceMap
    };

  } catch (e) {
    Logger.log(`getCorrectionData Error: ${e.toString()}`);
    return { success: false, message: e.toString() };
  }
}


// --- ASYNC QUEUE LOGIC ---

/**
 * Async Wrapper: Stores payload and schedules the core job.
 * ⚡️ RENAMED FROM savePoCorrections_AsyncWrapper TO savePoCorrections_AppendOnly
 */
function savePoCorrections_AppendOnly(poNumber, basicInfo, items) {
  try {
    const userEmail = Session.getActiveUser().getEmail(); 
    const timestamp = new Date().getTime();
    
    // 1. Store Payload in Cache
    const key = `poCorrection_${poNumber}_${timestamp}`;
    const dataPayload = { poNumber, basicInfo, items };
    CacheService.getUserCache().put(key, JSON.stringify(dataPayload), 3600);
    
    // 2. Record the job in the Queue Sheet
    const queueSheet = getOrCreateQueueSheet();
    queueSheet.appendRow([
      poNumber, 
      'Queued', // Initial status
      userEmail, 
      new Date(), 
      '', 
      key // Store the cache key
    ]);
    
    // 3. Create Trigger (only if one doesn't exist)
    _createProcessTrigger_(); 

    // 4. Return user-friendly English message
    return { success: true, message: "Your changes have been submitted. Processing will start shortly. Please check the 'PO Processing Queue' sheet for status." };

  } catch (e) {
    Logger.log(`savePoCorrections_AppendOnly Error: ${e.toString()}`);
    return { success: false, message: `Submission Failed: Please ensure full authorization or contact IT. (Error: ${e.toString()})` };
  }
}

/**
 * Background Processor: Executes queued jobs.
 */
function _processPoCorrectionTrigger() {
  const queueSheet = getOrCreateQueueSheet();
  const cache = CacheService.getUserCache();
  const lastRow = queueSheet.getLastRow();
  
  if (lastRow <= 1) {
    _deleteSelfTrigger();
    return;
  }
  
  const queueData = queueSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const updates = [];
  let jobsProcessed = false;

  queueData.forEach((row, index) => {
    const status = row[1];
    const cacheKey = row[5]; 
    
    if (status === 'Queued' && cacheKey) {
      jobsProcessed = true;
      const payload = cache.get(cacheKey);
      
      row[1] = 'In Progress';
      row[4] = new Date(); 
      
      let newStatus = 'Failed'; 
      let completionMessage = 'Processing failed: Data not found in cache.';
      
      if (payload) {
        try {
          const { poNumber, basicInfo, items } = JSON.parse(payload);
          const result = _savePoCorrectionsCore(poNumber, basicInfo, items);
          
          newStatus = result.success ? 'Success' : 'Failed';
          completionMessage = result.success ? `Successfully applied changes for PO #${result.newPoNumber}.` : `Processing failed (PO #${poNumber}): ${result.message}`;
          
        } catch (e) {
          completionMessage = `Unexpected System Error: ${e.toString()}`;
        }
        
        cache.remove(cacheKey);
      } 
      
      row[1] = newStatus;
      row[5] = completionMessage; 
    }
    updates.push(row.slice(1, 6)); // Collect columns B through F for final write
  });

  if (jobsProcessed) {
     const updateRange = queueSheet.getRange(2, 2, lastRow - 1, 5); 
     updateRange.setValues(updates);
  }
  
  _deleteSelfTrigger();
}


/**
 * Core function that performs the vectorized writes.
 */
function _savePoCorrectionsCore(poNumber, basicInfo, items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName(PO_RAW_DATA_SHEET);
  if (!poSheet) return { success: false, message: `System error: Could not find sheet '${PO_RAW_DATA_SHEET}'.` };

  const lastRow = poSheet.getLastRow();
  const numColumns = poSheet.getLastColumn();
  const newPoNumber = basicInfo.newPoNumber || poNumber;

  if (lastRow <= 1) {
    return { success: false, message: `The '${PO_RAW_DATA_SHEET}' sheet contains no data to process.` };
  }
  
  const existingDataRange = poSheet.getRange(2, 1, lastRow - 1, numColumns);
  const existingData = existingDataRange.getValues();
  
  let templateRow = null; 
  let changesMade = false;
  let poFound = false; 

  // 1. In-Memory Processing: Find template and update statuses
  for (let i = 0; i < existingData.length; i++) {
    const row = existingData[i];
    const currentPoNumber = row[PO_COL.PO_NUMBER - 1];
    const currentStatus = row[PO_COL.STATUS - 1];

    if (currentPoNumber === poNumber) {
        poFound = true;
        
        // Find template for File URL (P-column)
        if (!templateRow) {
            templateRow = row; 
        }
        
        // Mark active/unprocessed rows as 'Change'
        if (currentStatus === '') {
            row[PO_COL.STATUS - 1] = 'Change';
            row[PO_COL.CHANGE_NOTE - 1] = basicInfo.changeNote;
            row[PO_COL.TIMESTAMP - 1] = new Date();
            changesMade = true;
        }
    }
  }

  // CRITICAL CHECK: Ensure the PO was found AND we have a template
  if (!poFound || !templateRow) {
    return { success: false, message: `Original PO #${poNumber} not found in the raw data for updating. Check if it was archived/deleted or if the PO number is missing.` };
  }

  // 2. Vectorized Write I: Update old records' status
  if (changesMade) {
      poSheet.getRange(2, 1, existingData.length, numColumns).setValues(existingData);
  }

  // 3. Prepare and Append New Records (Uses guaranteed 'templateRow')
  const fileUrl = templateRow[PO_COL.FILE_URL - 1]; 
  const poTotal = items.reduce((sum, item) => sum + (parseFloat(item.qty) || 0) * (parseFloat(item.unitPrice) || 0), 0);

  const newRowsData = items.map(item => {
    const newRow = new Array(numColumns).fill('');
    
    // Fill in data based on PO_COL mappings
    newRow[PO_COL.PO_RECEIVED_DATE - 1] = new Date(basicInfo.poReceivedDate);
    newRow[PO_COL.BUYER_NAME - 1] = basicInfo.buyerName; 
    newRow[PO_COL.RSM - 1] = basicInfo.rsm;
    newRow[PO_COL.PO_NUMBER - 1] = newPoNumber;  
    newRow[PO_COL.PO_TOTAL - 1] = poTotal;
    newRow[PO_COL.PAYMENT_TERM - 1] = basicInfo.paymentTerm;
    newRow[PO_COL.COMPANY - 1] = basicInfo.company;
    newRow[PO_COL.P_O_LINE_ITEMS - 1] = item.model;
    newRow[PO_COL.P_O_UNIT_PRICE - 1] = parseFloat(item.unitPrice);
    newRow[PO_COL.P_O_QTY - 1] = parseFloat(item.qty);
    newRow[PO_COL.FILE_URL - 1] = fileUrl; // Critical data from template
    newRow[PO_COL.SHIP_TO_CONTACT - 1] = basicInfo.contact; 
    newRow[PO_COL.SHIP_TO_PHONE - 1] = basicInfo.phone;     
    newRow[PO_COL.STREET_ADDRESS - 1] = basicInfo.street;    
    newRow[PO_COL.CITY - 1] = basicInfo.city;
    newRow[PO_COL.STATE - 1] = basicInfo.state;
    newRow[PO_COL.ZIPCODE - 1] = basicInfo.zipcode;
    newRow[PO_COL.CHANGE_NOTE - 1] = basicInfo.changeNote;
    newRow[PO_COL.TIMESTAMP - 1] = new Date(); 
    newRow[PO_COL.SPIFF - 1] = basicInfo.spiff;

    return newRow;
  });

  // 4. Vectorized Write II: Append new records
  if (newRowsData.length > 0) {
    const startRowForNewData = poSheet.getLastRow() + 1;
    poSheet.getRange(startRowForNewData, 1, newRowsData.length, newRowsData[0].length).setValues(newRowsData);
  } else {
      return { success: false, message: `Failed: No line items were provided to save for PO #${newPoNumber}.` };
  }
  
  SpreadsheetApp.flush(); 

  return { success: true, newPoNumber: newPoNumber };
}


// --- ONOPEN ALIGNMENT & INITIAL SETUP ---

/**
 * Aligns with the user's custom menu to prevent errors.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Customed Order Tools')
    .addItem('Step 0: New/Revise/Void PO', 'openNewReviseVoidPO')
    .addItem('Step 0: Manual New PO', 'showManualPOSidebar')
    .addItem('Step 1: PO Editor', 'openPOEditor') // Aligned with the main editor
    .addSeparator()
    .addItem('Step 2: Shipping Mgt (Est.).', 'openShippingMgtTool')
    .addItem('Step 3: BOL# Entry (Act.)', 'openBolEntryTool')
    .addItem('Step 4: Serial Assignment', 'openSerialAssignmentTool')
    .addItem('Step 5: Create Estimate in QBO', 'createEstimateSidebar')
    .addSeparator()
    .addItem('GIT Mgt. Tool', 'openGitEditor')
    .addItem('Refresh PO Group Colors', 'recolorPoGroups')
    .addToUi();
}

/**
 * Opens the modal dialog interface for the PO Data Correction Tool.
 */
function openPOEditor() {
  const html = HtmlService.createTemplateFromFile('POEditor')
      .evaluate()
      .setTitle('PO Data Correction Tool')
      .setWidth(800)
      .setHeight(750); 
  SpreadsheetApp.getUi().showModalDialog(html, 'PO Data Correction Tool');
}


// ⚡️ STUB FUNCTIONS for Menu Alignment (These must exist to prevent "Script function not found" errors)
function openNewReviseVoidPO() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function showManualPOSidebar() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function openShippingMgtTool() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function openBolEntryTool() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function openSerialAssignmentTool() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function createEstimateSidebar() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function openGitEditor() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function recolorPoGroups() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}


/**
 * CORE: This function must be run manually ONCE for initial authorization.
 */
function createAllTriggers() {
  // 1. Delete all old triggers
  _deleteSelfTrigger();

  // 2. Try to create a test trigger and alert the user
  try {
    _createProcessTrigger_();
    SpreadsheetApp.getUi().alert('Success! Background processing is now set up. You can now use the PO Editor.');
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Setup Failed! Please ensure you have full authorization and try running this function again. Error: ${e.toString()}`);
  }
}
