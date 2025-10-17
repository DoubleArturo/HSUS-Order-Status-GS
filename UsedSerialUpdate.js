/**
 * @file AutoAssignOrder.gs
 * @version 3.0
 * @description Automatically assigns or removes an order number from serial numbers
 * when the list in 'Order Shipping Mgt. Table' is updated, added to, or deleted from.
 * @last-updated 2025-08-12
 */

/**
 * This is a special simple trigger (Simple Trigger) that runs automatically
 * when a user edits any cell in the spreadsheet.
 * @param {Object} e - The event object provided by Google Sheets, containing information about the edit.
 */
function onEditSerial(e) {
  // --- Configuration Section ---
  const sourceSheetName = "Order Shipping Mgt. Table"; // Name of the trigger sheet
  const sourceSerialCol = 24; // Trigger column: Serial number column (Col X)
  const sourceOrderCol = 2;   // Order number column (Col B)

  const targetSheetName = "Serial # | Raw Data"; // Name of the sheet to be updated
  const targetSerialCol = 4;  // Serial number column in the target sheet (Col D)
  const targetOrderCol = 14; // Column for the order number to be written to (Col N)
  // --- End of Configuration ---

  const range = e.range;
  const editedSheet = range.getSheet();
  
  // 1. Check: Was the edit in the correct sheet and column?
  if (editedSheet.getName() !== sourceSheetName || range.getColumn() !== sourceSerialCol) {
    return; // If not, abort the script
  }

  // 2. Get new and old data
  const orderNumber = editedSheet.getRange(range.getRow(), sourceOrderCol).getValue();
  const oldValue = e.oldValue || ''; // Content before the edit
  const newValue = e.value || '';   // Content after the edit

  // If the order number is empty, "add" operation cannot be performed
  if (!orderNumber && newValue) {
    SpreadsheetApp.getUi().alert(`Cannot tag serial number because the order number in column B is empty.`);
    range.setValue(oldValue);
    return;
  }
  
  // 3. Create sets of old and new serial numbers to easily compare the differences
  const oldSerials = new Set(oldValue.split(',').map(s => s.trim()).filter(Boolean));
  const newSerials = new Set(newValue.split(',').map(s => s.trim()).filter(Boolean));

  // 4. Calculate which serial numbers need to be "added" and "removed"
  const serialsToAdd = [...newSerials].filter(s => !oldSerials.has(s));
  const serialsToRemove = [...oldSerials].filter(s => !newSerials.has(s));

  if (serialsToAdd.length === 0 && serialsToRemove.length === 0) {
    return; // If there are no changes, abort the script
  }

  // 5. Prepare to update the target sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);

  if (!targetSheet) {
    SpreadsheetApp.getUi().alert(`Error: The sheet named "${targetSheetName}" was not found.`);
    return;
  }
  
  // --- Performance Optimization Start: Batch read and write ---
  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return;
  
  // Read all data from the target sheet into memory at once
  const serialData = targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn()).getValues();
  
  const serialMap = new Map();
  serialData.forEach((row, index) => {
    const serialValue = row[targetSerialCol - 1];
    if (serialValue) {
      serialMap.set(serialValue, index); // Store index for in-memory modification
    }
  });
  
  // Prepare an array to hold all the row indices to be updated
  const updates = [];

  // 6. Execute updates
  serialsToAdd.forEach(serial => {
    if (serialMap.has(serial)) {
      const rowIndex = serialMap.get(serial);
      serialData[rowIndex][targetOrderCol - 1] = orderNumber;
    }
  });

  serialsToRemove.forEach(serial => {
    if (serialMap.has(serial)) {
      const rowIndex = serialMap.get(serial);
      serialData[rowIndex][targetOrderCol - 1] = ''; // Clear the cell
    }
  });
  
  // Write all the modified data back to the sheet in one go
  targetSheet.getRange(2, 1, serialData.length, serialData[0].length).setValues(serialData);
  // --- Performance Optimization End: Batch read and write ---
}
