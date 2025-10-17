/**
 * Displays the HTML dialog for manual PO creation.
 */
function showManualPOSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Manual New PO')
      .setTitle('Manual New PO')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manual New PO');
}

/**
 * Serves the initial data for the sidebar.
 */
function getInitialData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const customersSheet = spreadsheet.getSheetByName('Customers(QBO)');
  const customerNames = customersSheet.getRange('B2:B300').getValues().flat().filter(String);

  const priceBookSheet = spreadsheet.getSheetByName('HSUS Price Book');
  const models = priceBookSheet.getRange('C2:C200').getValues().flat().filter(String);
  const prices = priceBookSheet.getRange('D2:D200').getValues().flat().filter(String);

  // 在後端生成 P/O # 和 Created Date
  const poNumber = `POM${Utilities.getUuid().substring(0, 4).toUpperCase()}`;
  const createdDate = new Date().toISOString().split('T')[0];

  return { customerNames, models, prices, poNumber, createdDate };
}

/**
 * Backend function to process PO data and save it to the sheet.
 * @param {Object} poData The PO data from the sidebar.
 * @returns {Object} A status object with success/error message.
 */
function processAndSavePo(poData) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Dealer PO | Raw Data');
  if (!sheet) {
    return { status: 'error', message: '工作表 "Dealer PO | Raw Data" 未找到。' };
  }

  try {
    const lastRow = sheet.getLastRow();
    const existingDataRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    const existingValues = existingDataRange.getValues();

    // 準備所有產品項目的資料列
    const lineItemRows = [];
    poData.lineItems.forEach(item => {
      // 確保每一行都有足夠的欄位數，避免索引超出範圍
      const lineItemRow = new Array(existingValues[0].length).fill('');
      
      // 填寫主要 PO 資訊
      lineItemRow[0] = poData.createdDate;       // Col A: Created Date
      lineItemRow[1] = poData.buyerName;         // Col B: Buyer Name
      lineItemRow[3] = poData.poNumber;          // Col D: P/O
      lineItemRow[7] = poData.total;             // Col H: P/O - Total
      lineItemRow[8] = poData.paymentTerm;       // Col I: Payment term
      lineItemRow[9] = poData.type;              // Col J: Type
      lineItemRow[18] = poData.shipToInfo.address; // Col S: Ship to
      lineItemRow[19] = poData.shipToInfo.contactPerson; // Col T: Contact Person
      lineItemRow[20] = poData.shipToInfo.phone;  // Col U: Phone
      lineItemRow[21] = poData.shipToInfo.email;  // Col V: Email

      // 填寫產品項目資訊
      lineItemRow[12] = item.model;      // Col M: item.model
      lineItemRow[13] = item.unitPrice;  // Col N: item.unitPrice
      lineItemRow[14] = item.quantity;   // Col O: item.quantity

      lineItemRows.push(lineItemRow);
    });

    if (lineItemRows.length > 0) {
      // 一次性寫入所有產品項目
      const newRange = sheet.getRange(lastRow + 1, 1, lineItemRows.length, lineItemRows[0].length);
      newRange.setValues(lineItemRows);
    }
    
    // After saving, send to Zapier
    // sendToZapier(poData); // TODO: Integrate Zapier call here

    return { status: 'success', message: `PO ${poData.poNumber} 已成功建立並儲存。` };

  } catch (e) {
    return { status: 'error', message: `儲存資料時發生錯誤：${e.message}` };
  }
}


/**
 * Generates a UUID for P/O Number.
 * Format: "POM" + 4 random characters.
 * @returns {string} The generated UUID.
 */
function generatePoUuid() {
  const randomPart = Utilities.getUuid().substring(0, 4).toUpperCase();
  return `POM${randomPart}`;
}

// -----------------------------------------------------------------
// 以下為舊函式，用於 Zapier 回傳。
// -----------------------------------------------------------------

/**
 * Zapier 回傳後，將 PDF 連結存回 Google Sheets。
 * @param {string} poNumber 訂單號碼。
 * @param {string} pdfUrl PandaDoc PDF 的 URL。
 * @returns {string} 成功或失敗訊息。
 */
function savePdfUrl(poNumber, pdfUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dealer PO | Raw Data');
  if (!sheet) {
    throw new Error('Sheet "Dealer PO | Raw Data" not found.');
  }

  const poColumn = 4; // D 欄
  const pdfUrlColumn = 16; // P 欄

  const range = sheet.getDataRange();
  const values = range.getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][poColumn - 1] == poNumber) {
      sheet.getRange(i + 1, pdfUrlColumn).setValue(pdfUrl);
      return `success: PDF URL for PO ${poNumber} saved.`;
    }
  }

  return `fail: PO ${poNumber} not found in sheet.`;
}
