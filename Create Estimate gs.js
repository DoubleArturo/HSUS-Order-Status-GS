// Replace with your Zapier Webhook URL
const ZAPIER_WEBHOOK_URL = 'https://hooks.zapier.com/hooks/catch/13989939/u4w68ta/';
const DASHBOARD_SHEET_NAME = 'Operation | Pending Order Dashboard';

// --- [V2 修改] 更新欄位索引 ---
// Define column indices (A=1, B=2, ...)
const PO_COLUMN = 16;       // O column
const BOL_NUMBER_COLUMN = 6;  // E column
const EST_NUMBER_COLUMN = 8;  // G column (新増)
// SHIP_TO_INFO_COLUMN 已被刪除

/**
 * Displays the HTML sidebar.
 */
function createEstimateSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Create Estimate')
      .setTitle('Create Estimate')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Backend function to get pending orders based on criteria.
 * [已修改 V2] 移除 Ship To Info 條件，增加 Estimate # 條件。
 * @returns {Array<Object>} An array of pending order objects with only the 'po' property.
 */
function getPendingOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet not found: ${DASHBOARD_SHEET_NAME}`);
  }
  
  const values = sheet.getDataRange().getValues();
  const pendingOrders = [];
  
  // Start the loop from index 2, which corresponds to row 3
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    
    // --- [V2 修改] 更新篩選條件 ---
    const isBolNumberNotEmpty = (String(row[BOL_NUMBER_COLUMN - 1] || '').trim() !== '');
    // 檢查 G 欄 (Estimate #) 是否為空字串
    const isEstNumberEmpty = (String(row[EST_NUMBER_COLUMN - 1] || '').trim() === '');
    
    // 新的條件：BOL# 必須有內容，且 Estimate # 必須為空
    if (isBolNumberNotEmpty && isEstNumberEmpty) {
      // Push only the PO number into the array
      pendingOrders.push({
        po: row[PO_COLUMN - 1],
      });
    }
  }
  
  return pendingOrders;
}

/**
 * Backend function to trigger the Zapier Webhook.
 * @param {object} order The selected order data.
 * @returns {string} Success or failure message.
 */
function createEstimate(order) {
  if (!order || !order.po) {
    return 'fail: No order data provided.';
  }
  
  // Re-verify the ZAPIER_WEBHOOK_URL
  if (ZAPIER_WEBHOOK_URL === 'YOUR_ZAPIER_WEBHOOK_HERE' || !ZAPIER_WEBHOOK_URL) {
      throw new Error('Zapier Webhook URL is not set. Please update the ZAPIER_WEBHOOK_URL constant in your script.');
  }

  const payload = {
    po_number: order.po,
    row_number: order.row,
    full_data: order.data 
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };
  
  try {
    UrlFetchApp.fetch(ZAPIER_WEBHOOK_URL, options);
    return `success:${order.po}`;
  } catch (e) {
    // Catch the error and log it, then re-throw it so the frontend can catch it
    Logger.log('Failed to fetch URL: ' + e.message);
    throw new Error(`Failed to trigger Zapier. Please check your URL and permissions. Error: ${e.message}`);
  }
}