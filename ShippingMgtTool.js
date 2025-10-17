/**

* @fileoverview Backend server-side script for the Shipping Management Tool.

* Handles all data interaction with Google Sheets.

* [MODIFIED to filter by SKU# and load existing plan data]

* [VERSION 2.1 - Corrected column ranges after address split]

*/



const PLANNING_SHEET_NAME1 = 'Shipment_Planning_DB';

const ORDER_SHEET_NAME = 'Order Shipping Mgt. Table';



/**

* Opens the sidebar interface for the Shipping Management Tool.

*/

function openShippingMgtTool() {

const html = HtmlService.createTemplateFromFile('ShippingMgtToolhtml.html')

.evaluate()

.setTitle('Shipping Management Tool');

SpreadsheetApp.getUi().showSidebar(html);

}



/**

* [CORRECTED] Fetches data for planning, filtering by non-empty SKU and including existing plan details.

* @returns {object} An object containing the list of pending items and their details.

*/

function getPlanningData() {

try {

const ss = SpreadsheetApp.getActiveSpreadsheet();

const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);

const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME1);



if (!orderSheet || !planningSheet) {

throw new Error("Required sheets ('Order Shipping Mgt. Table' or 'Shipment_Planning_DB') not found.");

}



// Step 1: Get all data from the planning sheet to find fulfilled items and existing plans.

const planningLastRow = planningSheet.getLastRow();

const fulfilledKeys = new Set();

const existingPlanDetails = {};

if (planningLastRow >= 2) {

const planningData = planningSheet.getRange('C2:G' + planningLastRow).getValues();

planningData.forEach(row => {

const key = row[0];

const estShipDate = row[1];

const qtyE = row[2];

const qtyW = row[3];

const status = row[4];

if (key) {

if (status === 'Fulfilled') {

fulfilledKeys.add(key);

}

existingPlanDetails[key] = {

estShipDate: estShipDate instanceof Date ? Utilities.formatDate(estShipDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '',

qtyE: qtyE,

qtyW: qtyW

};

}

});

}



// Step 2: Get total required quantities from the order sheet, filtering by SKU.

const lastRow = orderSheet.getLastRow();

if (lastRow < 3) {

return { success: true, pendingList: [], itemDetails: {} };

}


// --- ⭐ 修正點 1: 將讀取範圍擴大至 J3，以包含 Model Name ---

const fullOrderData = orderSheet.getRange('J3:U' + lastRow).getValues();

const itemDetails = {};



fullOrderData.forEach((row) => {

// 根據新的範圍 'K3:U' 調整索引

const modelName = row[1]; // K 欄 (Model Name)

const totalQty = row[2]; // L 欄 (Total Qty)

const sku = row[6]; // P 欄 (SKU)

const key = row[11]; // U 欄 (Key PO|SKU)


if (key && sku) {

if (!itemDetails[key]) {

// ⭐ 修正點 2: 首次建立 itemDetails 時，儲存 Model Name

itemDetails[key] = { totalQty: 0, modelName: modelName || 'N/A' };

}

itemDetails[key].totalQty += parseInt(totalQty, 10) || 0;

}

});


// Step 3 & 4 (合併計畫細節和過濾 fulfilled keys) 保持不變。



const allKeys = Object.keys(itemDetails);


// ⭐ 修正點 3: 建立最終待處理列表時，加入 Model Name 以供前端顯示

const rawPendingList = allKeys.filter(key => !fulfilledKeys.has(key));


const finalPendingList = rawPendingList.map(key => {

const model = itemDetails[key].modelName;

return `${key} (${model})`;

}).sort();


return {

success: true,

pendingList: finalPendingList, // 這是包含 Model Name 的顯示字串

itemDetails: itemDetails // 這是包含原始 key 和 modelName 的數據結構

};



} catch (e) {

Logger.log(e);

return { success: false, message: e.toString() };

}

}





/**

* [MODIFIED] Saves planning data by updating existing rows or appending new ones.

* @param {object} data The data object from the frontend form.

* @returns {object} An object containing the result of the operation.

*/

function savePlanningData(data) {

try {

const ss = SpreadsheetApp.getActiveSpreadsheet();

const planningSheet = ss.getSheetByName(PLANNING_SHEET_NAME1);

const user = Session.getActiveUser().getEmail();

const timestamp = new Date();



if (!planningSheet) {

throw new Error(`Sheet not found: ${PLANNING_SHEET_NAME1}`);

}


const totalQty = parseInt(data.totalQty, 10);

const qtyE = parseInt(data.qtyE, 10) || 0;

const qtyW = parseInt(data.qtyW, 10) || 0;



if (qtyE + qtyW !== totalQty) {

throw new Error(`Quantity mismatch! East (${qtyE}) + West (${qtyW}) does not equal Total Required (${totalQty}).`);

}



// Find if the row already exists

const lastRow = planningSheet.getLastRow();

let targetRowIndex = -1;

if (lastRow >= 2) {

const keys = planningSheet.getRange('C2:C' + lastRow).getValues();

for (let i = 0; i < keys.length; i++) {

if (keys[i][0] === data.poSkuKey) {

targetRowIndex = i + 2;

break;

}

}

}



const rowData = [

timestamp,

user,

data.poSkuKey,

new Date(data.estShipDate),

qtyE,

qtyW

];



if (targetRowIndex !== -1) {

// Update existing row (Columns A to F)

planningSheet.getRange(targetRowIndex, 1, 1, 6).setValues([rowData]);

} else {

// Append new row

planningSheet.appendRow(rowData);

}



SpreadsheetApp.flush();



return { success: true, message: 'Planning data saved successfully!' };



} catch (e) {

Logger.log(e);

return { success: false, message: e.toString() };

}

}