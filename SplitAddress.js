/**
 * @fileoverview
 * 將 'Dealer PO | Raw Data' S 欄中的地址拆分，並填入同一工作表的 AH-AK 欄。
 * 僅在 D 欄 (PO Number) 有值時執行。
 * 版本: 2.1
 */

// --- 常數設定 ---
const SHEET_NAME = 'Dealer PO | Raw Data';

// 欄位對應 (1-based index)
const COL = {
  PO_NUMBER: 4,         // D欄 (判斷條件)
  ADDRESS: 19,          // S欄 (來源)
  STREET_ADDRESS: 34,   // AH欄 (目標)
  CITY: 35,             // AI欄 (目標)
  STATE: 36,            // AJ欄 (目標)
  ZIPCODE: 37           // AK欄 (目標)
};

/**
 * 主函式：讀取來源資料、拆分地址，並寫回同一工作表。
 */
function splitAddressInPlace() {
  const ss = SpreadsheetApp.getActiveSpreadpreasheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`錯誤：找不到工作表 "${SHEET_NAME}"`);
    return;
  }

  // 1. 一次性讀取工作表中所有資料 (從第二列開始)
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();

  // 2. 遍歷每一列資料，進行地址拆分
  data.forEach(row => {
    const poNumber = row[COL.PO_NUMBER - 1]; // 讀取 D 欄的 PO 編號
    const fullAddress = row[COL.ADDRESS - 1]; // 讀取 S 欄的地址

    // --- ⭐ 核心修改：增加 PO 編號的判斷條件 ---
    // 只有當 D 欄有值，且 S 欄的地址也是有效的字串時，才執行拆分
    if (poNumber && fullAddress && typeof fullAddress === 'string' && fullAddress.trim() !== '') {
      const parsedAddress = parseUsAddress(fullAddress);
      
      // 將拆分後的結果，填入記憶體中 data 陣列的對應位置
      row[COL.STREET_ADDRESS - 1] = parsedAddress.street;
      row[COL.CITY - 1] = parsedAddress.city;
      row[COL.STATE - 1] = parsedAddress.state;
      row[COL.ZIPCODE - 1] = parsedAddress.zipcode;
    }
  });

  // 3. 將包含更新後地址的整個 data 陣列，一次性寫回工作表，效能最佳
  dataRange.setValues(data);
  
  SpreadsheetApp.getUi().alert('地址拆分並更新完成！');
}

/**
 * 使用正規表示式解析美國地址。(此函式保持不變)
 * @param {string} addressString - 完整的地址字串。
 * @returns {{street: string, city: string, state: string, zipcode: string}} 解析後的地址物件。
 */
function parseUsAddress(addressString) {
  // 清理地址，將換行符統一為逗號，並移除多餘空白
  const cleanedAddress = addressString.replace(/\n/g, ', ').replace(/,+/g, ',').trim();
  
  // 正規表示式，用於匹配 "City, ST ZIP" 或 "City ST ZIP"
  const regex = /([\w\s\.,#-]+?),\s*([\w\s]+),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)\s*$/;
  const match = cleanedAddress.match(regex);
  
  if (match) {
    return {
      street: (match[1] || '').trim().replace(/,$/, ''), // 清理尾隨的逗號
      city:   (match[2] || '').trim(),
      state:  (match[3] || '').trim(),
      zipcode:(match[4] || '').trim()
    };
  } else {
    // 如果正規表示式不匹配，將原始地址放入 Street 欄位，以便手動檢查
    return { street: addressString, city: '', state: '', zipcode: '' };
  }
}