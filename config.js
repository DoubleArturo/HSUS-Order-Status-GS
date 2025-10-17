/**
 * Config.js
 * * 集中管理所有 Google Sheet 工作表的名稱、ID，以及關鍵欄位的索引。
 * 在架構重構後，這個檔案將用於所有的資料存取抽象化 (SheetService.js)。
 * * 欄位標頭 (Headers) 將在 SheetService.js 中動態獲取，這裡主要定義 Sheet 名稱。
 * * 關鍵 Sheets (根據您的 Raw Data 和 DB)
 */

const SHEET_NAMES = {
  // 原始數據
  DEALER_PO_RAW: 'Dealer PO | Raw Data',
  DIRECT_QUOTE_RAW: 'Direct Quote | Raw Data',
  AR_QBO_RAW: 'AR from QBO | Raw Data',
  QBO_INVOICE_RAW: 'QBO Invoice | Raw Data',
  QBO_ESTIMATE_RAW: 'QBO Estimate | Raw Data',
  SERIAL_RAW: 'Serial # | Raw Data',

  // 資料庫 (將優先遷移至 Firestore，但目前仍視為 Sheets)
  GIT_DB: 'GIT Tool | DB',
  SERIAL_DB: 'Serial #_DB',
  BOL_DB: 'BOL_DB',
  SHIPMENT_PLANNING_DB: 'Shipment_Planning_DB',

  // 管理與儀表板 (Mgt. Table & Dashboard)
  PO_PROCESSING_QUEUE: 'PO Processing Queue',
  ORDER_SHIPPING_MGT: 'Order Shipping Mgt. Table',
  AR_MGT: 'AR ｜Mgt. Table',
  OPERATION_DASHBOARD: 'Operation | Pending Order Dashboard',
  MANAGER_DASHBOARD: 'Manager Dashboard -Orders',

  // 靜態資料
  PRICE_BOOK: 'HSUS Price Book',
  VENDORS: 'Vendors',
  RSM: 'RSM',
  CUSTOMERS_QBO: 'Customers(QBO)',
  SPIFF: 'SPIFF',
  STATUS_REMIND: 'Status & Remind',
  PANDADOC_STATUS: 'PandaDoc Status Definition'
};

// 由於您的核心操作是透過 PO# 或 Helper Key，我們定義主要的鍵 (Key) 欄位名稱。
const PRIMARY_KEYS = {
  PO_NUMBER: 'P/O',
  PO_SKU_KEY: 'Helper Key', // 通常是 PO#|SKU
  SERIAL_NUMBER: 'Serial #',
  BOL_NUMBER: 'BOL #'
};

// 導出配置供其他 .gs 檔案使用
// 在 App Script 中，所有變數都是全域的，但為了程式碼可讀性，我們用一個 Object 集中定義。
const Config = {
  SHEET_NAMES,
  PRIMARY_KEYS,
  // 您的 Google Sheet 文件 ID (請替換為您的實際 ID)
  // 您可以在瀏覽器 URL 中找到它：.../d/[YOUR_DOCUMENT_ID]/edit
  DOCUMENT_ID: 1WdJZkuiLCH- fM5LMX6leJ0XIpZCreItCSLAPN2VasqA
  // TODO: 請您手動替換上方 YOUR_GOOGLE_SHEET_ID_HERE
};

// 警告: App Script 舊版運行環境不支持 ES6 模組導出 (export default)。
// 我們將依賴 Apps Script 的全域作用域。
// 在其他 .gs 檔案中，直接使用 Config.SHEET_NAMES, Config.PRIMARY_KEYS 即可。
