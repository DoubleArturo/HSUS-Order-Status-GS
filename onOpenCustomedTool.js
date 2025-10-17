// Code.gs

/**
 * 當試算表被開啟時，這個特殊的函式會自動執行，用來建立自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Customed Order Tools')
    .addItem('Step 0: New/Revise/Void PO', 'openNewReviseVoidPO')
    .addItem('Step 0: Manual New PO', 'showManualPOSidebar')
    .addItem('Step 1: PO Editor', 'openPOEditor')
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

/**
 * ⚡️ [重要提醒] 您的菜單中也包含以下函數，為確保 onOpen 正常運作，它們也必須被定義：
 */
function openNewReviseVoidPO() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
function showManualPOSidebar() {
  SpreadsheetApp.getUi().alert('This feature is currently under development.');
}
//... (And all other functions called in your onOpen menu)

// --- 我們將共用的工具函式放在這裡 ---

/**
 * 安全地寫入受保護的儲存格。
 * 此函式會暫時為當前使用者授權，執行寫入操作，然後立刻恢復保護。
 * @param {GoogleAppsScript.Spreadsheet.Range} range 要寫入的目標儲存格 (一個 Range 物件)。
 * @param {any} value 要寫入的值。
 * @param {GoogleAppsScript.Base.Ui} ui 可選的 UI 物件，用於顯示提醒。
 * @returns {boolean} 如果寫入成功則返回 true，否則返回 false。
 */
function secureWrite(range, value, ui = SpreadsheetApp.getUi()) {
  const protection = range.getProtections(SpreadsheetApp.ProtectionType.RANGE)[0];
  const currentUser = Session.getEffectiveUser();

  // 情況 1: 儲存格沒有保護，或腳本擁有者本身就有權限
  if (!protection || protection.canEdit()) {
    try {
      range.setValue(value);
      return true;
    } catch (e) {
      ui.alert(`寫入儲存格 ${range.getA1Notation()} 時發生錯誤: ${e.message}`);
      return false;
    }
  }

  // 情況 2: 儲存格被保護，且腳本擁有者需要暫時授權
  protection.addEditor(currentUser);
  
  try {
    // 執行核心的寫入操作
    range.setValue(value);
    
    // 操作完成後，立刻移除使用者的編輯權限
    protection.removeEditor(currentUser);
    return true;

  } catch (e) {
    ui.alert(`寫入受保護儲存格 ${range.getA1Notation()} 時發生錯誤: ${e.message}`);
    return false;

  } finally {
    // 使用 'finally' 區塊確保無論成功或失敗，權限恢復的動作都一定會被執行
    protection.removeEditor(currentUser);
  }
}
