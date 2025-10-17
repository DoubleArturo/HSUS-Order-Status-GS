/**
 * @fileoverview Backend server-side script for the GIT Progress Editor.
 * Handles fetching and updating shipping progress for PIs.
 */

// --- 常數定義區 ---
const GIT_SHEET_NAME = 'GIT Tool | DB';
const GIT_CACHE_KEY = 'gitPendingPiData';


/**
 * Opens the sidebar interface for the GIT Progress Editor.
 * 您需要手動在 Google Sheet 中執行此函式，或建立一個自訂選單來觸發它。
 */
function openGitEditor() {
  const html = HtmlService.createTemplateFromFile('GITmanageTool.html')
    .evaluate()
    .setTitle('GIT Progress Editor');
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * 獲取所有未完成的 PI# 列表。
 * @returns {object} An object containing the list of pending PI numbers.
 */
function getGitData() {
  try {
    // 嘗試從快取讀取
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(GIT_CACHE_KEY);
    if (cachedData != null) {
      return { success: true, piList: JSON.parse(cachedData) };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gitSheet = ss.getSheetByName(GIT_SHEET_NAME);
    if (!gitSheet) throw new Error(`Sheet '${GIT_SHEET_NAME}' not found.`);
    
    const lastRow = gitSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, piList: [] }; // 工作表是空的
    }

    const dataRange = gitSheet.getRange('A2:G' + lastRow).getValues();
    const pendingPiList = [];

    dataRange.forEach(row => {
      const piNumber = row[0]; // Column A
      const isFinished = row[6]; // Column G (Finish)

      // 如果 PI# 存在且 'Finish' 欄不為 TRUE，則加入列表
      if (piNumber && isFinished !== true) {
        pendingPiList.push(piNumber);
      }
    });
    
    const sortedList = pendingPiList.sort();

    // 將結果存入快取，有效期限 5 分鐘
    cache.put(GIT_CACHE_KEY, JSON.stringify(sortedList), 300);

    return { success: true, piList: sortedList };
  } catch (e) {
    Logger.log(`getGitData Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 根據 PI# 獲取其詳細運輸資訊。
 * [已修正 V2 - 時區問題]
 * @param {string} piNumber The PI number to look up.
 * @returns {object} An object containing the details for the given PI.
 */
function getPiDetails(piNumber) {
  try {
    if (!piNumber) throw new Error("PI Number is required.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gitSheet = ss.getSheetByName(GIT_SHEET_NAME);
    if (!gitSheet) throw new Error(`Sheet '${GIT_SHEET_NAME}' not found.`);

    const piColumnValues = gitSheet.getRange('A2:A' + gitSheet.getLastRow()).getValues();
    let targetRow = -1;

    for (let i = 0; i < piColumnValues.length; i++) {
      if (piColumnValues[i][0] === piNumber) {
        targetRow = i + 2; // 找到 PI# 所在的列號
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: `PI# '${piNumber}' not found.` };
    }

    const rowData = gitSheet.getRange(targetRow, 2, 1, 6).getValues()[0]; // 讀取 B 到 G 欄
    
    // --- [V2 修正] ---
    // 原本使用 Session.getScriptTimeZone()，會導致時區轉換錯誤
    // 現在改用 ss.getSpreadsheetTimeZone() 來確保日期與您在表格中看到的完全一致
    const timeZone = ss.getSpreadsheetTimeZone();

    // 輔助函式，用於安全地格式化日期
    const formatDate = (date) => {
      if (date instanceof Date) {
        return Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
      }
      return null; // 如果不是有效的日期物件，返回 null
    };

    const details = {
      etc: formatDate(rowData[0]),          // Column B
      etd: formatDate(rowData[1]),          // Column C
      eta: formatDate(rowData[2]),          // Column D
      memo: rowData[3] || '',               // Column E
      inboundDate: formatDate(rowData[4]), // Column F
      isFinished: rowData[5] === true      // Column G
    };

    return { success: true, details: details };
  } catch (e) {
    Logger.log(`getPiDetails Error: ${e.message}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 儲存更新後的 GIT 詳細資訊。
 * @param {object} data The data object from the frontend form.
 * @returns {object} A result object indicating success or failure.
 */
function saveGitDetails(data) {
  try {
    const { piNumber, etc, etd, eta, memo, inboundDate, isFinished } = data;
    if (!piNumber) throw new Error("PI Number is missing.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gitSheet = ss.getSheetByName(GIT_SHEET_NAME);
    if (!gitSheet) throw new Error(`Sheet '${GIT_SHEET_NAME}' not found.`);

    const piColumnValues = gitSheet.getRange('A2:A' + gitSheet.getLastRow()).getValues();
    let targetRow = -1;

    for (let i = 0; i < piColumnValues.length; i++) {
      if (piColumnValues[i][0] === piNumber) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(`PI# '${piNumber}' could not be found for saving.`);
    }
    
    // 準備要寫入的資料
    // 如果日期字串為空，則寫入 null 來清空儲存格
    const valuesToSet = [
      etc ? new Date(etc) : null,
      etd ? new Date(etd) : null,
      eta ? new Date(eta) : null,
      memo,
      inboundDate ? new Date(inboundDate) : null,
      isFinished // isFinished 是布林值
    ];

    // 將更新的值寫入 B 到 G 欄
    gitSheet.getRange(targetRow, 2, 1, 6).setValues([valuesToSet]);
    SpreadsheetApp.flush();
    
    // 清除快取，以便下次打開時能獲取最新列表
    CacheService.getScriptCache().remove(GIT_CACHE_KEY);

    return { success: true, message: `Successfully updated PI# '${piNumber}'.` };
  } catch (e) {
    Logger.log(`saveGitDetails Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.toString() };
  }
}