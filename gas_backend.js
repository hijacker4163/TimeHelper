/**
 * Google Apps Script 後端程式碼 (Google 試算表資料庫)
 * 
 * 【使用步驟】：
 * 1. 建立一個新的 Google 試算表 (Google Sheets)。
 * 2. 在第一列 (Row 1) 依序填入欄位名稱（選填，方便您辨識）：
 *    A1: ID | B1: 建立時間 | C1: 事項備忘 | D1: 倒數長度 | E1: 目標時間 | F1: 目標日期 | G1: 目標時間戳記 | H1: 已推播
 * 3. 點擊頂端選單的「擴充功能」 -> 「Apps Script」。
 * 4. 將下方所有程式碼複製貼上到 Apps Script 編輯器中，覆蓋原有內容。
 * 5. 點擊右上角「部署」 -> 「新增部署作業」：
 *    - 類型選擇：「網頁應用程式 (Web App)」
 *    - 說明：時間計算器 API
 *    - 執行身分：「我」
 *    - 誰可以存取：「所有人 (Anyone)」 (※ 重要！這樣 GitHub Pages 與手機才能讀寫)
 * 6. 點擊「部署」，授權 Google 權限，複製產生的「網頁應用程式網址 (Web App URL)」。
 * 7. 將該網址貼入前端頁面的「⚙️ 雲端同步設定」中即可！
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // 使用 LockService 防止多裝置並行寫入衝突
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var params = e ? e.parameter : {};
    var action = params ? params.action : null;
    var postData = null;
    
    // 解析 POST 內容 (支援 JSON 與 text/plain 以避免 CORS 預檢限制)
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
        if (!action && postData.action) {
          action = postData.action;
        }
      } catch (parseErr) {
        // 若無法解析為 JSON，繼續嘗試參數
      }
    }
    
    // 動作 1：獲取所有歷史紀錄 (GET 或 action=get)
    if (!action || action === 'get') {
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return jsonResponse({ status: 'success', data: [] });
      }
      
      var records = [];
      // 從第二列開始讀取資料 (第一列為標題)
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0]) continue; // 略過空白列
        records.push({
          id: Number(row[0]),
          timestamp: String(row[1] || ''),
          task: String(row[2] || ''),
          input: String(row[3] || ''),
          resultTime: String(row[4] || ''),
          resultDate: String(row[5] || ''),
          targetTimestamp: Number(row[6]),
          notified: Boolean(row[7])
        });
      }
      return jsonResponse({ status: 'success', data: records });
    }
    
    // 動作 2：新增紀錄 (action=add)
    if (action === 'add') {
      var r = postData && postData.data ? postData.data : JSON.parse(params.data || '{}');
      sheet.appendRow([
        r.id,
        r.timestamp,
        r.task || '',
        r.input,
        r.resultTime,
        r.resultDate,
        r.targetTimestamp,
        false
      ]);
      return jsonResponse({ status: 'success', data: r });
    }
    
    // 動作 3：標記紀錄為已推播 (action=notify)
    if (action === 'notify') {
      var targetId = Number(postData && postData.id ? postData.id : params.id);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (Number(data[i][0]) === targetId) {
          sheet.getRange(i + 1, 8).setValue(true); // 標記第 8 欄 (已推播) 為 TRUE
          break;
        }
      }
      return jsonResponse({ status: 'success' });
    }
    
    // 動作 4：刪除單筆紀錄 (action=delete)
    if (action === 'delete') {
      var targetId = Number(postData && postData.id ? postData.id : params.id);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (Number(data[i][0]) === targetId) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ status: 'success', deleted_id: targetId });
    }
    
    // 動作 5：清空所有紀錄 (action=clear)
    if (action === 'clear') {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return jsonResponse({ status: 'success' });
    }
    
    return jsonResponse({ status: 'error', message: '未知操作: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// 輔助函式：回傳 JSON 回應與正確 MIME 類型
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
