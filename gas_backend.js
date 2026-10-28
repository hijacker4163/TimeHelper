/**
 * Google Apps Script 後端程式碼 (Google 試算表資料庫 + LINE Bot 自動推播)
 * 
 * ====================================================================
 * ⚙️ 【安全性設計：金鑰與環境變數 (防 GitHub 外洩)】
 * ====================================================================
 * 本程式碼優先由 Google 雲端的「指令碼屬性 (Script Properties)」讀取金鑰。
 * 本地檔案可完全「留空」，推送到 GitHub 也不會有任何金鑰洩漏風險！
 * 
 * 💡 如何在 Google 雲端設定金鑰（只需設定一次）：
 * 1. 進入 Google Apps Script，點擊左側 ⚙️「專案設定」
 * 2. 往下滑找到「指令碼屬性」-> 點擊「新增指令碼屬性」：
 *    - 屬性：LINE_CHANNEL_ACCESS_TOKEN，值：[您的 LINE Access Token]
 *    - 屬性：LINE_USER_ID，值：[您的 LINE User ID]
 * 3. 點擊「儲存指令碼屬性」即完成！
 */

// 從 Google Apps Script 雲端環境變數讀取金鑰；此處留空防洩漏
var scriptProperties = PropertiesService.getScriptProperties();
var LINE_CHANNEL_ACCESS_TOKEN = scriptProperties.getProperty("LINE_CHANNEL_ACCESS_TOKEN") || "";
var LINE_USER_ID = scriptProperties.getProperty("LINE_USER_ID") || "";

/**
 * ====================================================================
 * 🚀 【測試與建立背景定時檢查排程】
 * ====================================================================
 * 1. 設定好指令碼屬性後，在頂端函式選單選「testSendLine」點擊「執行」：
 *    - 若手機收到 LINE 測試訊息，代表雲端金鑰讀取與 LINE 設定成功！
 * 2. 接著在下拉選單選擇「setupAutoNotifyTrigger」點擊「執行」：
 *    - 系統會自動建立 24 小時每分鐘背景檢查排程。
 *    - 就算手機電腦關閉網頁，時間到 Google 雲端也會準時傳送 LINE 通知！
 * 3. 部署 Web App：右上角「部署」->「管理部署作業」-> 編輯 -> 選擇「新版本」->「部署」。
 */

// 【重要！解鎖外部連線權限函式】
// 請在頂端選單選擇「authorizeExternalRequest」並點擊「執行」
// Google 就會立刻跳出「需要授權」視窗，點擊「審查權限 ➔ 帳號 ➔ 進階 ➔ 前往(不安全) ➔ 允許」即可解鎖！
function authorizeExternalRequest() {
  UrlFetchApp.fetch("https://api.line.me");
}

// 【測試函式 1】：驗證 LINE Bot 是否能成功發送訊息到您的手機
function testSendLine() {
  sendLineMessage("⏱️【剩餘時間計算器】LINE Bot 測試訊息！\n看到這則訊息代表 LINE 推播設定完全成功！🎉");
}

// 【測試函式 2】：測試直接發送試算表中「最後一筆紀錄」的 LINE 推播
function testNotifyLastRecord() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  if (rawData.length <= 0) {
    Logger.log("⚠️ 試算表中沒有任何資料");
    return;
  }
  var lastIdx = rawData.length - 1;
  var taskName = String(displayData[lastIdx][2] || rawData[lastIdx][2] || '').trim();
  var resultTime = formatCellTime(displayData[lastIdx][4], rawData[lastIdx][4]);
  var inputStr = String(displayData[lastIdx][3] || rawData[lastIdx][3] || '').trim();
  
  var msg = "⏱️【時間到了！】(測試)\n" +
            (taskName ? "📌 事項：" + taskName + "\n" : "") +
            "⏰ 到期時間：" + resultTime + "\n" +
            "⌛ 倒數時長：" + inputStr;
  sendLineMessage(msg);
  Logger.log("✅ 已成功推送最後一筆資料至 LINE！");
}

// 【設定函式 3】：一鍵建立每分鐘自動於 Google 背景檢查到期的排程
function setupAutoNotifyTrigger() {
  // 清除舊有的定時排程避免重複
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDueRecordsAndNotify') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 建立每 1 分鐘自動執行一次的觸發器
  ScriptApp.newTrigger('checkDueRecordsAndNotify')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("✅ 成功建立每分鐘自動排程！未來時間到達時，Google 將主動傳 LINE 給您。");
}

// 【初次授權與測試試算表讀取】
function testRun() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log("目前試算表名稱：" + sheet.getName());
  Logger.log("✅ 成功存取試算表！Google 帳號權限已正式授權！");
}

// 【手動執行 4】：手動清除試算表中所有「已到期」的紀錄（保留進行中的倒數）
function manualClearExpiredRecords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rawData = sheet.getDataRange().getValues();
  if (rawData.length <= 0) {
    Logger.log("⚠️ 試算表中沒有任何資料");
    return;
  }
  var hasHeader = false;
  var firstCell = String(rawData[0][0]).trim().toLowerCase();
  if (isNaN(Number(firstCell)) || firstCell === 'id') {
    hasHeader = true;
  }
  var startIdx = hasHeader ? 1 : 0;
  var now = Date.now();
  var deletedCount = 0;
  
  // 由後往前倒序刪除，避免列號移位
  for (var i = rawData.length - 1; i >= startIdx; i--) {
    var targetTs = Number(rawData[i][6]);
    if (targetTs && targetTs <= now) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  Logger.log("✅ 已手動清除 " + deletedCount + " 筆已到期紀錄！（進行中的倒數已安全保留）");
}

// 【手動執行 5】：手動清空試算表中「所有」紀錄（保留第一列標題）
function manualClearAllRecords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rawData = sheet.getDataRange().getValues();
  if (rawData.length <= 0) {
    Logger.log("⚠️ 試算表中沒有任何資料");
    return;
  }
  var hasHeader = false;
  var firstCell = String(rawData[0][0]).trim().toLowerCase();
  if (isNaN(Number(firstCell)) || firstCell === 'id') {
    hasHeader = true;
  }
  var startIdx = hasHeader ? 1 : 0;
  var deletedCount = 0;
  for (var i = rawData.length - 1; i >= startIdx; i--) {
    sheet.deleteRow(i + 1);
    deletedCount++;
  }
  Logger.log("✅ 已手動清空試算表共 " + deletedCount + " 筆資料！");
}

/**
 * ====================================================================
 * ⏰ 【背景定時觸發器】：每分鐘由 Google 雲端自動執行
 * ====================================================================
 */
function checkDueRecordsAndNotify() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rawData = sheet.getDataRange().getValues();
    var displayData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 0) return;
    
    var now = Date.now();
    var hasHeader = false;
    var firstCell = String(rawData[0][0]).trim().toLowerCase();
    if (isNaN(Number(firstCell)) || firstCell === 'id') {
      hasHeader = true;
    }
    
    var startIdx = hasHeader ? 1 : 0;
    
    for (var i = startIdx; i < rawData.length; i++) {
      var targetTs = Number(rawData[i][6]);
      var notified = Boolean(rawData[i][7]);
      
      // 目標時間已到達且尚未標記已通知
      if (targetTs && targetTs <= now && !notified) {
        var taskName = String(displayData[i][2] || rawData[i][2] || '').trim();
        var resultTime = formatCellTime(displayData[i][4], rawData[i][4]);
        var inputStr = String(displayData[i][3] || rawData[i][3] || '').trim();
        
        var message = "";
        if (taskName) {
          message = "⏱️【時間到了！】\n" +
                    "📌 事項：" + taskName + "\n" +
                    "⏰ 到期時間：" + resultTime + "\n" +
                    "⌛ 倒數時長：" + inputStr;
        } else {
          message = "⏱️【時間到了！】\n" +
                    "⏰ 到期時間：" + resultTime + "\n" +
                    "⌛ 倒數時長：" + inputStr;
        }
        
        // 發送 LINE Bot 訊息到手機
        sendLineMessage(message);
        
        // 標記試算表第 8 欄 (已推播) 為 TRUE
        sheet.getRange(i + 1, 8).setValue(true);
      }
    }
  } catch (err) {
    Logger.log("checkDueRecordsAndNotify 錯誤: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * 透過 LINE Messaging API 發送 Push Message
 */
function sendLineMessage(text) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN.indexOf("請在此填入") !== -1) {
    Logger.log("⚠️ 尚未設定 LINE_CHANNEL_ACCESS_TOKEN");
    return;
  }
  if (!LINE_USER_ID || LINE_USER_ID.indexOf("請在此填入") !== -1) {
    Logger.log("⚠️ 尚未設定 LINE_USER_ID");
    return;
  }
  
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    to: LINE_USER_ID,
    messages: [{
      type: 'text',
      text: text
    }]
  };
  
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code === 200) {
      Logger.log("✅ LINE 訊息發送成功！");
    } else {
      Logger.log("❌ LINE 發送失敗 (" + code + "): " + res.getContentText());
    }
  } catch (e) {
    Logger.log("發送 LINE 請求異常: " + e.toString());
  }
}

/**
 * ====================================================================
 * 🌐 【Web App 請求處理】：前端讀寫 API
 * ====================================================================
 */
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var params = e ? e.parameter : {};
    var action = params ? params.action : null;
    var postData = null;
    
    // 解析 POST 內容 (支援 JSON 與 text/plain)
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
        if (!action && postData.action) {
          action = postData.action;
        }
      } catch (parseErr) {}
    }

    var rawData = sheet.getDataRange().getValues();
    var displayData = sheet.getDataRange().getDisplayValues();

    // 判斷第一列是否為標題列
    var hasHeader = false;
    if (rawData.length > 0 && rawData[0].length > 0) {
      var firstCell = String(rawData[0][0]).trim().toLowerCase();
      if (isNaN(Number(firstCell)) || firstCell === 'id') {
        hasHeader = true;
      }
    }
    var startIdx = hasHeader ? 1 : 0;
    
    // 動作 1：獲取所有歷史紀錄 (GET 或 action=get)
    if (!action || action === 'get') {
      if (rawData.length === 0 || (rawData.length === 1 && hasHeader)) {
        return jsonResponse({ status: 'success', data: [] });
      }
      
      var records = [];
      
      for (var i = startIdx; i < rawData.length; i++) {
        var rawRow = rawData[i];
        var dispRow = displayData[i] || [];
        if (!rawRow[0] && rawRow[0] !== 0) continue;

        records.push({
          id: Number(rawRow[0]),
          timestamp: formatCellTimestamp(dispRow[1], rawRow[1]),
          task: String(dispRow[2] || rawRow[2] || ''),
          input: String(dispRow[3] || rawRow[3] || ''),
          resultTime: formatCellTime(dispRow[4], rawRow[4]),
          resultDate: formatCellDate(dispRow[5], rawRow[5]),
          targetTimestamp: Number(rawRow[6]),
          notified: Boolean(rawRow[7])
        });
      }
      return jsonResponse({ status: 'success', data: records });
    }
    
    // 動作 2：新增紀錄 (action=add)
    if (action === 'add') {
      var r = postData && postData.data ? postData.data : JSON.parse(params.data || '{}');
      sheet.appendRow([
        r.id,
        "'" + r.timestamp,
        r.task || '',
        r.input,
        "'" + r.resultTime,
        "'" + r.resultDate,
        r.targetTimestamp,
        false
      ]);
      return jsonResponse({ status: 'success', data: r });
    }
    
    // 動作 3：標記紀錄為已推播 (action=notify)
    if (action === 'notify') {
      // 關鍵修復：當前端倒數時間到達時，立即呼叫 LINE 發送推播，無需苦等每分鐘排程！
      checkDueRecordsAndNotify();
      return jsonResponse({ status: 'success' });
    }
    
    // 動作 4：刪除單筆紀錄 (action=delete)
    if (action === 'delete') {
      var targetId = postData && postData.id ? postData.id : params.id;
      var targetIdStr = String(targetId).trim();
      var deleted = false;
      for (var i = rawData.length - 1; i >= startIdx; i--) {
        var cellIdStr = String(rawData[i][0]).trim();
        if (cellIdStr === targetIdStr || (targetIdStr && Number(cellIdStr) === Number(targetIdStr))) {
          sheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
      return jsonResponse({ status: 'success', deleted: deleted, deleted_id: targetId });
    }
    
    // 動作 5：只清空已過期的紀錄 (action=clear 或 clear_expired)，保留進行中的倒數！
    if (action === 'clear' || action === 'clear_expired') {
      var checkTime = Number(postData && postData.now ? postData.now : (params.now || Date.now()));
      var deletedCount = 0;
      // 由最後一列倒序往前刪除，避免刪除列時造成的 index 錯位
      for (var i = rawData.length - 1; i >= startIdx; i--) {
        var targetTs = Number(rawData[i][6]);
        // 若該筆紀錄的目標時間已小於或等於當前時間 (已到期)，則刪除
        if (targetTs && targetTs <= checkTime) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
      return jsonResponse({ status: 'success', deleted_count: deletedCount });
    }

    // 動作 6：清空所有紀錄 (action=clear_all)
    if (action === 'clear_all') {
      var deletedCount = 0;
      for (var i = rawData.length - 1; i >= startIdx; i--) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
      return jsonResponse({ status: 'success', deleted_count: deletedCount });
    }

    // 動作 7：編輯備忘事項 (action=update_task 或 edit_task)
    if (action === 'update_task' || action === 'edit_task') {
      var targetId = postData && postData.id ? postData.id : params.id;
      var newTask = postData && (postData.task !== undefined) ? postData.task : (params.task !== undefined ? params.task : '');
      var targetIdStr = String(targetId).trim();
      var updated = false;
      for (var i = startIdx; i < rawData.length; i++) {
        var cellIdStr = String(rawData[i][0]).trim();
        if (cellIdStr === targetIdStr || (targetIdStr && Number(cellIdStr) === Number(targetIdStr))) {
          // 試算表第 3 欄 (C 欄) 為事項備忘文字 (task)
          sheet.getRange(i + 1, 3).setValue(String(newTask).trim());
          updated = true;
          break;
        }
      }
      return jsonResponse({ status: 'success', updated: updated, task: newTask });
    }
    
    return jsonResponse({ status: 'error', message: '未知操作: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// 輔助函式：時間與日期格式化
function formatCellTime(dispVal, rawVal) {
  if (rawVal instanceof Date) {
    var h = String(rawVal.getHours()).padStart(2, '0');
    var m = String(rawVal.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  var s = String(dispVal || rawVal || '').trim();
  if (s.indexOf('1899') !== -1 || s.indexOf('GMT') !== -1) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      var h = String(d.getHours()).padStart(2, '0');
      var m = String(d.getMinutes()).padStart(2, '0');
      return h + ':' + m;
    }
  }
  return s || '--:--';
}

function formatCellDate(dispVal, rawVal) {
  if (rawVal instanceof Date) {
    return rawVal.getFullYear() + '/' + (rawVal.getMonth() + 1) + '/' + rawVal.getDate();
  }
  var s = String(dispVal || rawVal || '').trim();
  if (s.indexOf('GMT') !== -1) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
    }
  }
  return s;
}

function formatCellTimestamp(dispVal, rawVal) {
  if (rawVal instanceof Date) {
    var m = String(rawVal.getMonth() + 1);
    var d = String(rawVal.getDate());
    var h = String(rawVal.getHours()).padStart(2, '0');
    var min = String(rawVal.getMinutes()).padStart(2, '0');
    return m + '/' + d + ' ' + h + ':' + min;
  }
  var s = String(dispVal || rawVal || '').trim();
  if (s.indexOf('GMT') !== -1) {
    var dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      var m = String(dt.getMonth() + 1);
      var d = String(dt.getDate());
      var h = String(dt.getHours()).padStart(2, '0');
      var min = String(dt.getMinutes()).padStart(2, '0');
      return m + '/' + d + ' ' + h + ':' + min;
    }
  }
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
