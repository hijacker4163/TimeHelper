/**
 * Google Apps Script 後端程式碼 (Google 試算表資料庫)
 * 
 * 【更新說明】：
 * 1. 自動辨識第一列是否為標題列：若第一列是資料（沒有填 ID/標題），也會完整讀取第一筆資料！
 * 2. 徹底解決 Google 試算表將時間自動轉為「Sat Dec 30 1899」長字串的格式問題，統一還原為簡潔清晰的時間（如 14:59、2026/10/22）。
 * 3. 儲存時以純文字格式儲存，避免被 Google 試算表自動套用錯誤日期格式。
 * 
 * 【部署方式】：
 * 1. 複製本檔案全部內容，覆蓋到 Apps Script 編輯器中。
 * 2. 點選「testRun」執行一次以確保授權完成。
 * 3. 點擊「部署」 -> 「管理部署作業」 -> 點選鉛筆 ✏️ 編輯 -> 版本選擇「新版本」 -> 點擊「部署」即可！
 */

// 【初次授權與測試用函式】
function testRun() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log("目前試算表名稱：" + sheet.getName());
  Logger.log("✅ 成功存取試算表！Google 帳號權限已正式授權！");
}

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
      } catch (parseErr) {}
    }

    var rawData = sheet.getDataRange().getValues();
    var displayData = sheet.getDataRange().getDisplayValues();

    // 判斷第一列是否為標題列 (若第一格不是純數字，或寫著 ID/編號，則視為標題列)
    var hasHeader = false;
    if (rawData.length > 0 && rawData[0].length > 0) {
      var firstCell = String(rawData[0][0]).trim().toLowerCase();
      if (isNaN(Number(firstCell)) || firstCell === 'id') {
        hasHeader = true;
      }
    }
    
    // 動作 1：獲取所有歷史紀錄 (GET 或 action=get)
    if (!action || action === 'get') {
      if (rawData.length === 0 || (rawData.length === 1 && hasHeader)) {
        return jsonResponse({ status: 'success', data: [] });
      }
      
      var records = [];
      var startIdx = hasHeader ? 1 : 0; // 若有標題從第二列開始，若無標題從第一列完整讀取
      
      for (var i = startIdx; i < rawData.length; i++) {
        var rawRow = rawData[i];
        var dispRow = displayData[i] || [];
        if (!rawRow[0] && rawRow[0] !== 0) continue; // 略過無 ID 之空白列

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
      // 時間與日期欄位加上單引號 "'" 儲存為文字，避免 Google Sheets 自動轉型為 1899 年之 Date 物件
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
      var targetId = Number(postData && postData.id ? postData.id : params.id);
      for (var i = 0; i < rawData.length; i++) {
        if (Number(rawData[i][0]) === targetId) {
          sheet.getRange(i + 1, 8).setValue(true); // 標記第 8 欄 (已推播) 為 TRUE
          break;
        }
      }
      return jsonResponse({ status: 'success' });
    }
    
    // 動作 4：刪除單筆紀錄 (action=delete)
    if (action === 'delete') {
      var targetId = Number(postData && postData.id ? postData.id : params.id);
      for (var i = 0; i < rawData.length; i++) {
        if (Number(rawData[i][0]) === targetId) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ status: 'success', deleted_id: targetId });
    }
    
    // 動作 5：清空所有紀錄 (action=clear)
    if (action === 'clear') {
      var lastRow = sheet.getLastRow();
      if (hasHeader && lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      } else if (!hasHeader && lastRow >= 1) {
        sheet.clear();
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

// --- 輔助函式：時間與日期精準格式化 ---

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

// 輔助函式：回傳 JSON 回應與正確 MIME 類型
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
