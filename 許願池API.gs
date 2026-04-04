/**
 * 許願池 API - Google Apps Script
 * 貼到試算表：擴充功能 → Apps Script → 新增 .gs 檔貼上後部署為「網路應用程式」
 * 重要：Sheet 第一列必須是標題 id, title, note, category, link, region, status, image1, image2, image3, supportCount, createdAt
 * 第二張分頁「集氣動態」：time, wishId, title, nick（由 addSupport 寫入；GET ?type=supportFeed 讀取）
 */

/**
 * 【必讀】試算表綁定（二選一，否則集氣動態可能無法寫入）：
 * 1) 建議：在「該本 Google 試算表」→ 擴充功能 → Apps Script → 貼上程式 → 部署。SPREADSHEET_ID 請保持空白。
 * 2) 若腳本在 script.google.com 是「獨立專案」，Web App 沒有使用中試算表，必須把 ID 填在下方。
 *    ID：試算表網址 https://docs.google.com/spreadsheets/d/【這串】/edit
 */
var SPREADSHEET_ID = "1lbAqCBnYuOkzLFyn3DWMdIbZRAm7kyvnOgILQzmMDEY";

function _getSpreadsheet_() {
  var id = String(SPREADSHEET_ID || "").replace(/^\s+|\s+$/g, "");
  if (!id) {
    try {
      id = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").replace(/^\s+|\s+$/g, "");
    } catch (e0) {}
  }
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 在編輯器執行一次即可（獨立腳本）：setSpreadsheetId_("1lbAqCBnYuOkzLFyn3DWMdIbZRAm7kyvnOgILQzmMDEY")
 */
function setSpreadsheetId_(sheetId) {
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", String(sheetId || "").trim());
}

/** 忽略分頁名稱中的空白差異 */
function _findSheetByNormalizedName_(ss, wanted) {
  var w = String(wanted).replace(/\s/g, "");
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).replace(/\s/g, "") === w) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * 取得集氣動態分頁：優先名稱「集氣動態」（含寬鬆比對），否則用第二個分頁；僅一張表時自動新增「集氣動態」
 */
function _getFeedSheet_(ss) {
  if (!ss) return null;
  var sh = ss.getSheetByName("集氣動態") || _findSheetByNormalizedName_(ss, "集氣動態");
  if (sh) return sh;
  var sheets = ss.getSheets();
  if (sheets.length > 1) {
    return sheets[1];
  }
  try {
    return ss.insertSheet("集氣動態");
  } catch (err) {
    return _findSheetByNormalizedName_(ss, "集氣動態") || ss.getSheetByName("集氣動態");
  }
}

/** 分頁完全空白時寫入標題列，否則讀取不到欄位 */
function _ensureFeedHeaders_(sh) {
  if (!sh) return;
  if (sh.getLastRow() < 1) {
    sh.appendRow(["time", "wishId", "title", "nick"]);
  }
}

/**
 * 由第一個分頁（許願列表）產生「有集氣的品項」摘要列（supportCount>0），補足集氣動態分頁未寫入時的顯示
 */
function _supportFeedFromWishSheet_(ss) {
  var out = [];
  try {
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return out;
    var headers = data[0];
    var idIdx = headers.indexOf("id");
    var titleIdx = headers.indexOf("title");
    var scIdx = headers.indexOf("supportCount");
    var createdIdx = headers.indexOf("createdAt");
    if (idIdx === -1 || scIdx === -1) return out;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var n = parseInt(row[scIdx], 10) || 0;
      if (n <= 0) continue;
      var id = String(row[idIdx] != null ? row[idIdx] : "");
      var titleStr = titleIdx !== -1 && row[titleIdx] != null ? String(row[titleIdx]) : "";
      var t = Date.now() - r * 5000;
      if (createdIdx !== -1 && row[createdIdx]) {
        var d0 = new Date(row[createdIdx]);
        if (!isNaN(d0.getTime())) t = d0.getTime();
      }
      out.push({
        time: t,
        wishId: id,
        title: titleStr,
        nick: "累計 " + n + " 集氣"
      });
    }
    out.sort(function (a, b) {
      return b.time - a.time;
    });
  } catch (e) {
    return [];
  }
  return out;
}

/** 合併：事件列優先；若某 wishId 尚無任何事件列，補上許願表的累計摘要 */
function _mergeEventFeedAndWishSummary_(eventRows, wishRows) {
  var seen = {};
  var merged = [];
  var i;
  for (i = 0; i < eventRows.length; i++) {
    seen[String(eventRows[i].wishId)] = true;
    merged.push(eventRows[i]);
  }
  for (i = 0; i < wishRows.length; i++) {
    var w = wishRows[i];
    if (!seen[String(w.wishId)]) {
      merged.push(w);
      seen[String(w.wishId)] = true;
    }
  }
  merged.sort(function (a, b) {
    return b.time - a.time;
  });
  return merged;
}

/**
 * GET：讀取許願列表。加上 ?callback=函數名 可回傳 JSONP（避開 CORS）
 * ?type=supportFeed 讀取「集氣動態」分頁（與許願列表分開）
 */
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var callback = params.callback || null;

  // type=supportFeed 或 feed=1（備用，避免舊快取／參數遺失）
  if (params.type === "supportFeed" || params.feed === "1") {
    return _getSupportFeed(callback);
  }

  // 許願列表一律讀「第一個分頁」，避免編輯器目前選在「集氣動態」時讀錯表
  var ss = _getSpreadsheet_();
  var sheet = ss.getSheets()[0];
  var data = [];
  try {
    data = sheet.getDataRange().getValues();
  } catch (err) {
    data = [];
  }
  if (!data || data.length === 0) {
    return _jsonResponse({ wishes: [] }, callback);
  }
  var headers = data[0];
  var rows = data.slice(1);
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var val = row[j];
      obj[key] = (val != null && val !== "") ? val : "";
    }
    list.push(obj);
  }
  return _jsonResponse({ wishes: list }, callback);
}

/**
 * 讀取集氣動態：① 集氣動態分頁的事件列 ② 合併許願表 supportCount>0 的摘要（無事件列的品項仍會顯示）
 */
function _getSupportFeed(callback) {
  var eventRows = [];
  try {
    var ss = _getSpreadsheet_();
    var sh = _getFeedSheet_(ss);
    if (sh) {
      var data = sh.getDataRange().getValues();
      if (data && data.length >= 2) {
        var headers = data[0];
        var timeIdx = headers.indexOf("time");
        var wishIdIdx = headers.indexOf("wishId");
        var titleIdx = headers.indexOf("title");
        var nickIdx = headers.indexOf("nick");
        if (timeIdx !== -1 && wishIdIdx !== -1) {
          for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var t = row[timeIdx];
            var timeMs = 0;
            if (typeof t === "number") {
              timeMs = t;
            } else if (t instanceof Date) {
              timeMs = t.getTime();
            } else if (t != null && t !== "") {
              var parsed = new Date(t);
              timeMs = isNaN(parsed.getTime()) ? 0 : parsed.getTime();
            }
            eventRows.push({
              time: timeMs,
              wishId: String(row[wishIdIdx] != null ? row[wishIdIdx] : ""),
              title: String(titleIdx !== -1 && row[titleIdx] != null ? row[titleIdx] : ""),
              nick: String(nickIdx !== -1 && row[nickIdx] != null && String(row[nickIdx]).trim() !== "" ? row[nickIdx] : "有人")
            });
          }
        }
      }
    }
    eventRows.sort(function (a, b) {
      return b.time - a.time;
    });
    var wishSummary = _supportFeedFromWishSheet_(ss);
    var merged = _mergeEventFeedAndWishSummary_(eventRows, wishSummary);
    if (merged.length > 300) {
      merged = merged.slice(0, 300);
    }
    return _jsonResponse({ feed: merged }, callback);
  } catch (err) {
    return _jsonResponse({ feed: [] }, callback);
  }
}

function _jsonResponse(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    var text = callback + "(" + json + ");";
    return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST：可為 (1) 上傳圖片 action=uploadImage  (2) 送出許願（JSON 或 form）
 */
function doPost(e) {
  e = e || {};
  var params = e.parameter || {};

  if (params.action === "uploadImage") {
    return _handleImageUpload(e);
  }

  var json = null;
  var returnHtml = false;

  if (params.source === "form" && params.data) {
    returnHtml = true;
    try {
      json = JSON.parse(e.parameter.data);
    } catch (err) {
      return _postResponse({ ok: false, error: "資料格式錯誤" }, returnHtml);
    }
  } else if (e && e.postData && e.postData.contents) {
    try {
      json = JSON.parse(e.postData.contents);
    } catch (err) {
      return _postResponse({ ok: false, error: err.toString() }, returnHtml);
    }
  } else if (params.data && (!params.source || params.source !== "form")) {
    try {
      json = JSON.parse(params.data);
    } catch (err) {
      return _postResponse({ ok: false, error: "資料格式錯誤" }, returnHtml);
    }
  }

  if (!json) {
    return _postResponse({ ok: false, error: "沒有收到表單資料" }, returnHtml);
  }

  // 集氣：對指定許願 +1 supportCount，寫回試算表（鎖定避免手機／電腦同時集氣寫入衝突）
  if (json.action === "addSupport") {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (lockErr) {
      return _postResponse({ ok: false, error: "系統忙碌，請稍後再試集氣" }, returnHtml);
    }
    try {
      var wishId = String(json.wishId || "").trim();
      if (!wishId) {
        return _postResponse({ ok: false, error: "缺少許願編號" }, returnHtml);
      }
      var ss = _getSpreadsheet_();
      var sheet = ss.getSheets()[0];
      var data = sheet.getDataRange().getValues();
      if (!data || data.length < 2) {
        return _postResponse({ ok: false, error: "找不到許願資料" }, returnHtml);
      }
      var headers = data[0];
      var idIdx = headers.indexOf("id");
      var supportCountIdx = headers.indexOf("supportCount");
      if (idIdx === -1 || supportCountIdx === -1) {
        return _postResponse({ ok: false, error: "試算表缺少 id 或 supportCount 欄位" }, returnHtml);
      }
      var targetRow = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idIdx]) === wishId) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) {
        return _postResponse({ ok: false, error: "找不到指定編號的許願" }, returnHtml);
      }
      var current = parseInt(data[targetRow - 1][supportCountIdx], 10) || 0;
      var newCount = current + 1;
      sheet.getRange(targetRow, supportCountIdx + 1).setValue(newCount);

      var feedSheet = _getFeedSheet_(ss);
      var feedAppended = false;
      var feedAppendError = "";
      if (feedSheet) {
        _ensureFeedHeaders_(feedSheet);
        var titleSnap = String(json.title || "").trim();
        if (titleSnap.length > 200) {
          titleSnap = titleSnap.substring(0, 200);
        }
        var nickVal = String(json.nick != null ? json.nick : "有人").trim();
        if (nickVal.length > 40) {
          nickVal = nickVal.substring(0, 40);
        }
        if (!nickVal) {
          nickVal = "有人";
        }
        // 使用 Date 物件寫入，試算表顯示為日期時間；讀取時 _getSupportFeed 會正確轉成毫秒
        var timeVal = new Date();
        try {
          feedSheet.appendRow([timeVal, wishId, titleSnap, nickVal]);
          feedAppended = true;
        } catch (appendErr) {
          feedAppendError = String(appendErr);
        }
      } else {
        feedAppendError = "找不到集氣動態分頁";
      }

      return _postResponse(
        {
          ok: true,
          supportCount: newCount,
          feedAppended: feedAppended,
          feedAppendError: feedAppendError || undefined
        },
        returnHtml
      );
    } catch (err) {
      return _postResponse({ ok: false, error: err.toString() }, returnHtml);
    } finally {
      try {
        lock.releaseLock();
      } catch (releaseErr) {}
    }
  }

  // 管理員：更新單筆許願（狀態 / 圖片）
  if (json.action === "updateWish") {
    try {
      var sheet = _getSpreadsheet_().getSheets()[0];
      var data = sheet.getDataRange().getValues();
      if (!data || data.length < 2) {
        return _postResponse({ ok: false, error: "目前沒有資料可更新" }, returnHtml);
      }
      var headers = data[0];
      var idIndex = headers.indexOf("id");
      var statusIndex = headers.indexOf("status");
      var img1Index = headers.indexOf("image1");
      var img2Index = headers.indexOf("image2");
      var img3Index = headers.indexOf("image3");
      if (idIndex === -1) {
        return _postResponse({ ok: false, error: "找不到 id 欄位" }, returnHtml);
      }
      var targetRow = -1;
      var targetId = String(json.id || "");
      for (var i = 1; i < data.length; i++) {
        var rowId = String(data[i][idIndex]);
        if (rowId === targetId) {
          targetRow = i + 1; // 轉成試算表列號（從 1 開始）
          break;
        }
      }
      if (targetRow === -1) {
        return _postResponse({ ok: false, error: "找不到指定編號的許願" }, returnHtml);
      }
      var rowRange = sheet.getRange(targetRow, 1, 1, headers.length);
      var rowValues = rowRange.getValues()[0];
      if (statusIndex !== -1 && typeof json.status === "string" && json.status !== "") {
        rowValues[statusIndex] = json.status;
      }
      if (img1Index !== -1 && typeof json.image1 === "string" && json.image1 !== "") {
        rowValues[img1Index] = json.image1;
      }
      if (img2Index !== -1 && typeof json.image2 === "string" && json.image2 !== "") {
        rowValues[img2Index] = json.image2;
      }
      if (img3Index !== -1 && typeof json.image3 === "string" && json.image3 !== "") {
        rowValues[img3Index] = json.image3;
      }
      rowRange.setValues([rowValues]);
      return _postResponse({ ok: true, id: targetId }, returnHtml);
    } catch (err) {
      return _postResponse({ ok: false, error: err.toString() }, returnHtml);
    }
  }

  try {
    var sheet = _getSpreadsheet_().getSheets()[0];
    var lastRow = sheet.getLastRow();
    var newId = (lastRow < 1) ? 1 : lastRow;
    var now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm");

    var row = [
      newId,
      json.title || "",
      json.note || "",
      json.category || "其他",
      json.link || "",
      json.region || "",
      "許願中",
      json.image1 || "",
      json.image2 || "",
      json.image3 || "",
      0,
      now
    ];
    sheet.appendRow(row);

    return _postResponse({ ok: true, id: newId }, returnHtml);
  } catch (err) {
    return _postResponse({ ok: false, error: err.toString() }, returnHtml);
  }
}

function _postResponse(obj, asHtml) {
  if (asHtml) {
    var script = "window.parent.postMessage(" + JSON.stringify(obj) + ", '*');";
    var html = "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body><script>" + script + "<\/script><\/body><\/html>";
    return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 上傳圖片：接受 (1) POST body JSON { image: "data:image/...;base64,..." } 或 (2) 表單 source=form 且參數 image= dataURL
 */
function _handleImageUpload(e) {
  var dataUrl = "";
  var returnHtml = (e.parameter && e.parameter.source === "form");

  if (e.parameter && e.parameter.source === "form" && e.parameter.image) {
    dataUrl = e.parameter.image;
  } else if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      dataUrl = body.image || "";
    } catch (err) {
      return _uploadResponse({ ok: false, error: "格式錯誤" }, returnHtml);
    }
  }

  try {
    if (!dataUrl || dataUrl.indexOf("base64,") === -1) {
      return _uploadResponse({ ok: false, error: "圖片格式錯誤" }, returnHtml);
    }
    var base64 = dataUrl.split("base64,")[1];
    if (!base64) {
      return _uploadResponse({ ok: false, error: "圖片格式錯誤" }, returnHtml);
    }
    var mime = "image/jpeg";
    var ext = "jpg";
    if (dataUrl.indexOf("image/png") !== -1) { mime = "image/png"; ext = "png"; }
    if (dataUrl.indexOf("image/gif") !== -1) { mime = "image/gif"; ext = "gif"; }
    if (dataUrl.indexOf("image/webp") !== -1) { mime = "image/webp"; ext = "webp"; }
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, "wish-" + new Date().getTime() + "." + ext);
    var folder = _getOrCreateWishFolder();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fid = file.getId();
    var url = "https://drive.google.com/thumbnail?id=" + fid + "&sz=w800";
    return _uploadResponse({ ok: true, url: url }, returnHtml);
  } catch (err) {
    return _uploadResponse({ ok: false, error: err.toString() }, returnHtml);
  }
}

function _uploadResponse(obj, asHtml) {
  if (asHtml) {
    var payload = { upload: true, ok: obj.ok, url: obj.url || "", error: obj.error || "" };
    var script = "window.parent.postMessage(" + JSON.stringify(payload) + ", '*');";
    var html = "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body><script>" + script + "<\/script><\/body><\/html>";
    return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
  }
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _getOrCreateWishFolder() {
  var name = "許願池圖片";
  var iter = DriveApp.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return DriveApp.getRootFolder().createFolder(name);
}
