/**
 * 許願池 API - Google Apps Script
 * 貼到試算表：擴充功能 → Apps Script → 新增 .gs 檔貼上後部署為「網路應用程式」
 * 重要：Sheet 第一列必須是標題 id, title, note, category, link, region, status, image1, image2, image3, supportCount, createdAt
 * 第二張分頁「集氣動態」：time, wishId, title, nick（由 addSupport 寫入；GET ?type=supportFeed 讀取）
 */

/**
 * GET：讀取許願列表。加上 ?callback=函數名 可回傳 JSONP（避開 CORS）
 * ?type=supportFeed 讀取「集氣動態」分頁（與許願列表分開）
 */
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var callback = params.callback || null;

  if (params.type === "supportFeed") {
    return _getSupportFeed(callback);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
 * 讀取第二張分頁「集氣動態」：欄位 time, wishId, title, nick；回傳最新 30 筆（新→舊）
 */
function _getSupportFeed(callback) {
  var rows = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("集氣動態");
    if (!sh) {
      return _jsonResponse({ feed: [] }, callback);
    }
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) {
      return _jsonResponse({ feed: [] }, callback);
    }
    var headers = data[0];
    var timeIdx = headers.indexOf("time");
    var wishIdIdx = headers.indexOf("wishId");
    var titleIdx = headers.indexOf("title");
    var nickIdx = headers.indexOf("nick");
    if (timeIdx === -1 || wishIdIdx === -1) {
      return _jsonResponse({ feed: [] }, callback);
    }
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
      rows.push({
        time: timeMs,
        wishId: String(row[wishIdIdx] != null ? row[wishIdIdx] : ""),
        title: String(titleIdx !== -1 && row[titleIdx] != null ? row[titleIdx] : ""),
        nick: String(nickIdx !== -1 && row[nickIdx] != null && String(row[nickIdx]).trim() !== "" ? row[nickIdx] : "有人")
      });
    }
    rows.sort(function (a, b) {
      return b.time - a.time;
    });
    if (rows.length > 30) {
      rows = rows.slice(0, 30);
    }
  } catch (err) {
    rows = [];
  }
  return _jsonResponse({ feed: rows }, callback);
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
  }

  if (!json) {
    return _postResponse({ ok: false, error: "沒有收到表單資料" }, returnHtml);
  }

  // 集氣：對指定許願 +1 supportCount，寫回試算表
  if (json.action === "addSupport") {
    try {
      var wishId = String(json.wishId || "").trim();
      if (!wishId) {
        return _postResponse({ ok: false, error: "缺少許願編號" }, returnHtml);
      }
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

      var feedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("集氣動態");
      if (feedSheet) {
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
        var timeVal = new Date().getTime();
        feedSheet.appendRow([timeVal, wishId, titleSnap, nickVal]);
      }

      return _postResponse({ ok: true, supportCount: newCount }, returnHtml);
    } catch (err) {
      return _postResponse({ ok: false, error: err.toString() }, returnHtml);
    }
  }

  // 管理員：更新單筆許願（狀態 / 圖片）
  if (json.action === "updateWish") {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
