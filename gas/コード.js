// ============================================================
// eikaiwa-vocab クラウド保存用 GAS（こづかいノート方式）v3
// v2に「最終更新時刻(lastModified)」の記録を追加。データ構造は不変。
// ============================================================

var SHEET_NAME = 'words';
var HEADERS = ['id','term','pos','meaning','example','status','stage','inBox','boxStreak','source','addedAt','nextDue','history'];
var META_SHEET = '_meta';   // ★追加：時刻などを記録する裏方シート

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

// ★追加：時刻を保存する裏方シートを用意する
function getMetaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(META_SHEET);
  if (!sh) sh = ss.insertSheet(META_SHEET);
  return sh;
}
function getLastModified_() {
  var v = Number(getMetaSheet_().getRange('A1').getValue());
  return v || 0;
}
function setLastModified_(t) {
  getMetaSheet_().getRange('A1').setValue(Number(t) || Date.now());
}

// 読み込み要求（GET）。callbackパラメータがあればJSONPで返す。
function doGet(e) {
  var result;
  try {
    var sh = getSheet_();
    var lastRow = sh.getLastRow();
    var words = [];
    if (lastRow >= 2) {
      var values = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var w = {};
        for (var c = 0; c < HEADERS.length; c++) {
          w[HEADERS[c]] = row[c];
        }
        w.stage = Number(w.stage) || 0;
        w.boxStreak = Number(w.boxStreak) || 0;
        w.addedAt = Number(w.addedAt) || 0;
        w.nextDue = Number(w.nextDue) || 0;
        w.inBox = (w.inBox === true || w.inBox === 'true' || w.inBox === 'TRUE');
        try { w.history = w.history ? JSON.parse(w.history) : []; } catch (e2) { w.history = []; }
        if (w.term) words.push(w);
      }
    }
    result = { ok: true, words: words, count: words.length, lastModified: getLastModified_() }; // ★時刻を返す
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return output_(e, result);
}

// 保存要求（POST）
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var words = body.words || [];
    // ★端末から送られた時刻があればそれを、無ければサーバ時刻を採用
    var lastModified = Number(body.lastModified) || Date.now();
    var sh = getSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      sh.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
    }
    if (words.length > 0) {
      var rows = [];
      for (var i = 0; i < words.length; i++) {
        var w = words[i];
        var row = [];
        for (var c = 0; c < HEADERS.length; c++) {
          var key = HEADERS[c];
          var v = w[key];
          if (key === 'history') {
            v = JSON.stringify(v || []);
          } else if (v === undefined || v === null) {
            v = '';
          }
          row.push(v);
        }
        rows.push(row);
      }
      sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    }
    setLastModified_(lastModified); // ★保存した時刻を記録
    result = { ok: true, saved: words.length, lastModified: lastModified }; // ★時刻を返す
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return output_(e, result);
}

// callback があれば JSONP（text/javascript）で、無ければ通常JSONで返す
function output_(e, obj) {
  var jsonStr = JSON.stringify(obj);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + jsonStr + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}