/**
 * حسن الحوت للآلات الزراعية — واجهة برمجية (API) على Google Sheets
 * ------------------------------------------------------------------
 * انسخ هذا الكود كامل داخل: الشيت > Extensions > Apps Script
 * ثم اضغط Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * وانسخ رابط الـ /exec وأعطه للمطور ليضعه في التطبيق.
 */

var SHEET_METHODS = 'Methods';
var SHEET_TRANS = 'Transactions';

var METHOD_HEADERS = ['id', 'name', 'type', 'balance', 'color'];
var TRANS_HEADERS = ['id', 'date', 'type', 'methodId', 'methodToId', 'amount', 'category', 'description', 'timestamp'];

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // امنع جوجل من تحويل النصوص (التاريخ / المعرفات) إلى أرقام أو تواريخ
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c];
      if (h === 'date' || h === 'id' || h === 'methodId' || h === 'methodToId' || h === 'timestamp') {
        sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
      }
    }
  }
  return sh;
}

function readAll_(name, headers) {
  var sh = getSheet_(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0] === '' || row[0] === null) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    out.push(obj);
  }
  return out;
}

function appendRow_(name, headers, obj) {
  var sh = getSheet_(name, headers);
  var row = headers.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
  sh.appendRow(row);
}

function deleteById_(name, headers, id) {
  var sh = getSheet_(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
    }
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    return json_({
      ok: true,
      methods: readAll_(SHEET_METHODS, METHOD_HEADERS),
      transactions: readAll_(SHEET_TRANS, TRANS_HEADERS)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'addTransaction') {
      appendRow_(SHEET_TRANS, TRANS_HEADERS, body.payload);
    } else if (action === 'deleteTransaction') {
      deleteById_(SHEET_TRANS, TRANS_HEADERS, body.id);
    } else if (action === 'addMethod') {
      appendRow_(SHEET_METHODS, METHOD_HEADERS, body.payload);
    } else if (action === 'deleteMethod') {
      deleteById_(SHEET_METHODS, METHOD_HEADERS, body.id);
    } else if (action === 'seedMethods') {
      // يستخدم مرة واحدة لإضافة الوسائل الافتراضية لو الشيت فاضي
      var existing = readAll_(SHEET_METHODS, METHOD_HEADERS);
      if (existing.length === 0 && body.payload && body.payload.length) {
        for (var i = 0; i < body.payload.length; i++) {
          appendRow_(SHEET_METHODS, METHOD_HEADERS, body.payload[i]);
        }
      }
    } else {
      return json_({ ok: false, error: 'unknown action' });
    }

    return json_({
      ok: true,
      methods: readAll_(SHEET_METHODS, METHOD_HEADERS),
      transactions: readAll_(SHEET_TRANS, TRANS_HEADERS)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
