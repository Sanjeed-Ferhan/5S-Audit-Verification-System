/**
 * WorkCenter Map — Google Sheet sync backend
 * ------------------------------------------
 * SETUP (one time)
 * 1. Create a Google Sheet (or open an existing one).
 * 2. Extensions > Apps Script. Delete the starter code and paste this file.
 * 3. Set TOKEN below to a long random value (or set SYNC_TOKEN in
 *    Project Settings > Script properties). Use the same token in the app.
 * 4. Deploy > New deployment > Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    Deploy, authorize, and copy the URL ending in /exec into the app.
 *
 * The app creates two sheets automatically:
 *   WorkCenters  — active work centers
 *   _Deleted     — deletion tombstones so every device removes them
 */

var TOKEN = 'PASTE_A_LONG_RANDOM_TOKEN_HERE';
var SHEET_NAME = 'WorkCenters';
var DELETED_SHEET = '_Deleted';
var HEADERS = ['id', 'sbu', 'name', 'code', 'zone', 'category', 'lat', 'lng',
               'acc', 'notes', 'createdAt', 'updatedAt'];
var DELETED_HEADERS = ['id', 'updatedAt'];


function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    if (!checkToken_(p.token)) throw new Error('Invalid token. Check TOKEN in the script and the app settings.');
    var action = p.action || 'ping';
    if (action === 'ping') {
      var sh = getSheet_();
      result = { ok: true, sheet: sh.getName(), rows: Math.max(0, sh.getLastRow() - 1) };
    } else if (action === 'list') {
      result = { ok: true, workcenters: listCenters_(), deleted: listDeleted_() };
    } else if (action === 'upsert') {
      upsertCenter_(JSON.parse(p.record || '{}'));
      result = { ok: true };
    } else if (action === 'delete') {
      deleteCenter_(p.id, p.updatedAt);
      result = { ok: true };
    } else {
      throw new Error('Unknown action: ' + action);
    }
  } catch (err) {
    result = { ok: false, error: errMessage_(err) };
  }
  return respond_(result, p.callback);
}


function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) { /* handled below */ }
  var result;
  try {
    if (!checkToken_(body.token)) throw new Error('Invalid token. Check TOKEN in the script and the app settings.');
    var action = body.action;
    if (action === 'upsert') {
      upsertCenter_(body.record || {});
      result = { ok: true };
    } else if (action === 'delete') {
      deleteCenter_(body.id, body.updatedAt);
      result = { ok: true };
    } else if (action === 'upsertMany') {
      var records = body.records || [];
      records.forEach(upsertCenter_);
      result = { ok: true, count: records.length };
    } else {
      throw new Error('Unknown action: ' + action);
    }
  } catch (err) {
    result = { ok: false, error: errMessage_(err) };
  }
  return respond_(result, null);
}


/* ------------------------------------------------------------------ */
/* responses                                                           */
/* ------------------------------------------------------------------ */

function respond_(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function errMessage_(err) {
  return String((err && err.message) || err || 'Unknown error');
}


/* ------------------------------------------------------------------ */
/* auth                                                                */
/* ------------------------------------------------------------------ */

function checkToken_(given) {
  var expected = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN') || TOKEN;
  if (!expected || expected === 'PASTE_A_LONG_RANDOM_TOKEN_HERE') return false;
  return String(given || '') === String(expected);
}


/* ------------------------------------------------------------------ */
/* sheets                                                              */
/* ------------------------------------------------------------------ */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  ensureHeader_(sh, HEADERS);
  return sh;
}

function getDeletedSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DELETED_SHEET);
  if (!sh) sh = ss.insertSheet(DELETED_SHEET);
  ensureHeader_(sh, DELETED_HEADERS);
  return sh;
}

function ensureHeader_(sh, headers) {
  if (sh.getLastRow() > 0) return;
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');
}


/* ------------------------------------------------------------------ */
/* reads                                                               */
/* ------------------------------------------------------------------ */

function listCenters_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var rec = {};
    for (var h = 0; h < HEADERS.length; h++) rec[HEADERS[h]] = row[h];
    rec.lat = toNumber_(rec.lat);
    rec.lng = toNumber_(rec.lng);
    rec.acc = toNumber_(rec.acc);
    rec.updatedAt = toNumber_(rec.updatedAt);
    rec.createdAt = toNumber_(rec.createdAt);
    out.push(rec);
  }
  return out;
}

function listDeleted_() {
  var sh = getDeletedSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, DELETED_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ id: String(values[i][0]), updatedAt: toNumber_(values[i][1]) });
  }
  return out;
}


/* ------------------------------------------------------------------ */
/* writes                                                              */
/* ------------------------------------------------------------------ */

function upsertCenter_(rec) {
  if (!rec || !rec.id) throw new Error('Missing work center id');
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = getSheet_();
    var rowIndex = findRow_(sh, rec.id);
    var row = [
      String(rec.id),
      safeText_(rec.sbu || 'ACCL'),
      safeText_(rec.name || ''),
      safeText_(rec.code || ''),
      safeText_(rec.zone || ''),
      String(rec.category || 'other'),
      toNumber_(rec.lat),
      toNumber_(rec.lng),
      toNumber_(rec.acc),
      safeText_(rec.notes || ''),
      toNumber_(rec.createdAt),
      toNumber_(rec.updatedAt) || Date.now()
    ];
    if (rowIndex > 0) {
      sh.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    removeTombstone_(rec.id);
  } finally {
    lock.releaseLock();
  }
}

function deleteCenter_(id, updatedAt) {
  if (!id) throw new Error('Missing work center id');
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = getSheet_();
    var rowIndex = findRow_(sh, id);
    if (rowIndex > 0) sh.deleteRow(rowIndex);
    addTombstone_(id, toNumber_(updatedAt) || Date.now());
  } finally {
    lock.releaseLock();
  }
}


/* ------------------------------------------------------------------ */
/* helpers                                                             */
/* ------------------------------------------------------------------ */

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function addTombstone_(id, updatedAt) {
  var sh = getDeletedSheet_();
  var last = sh.getLastRow();
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.getRange(i + 2, 1, 1, DELETED_HEADERS.length).setValues([[String(id), updatedAt]]);
        return;
      }
    }
  }
  sh.appendRow([String(id), updatedAt]);
}

function removeTombstone_(id) {
  var sh = getDeletedSheet_();
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      return;
    }
  }
}

function toNumber_(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = Number(v);
  return isFinite(n) ? n : '';
}

function safeText_(v) {
  var s = String(v === null || v === undefined ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
