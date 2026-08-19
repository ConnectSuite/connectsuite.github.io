// ============================================================
//  行動予定表（arch）GAS サーバー
//  スプレッドシートのシート構成：
//    「予定」「request」「information」「staff」
// ============================================================

const SHEET_ID     = '1rsMvJKaKETBwW_3jSqtNkPPKCpdOIphbmDFiPT7Y0as';
const ACCESS_TOKEN = 'arch-2026-x7kP9mQ2vL';

const SHEET_SCHEDULE = '予定';
const SHEET_REQUEST  = 'request';
const SHEET_INFO     = 'information';
const SHEET_STAFF    = 'staff';

// ============================================================
//  エントリポイント
// ============================================================
function doGet(e) {
  try {
    if (!checkToken(e.parameter.token)) return jsonResponse({ error: 'unauthorized' });
    if (e.parameter.action === 'get') return jsonResponse(getAllData());
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken(body.token)) return jsonResponse({ error: 'unauthorized' });
    const { action } = body;

    if (action === 'addSchedule')    return jsonResponse(addSchedule(body.entry));
    if (action === 'editSchedule')   return jsonResponse(editRow(SHEET_SCHEDULE, body.id, body.patch));
    if (action === 'deleteSchedule') return jsonResponse(deleteRow(SHEET_SCHEDULE, body.id));

    if (action === 'addRequest')     return jsonResponse(addRequest(body.entry));
    if (action === 'editRequest')    return jsonResponse(editRow(SHEET_REQUEST, body.id, body.patch));
    if (action === 'assignStaff')    return jsonResponse(editRow(SHEET_REQUEST, body.id, { staff: body.staff }));
    if (action === 'deleteRequest')  return jsonResponse(deleteRow(SHEET_REQUEST,  body.id));

    if (action === 'addInfo')        return jsonResponse(addInfo(body.entry));
    if (action === 'editInfo')       return jsonResponse(editRow(SHEET_INFO, body.id, body.patch));
    if (action === 'deleteInfo')     return jsonResponse(deleteRow(SHEET_INFO,     body.id));

    if (action === 'addStaff')       return jsonResponse(addStaff(body.entry));
    if (action === 'editStaff')      return jsonResponse(editRow(SHEET_STAFF, body.id, body.patch));
    if (action === 'deleteStaff')    return jsonResponse(deleteRow(SHEET_STAFF, body.id));
    if (action === 'bulkUpdateStaffOrder') return jsonResponse(bulkUpdateStaffOrder(body.updates));

    return jsonResponse({ error: 'unknown action' });
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
//  全データ取得
// ============================================================
function getAllData() {
  // SpreadsheetApp.openById() は呼び出しごとに数百ms〜数秒かかる重い処理のため、
  // 4シート分まとめて読む際は1回だけ開いて使い回す（以前は4回開いていて遅かった）
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    schedules: getSheetData(ss, SHEET_SCHEDULE),
    requests:  getSheetData(ss, SHEET_REQUEST),
    infos:     getSheetData(ss, SHEET_INFO),
    staff:     getSheetData(ss, SHEET_STAFF),
  };
}

// ============================================================
//  予定
// ============================================================
function addSchedule(entry) {
  const sheet = getOrCreateSheet(SHEET_SCHEDULE, ['id','floor','name','dest','return','start','dateStr','updatedAt']);
  appendRowAsText(sheet, [
    entry.id, entry.floor, entry.name, entry.dest,
    entry.return, entry.start || '',
    entry.dateStr, entry.updatedAt
  ]);
  return { ok: true };
}

// ============================================================
//  ご依頼（request）
//  新規登録時はstaff（担当者）を空で保存する運用
// ============================================================
function addRequest(entry) {
  const sheet = getOrCreateSheet(SHEET_REQUEST, ['id','client','time','purpose','staff','dateStr','updatedAt']);
  appendRowAsText(sheet, [
    entry.id, entry.client, entry.time,
    entry.purpose, entry.staff || '',
    entry.dateStr, entry.updatedAt
  ]);
  return { ok: true };
}

// ============================================================
//  お知らせ（information）
// ============================================================
function addInfo(entry) {
  const sheet = getOrCreateSheet(SHEET_INFO, ['id','text','createdAt']);
  appendRowAsText(sheet, [entry.id, entry.text, entry.createdAt]);
  return { ok: true };
}

// ============================================================
//  従業員マスタ（staff）
//  列: id, last(苗字), first(名前), floor(フロア), order(表示順)
// ============================================================
function addStaff(entry) {
  const sheet = getOrCreateSheet(SHEET_STAFF, ['id','last','first','floor','order']);
  appendRowAsText(sheet, [entry.id, entry.last, entry.first || '', entry.floor, entry.order]);
  return { ok: true };
}

// 表示順の一括更新（並べ替え時、複数行のorderをまとめて書き換える）
// updates: [{id, order}, ...]
function bulkUpdateStaffOrder(updates) {
  const sheet = getOrCreateSheet(SHEET_STAFF, ['id','last','first','floor','order']);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol    = headers.indexOf('id');
  const orderCol = headers.indexOf('order');
  const orderMap = {};
  updates.forEach(u => { orderMap[String(u.id)] = u.order; });

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][idCol]);
    if (orderMap.hasOwnProperty(rowId)) {
      sheet.getRange(i + 1, orderCol + 1).setValue(orderMap[rowId]);
    }
  }
  return { ok: true };
}

// ============================================================
//  汎用：行の一部フィールドを更新
//  patch: 更新したい列だけを持つオブジェクト（例: {dest:'新行先', return:'12:00'}）
// ============================================================
function editRow(sheetName, id, patch) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false };
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      Object.keys(patch).forEach(key => {
        const col = headers.indexOf(key);
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(patch[key]);
      });
      return { ok: true };
    }
  }
  return { ok: false };
}

// ============================================================
//  汎用：行削除
// ============================================================
function deleteRow(sheetName, id) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ============================================================
//  汎用：シートデータ取得（オブジェクト配列）
// ============================================================
function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = formatCellValue(row[i]); });
    return obj;
  });
}

// セルの値を文字列化する。Dateオブジェクトの場合は "YYYY-M-D" 形式（先頭ゼロなし）に変換する。
// ※スプレッドシートが "2026-8-10" のような文字列を自動的に日付型へ変換してしまうケースへの対策。
function formatCellValue(val) {
  if (val === null || val === undefined) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return `${val.getFullYear()}-${val.getMonth()+1}-${val.getDate()}`;
  }
  return String(val);
}

// ============================================================
//  ユーティリティ
// ============================================================
function checkToken(token) {
  return token === ACCESS_TOKEN;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    appendRowAsText(sheet, headers);
  }
  return sheet;
}

// 行を追加する前に、書き込み先セル範囲をテキスト形式(@)にしてから値を入れる。
// これにより "2026-8-10" のような日付に見える文字列が、
// スプレッドシート側で自動的にDate型へ変換されるのを防ぐ。
function appendRowAsText(sheet, rowValues) {
  const rowIndex = sheet.getLastRow() + 1;
  const range = sheet.getRange(rowIndex, 1, 1, rowValues.length);
  range.setNumberFormat('@');
  range.setValues([rowValues.map(v => v === undefined || v === null ? '' : String(v))]);
}

function setTextFormat(sheet) {
  const range = sheet.getDataRange();
  range.setNumberFormat('@');
}
