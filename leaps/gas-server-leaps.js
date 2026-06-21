// ============================================================
//  不在管理システム - Google Apps Script サーバー
//  行動予定表システムと同じ構成（GitHub Pages + GAS API + トークン認証）
//  【設定】下記2つを自分用に変更する
// ============================================================

const SHEET_ID = '1qu4z6rK2Rt2tVeXM8Jmk7eQgx1IovRpeNftlPnKC4zQ'; // 不在管理用スプレッドシート

// アクセストークン（URLを知らない第三者からのアクセスを防ぐ）
// 【設定】下のランダム文字列を自分用に変更し、HTML側のGAS_TOKENとも一致させること
const ACCESS_TOKEN = 'leaps-2026-h7vK3qN9wL';

// シート名（変更不要）
const SHEET_MASTER = '住宅マスタ';
const SHEET_LOG    = '不在管理ログ';
const SHEET_ABSENT = '現在不在中';
const SHEET_TOSEKI = '透析スケジュール';

// グレーアウト色
const COLOR_RETURNED = '#CCCCCC';
const COLOR_ABSENT   = '#FFFFFF';

// ============================================================
//  GET: データ取得系（houses / absent / dashboard / checkDup）
// ============================================================
function doGet(e) {
  try {
    if (!checkToken(e.parameter.token)) {
      return jsonResponse({ error: 'unauthorized' });
    }
    const action = e.parameter.action || '';

    if (action === 'houses')    return jsonResponse(getHouseList());
    if (action === 'absent')    return jsonResponse(getAbsentList(e.parameter.house || ''));
    if (action === 'dashboard') return jsonResponse(getDashboardData());
    if (action === 'checkDup')  return jsonResponse(checkDuplicate({
      house: e.parameter.house || '',
      name:  e.parameter.name  || ''
    }));

    return jsonResponse({ error: 'invalid action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
//  POST: 登録系（registerOut / registerIn）
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkToken(body.token)) {
      return jsonResponse({ error: 'unauthorized' });
    }
    const action = body.action;

    if (action === 'registerOut') return jsonResponse(registerOut(body.data));
    if (action === 'registerIn')  return jsonResponse(registerIn(body.data));

    return jsonResponse({ error: 'invalid action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
//  住宅マスタ読み込み（不在人数つき）
// ============================================================
function getHouseList() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateMasterSheet(ss);
  const data  = sheet.getDataRange().getValues();

  // 現在不在中シートから住宅ごとの不在人数を集計
  const absSheet = getOrCreateSheet(ss, SHEET_ABSENT);
  const absData  = absSheet.getDataRange().getValues();
  const countMap = {};
  for (let i = 1; i < absData.length; i++) {
    const h = String(absData[i][1]).trim();
    if (h) countMap[h] = (countMap[h] || 0) + 1;
  }

  const houses = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (name) houses.push({ name: name, absentCount: countMap[name] || 0 });
  }
  return houses;
}

// ============================================================
//  二重登録チェック
// ============================================================
function checkDuplicate(data) {
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID);
    const absSheet = getOrCreateSheet(ss, SHEET_ABSENT);
    const absData  = absSheet.getDataRange().getValues();

    for (let i = 1; i < absData.length; i++) {
      const house = String(absData[i][1]).trim();
      const name  = String(absData[i][2]).trim();
      if (house === data.house && name === data.name) {
        return {
          duplicate: true,
          absenceType: absData[i][3],
          outTime: absData[i][4] instanceof Date ? formatDate(absData[i][4]) : String(absData[i][4])
        };
      }
    }
    return { duplicate: false };
  } catch (err) {
    return { duplicate: false, error: err.toString() };
  }
}

// ============================================================
//  ダッシュボード用データ取得（全住宅の不在状況）
// ============================================================
function getDashboardData() {
  try {
    const ss          = SpreadsheetApp.openById(SHEET_ID);
    const absSheet     = getOrCreateSheet(ss, SHEET_ABSENT);
    const absData      = absSheet.getDataRange().getValues();
    const masterSheet  = getOrCreateMasterSheet(ss);
    const masterData   = masterSheet.getDataRange().getValues();

    const houses = [];
    for (let i = 1; i < masterData.length; i++) {
      const name = String(masterData[i][0]).trim();
      if (name) houses.push(name);
    }

    const houseMap = {};
    houses.forEach(h => { houseMap[h] = []; });

    for (let i = 1; i < absData.length; i++) {
      const row   = absData[i];
      const house = String(row[1]).trim();
      if (houseMap[house] !== undefined) {
        houseMap[house].push({
          name:           String(row[2]),
          absenceType:    String(row[3]),
          outTime:        row[4] instanceof Date ? formatDate(row[4]) : String(row[4]),
          expectedReturn: row[5] instanceof Date ? formatDate(row[5]) : String(row[5]),
          note:           String(row[6] || '')
        });
      }
    }

    const result = houses.map(h => ({
      house:       h,
      absentCount: houseMap[h].length,
      absentees:   houseMap[h]
    }));

    return { success: true, data: result, updatedAt: formatDate(new Date()) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 住宅マスタシートの初期化（初回のみ）
function getOrCreateMasterSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MASTER);
    sheet.appendRow(['住宅名', '表示順', '備考']);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#1A1A1A')
      .setFontColor('#FFFFFF');
    const samples = [
      ['鹿沼台(T01)',   1, ''],
      ['大和(T02)',     2, ''],
      ['新戸(T03)',     3, ''],
      ['上溝(T04)',     4, ''],
      ['相原(T05)',     5, ''],
      ['水郷田名(T06)', 6, ''],
      ['上田名(T07)',   7, ''],
      ['愛川中津(T08)', 8, ''],
      ['上今泉(T09)',   9, ''],
      ['原当麻(T10)',  10, ''],
      ['海老名(T11)',  11, ''],
      ['愛川(T12)',    12, ''],
      ['椿森(T13)',    13, ''],
      ['NIZIハウス(N01)', 14, '']
    ];
    samples.forEach(row => sheet.appendRow(row));
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 70);
    sheet.setColumnWidth(3, 200);
    sheet.getRange('A1').setNote('住宅名を追加・変更・削除するとフォームに即反映されます');
    SpreadsheetApp.flush();
  }
  return sheet;
}

// ============================================================
//  外出登録
// ============================================================
function registerOut(data) {
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = getOrCreateSheet(ss, SHEET_LOG);
    const absSheet = getOrCreateSheet(ss, SHEET_ABSENT);
    const now      = new Date();
    const id       = Utilities.getUuid();

    ensureLogHeader(logSheet);

    const newRow = [
      id,
      data.house,
      data.name,
      data.absenceType,
      formatDate(now),
      data.expectedReturn || '',
      '',
      '',
      data.note || '',
      '不在中'
    ];
    logSheet.appendRow(newRow);
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1, 1, newRow.length).setBackground(COLOR_ABSENT);

    if (data.absenceType === '透析') {
      const tosSheet = getOrCreateSheet(ss, SHEET_TOSEKI);
      ensureTosekiHeader(tosSheet);
      tosSheet.appendRow([id, data.house, data.name, formatDate(now), data.expectedReturn || '', '', data.note || '']);
    }

    ensureAbsentHeader(absSheet);
    absSheet.appendRow([id, data.house, data.name, data.absenceType, formatDate(now), data.expectedReturn || '', data.note || '']);

    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
//  現在不在中リスト取得
// ============================================================
function getAbsentList(house) {
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID);
    const absSheet = getOrCreateSheet(ss, SHEET_ABSENT);
    const data     = absSheet.getDataRange().getValues();

    if (data.length <= 1) return { success: true, list: [] };

    const list = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!house || row[1] === house) {
        list.push({
          id:             row[0],
          house:          row[1],
          name:           row[2],
          absenceType:    row[3],
          outTime:        row[4] instanceof Date ? formatDate(row[4]) : String(row[4]),
          expectedReturn: row[5] instanceof Date ? formatDate(row[5]) : String(row[5]),
          note:           row[6]
        });
      }
    }
    return { success: true, list: list };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
//  帰着登録
// ============================================================
function registerIn(data) {
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = getOrCreateSheet(ss, SHEET_LOG);
    const absSheet = getOrCreateSheet(ss, SHEET_ABSENT);
    const now      = new Date();

    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) {
      if (logData[i][0] === data.absentId) {
        const row = i + 1;
        logSheet.getRange(row, 7).setValue(formatDate(now));
        logSheet.getRange(row, 8).setValue(data.condition || '');
        logSheet.getRange(row, 10).setValue('帰着済み');
        logSheet.getRange(row, 1, 1, 10).setBackground(COLOR_RETURNED);
        break;
      }
    }

    if (data.absenceType === '透析') {
      const tosSheet = getOrCreateSheet(ss, SHEET_TOSEKI);
      const tosData  = tosSheet.getDataRange().getValues();
      for (let i = 1; i < tosData.length; i++) {
        if (tosData[i][0] === data.absentId) {
          tosSheet.getRange(i + 1, 6).setValue(formatDate(now));
          tosSheet.getRange(i + 1, 1, 1, 7).setBackground(COLOR_RETURNED);
          break;
        }
      }
    }

    const absData = absSheet.getDataRange().getValues();
    for (let i = absData.length - 1; i >= 1; i--) {
      if (absData[i][0] === data.absentId) {
        absSheet.deleteRow(i + 1);
        break;
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
//  ユーティリティ
// ============================================================
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureLogHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const h = ['ID', '住宅', '氏名', '不在区分', '外出日時', '予定帰着日時', '実際の帰着日時', '帰着時の状態', '備考', 'ステータス'];
  sheet.appendRow(h);
  sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#2D4A7A').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  [220,70,100,80,130,130,130,160,200,80].forEach((w,i) => sheet.setColumnWidth(i+1, w));
}

function ensureAbsentHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const h = ['ID', '住宅', '氏名', '不在区分', '外出日時', '予定帰着日時', '備考'];
  sheet.appendRow(h);
  sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#7A2D2D').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function ensureTosekiHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const h = ['ID', '住宅', '氏名', '外出日時', '予定帰着', '実際の帰着', '備考'];
  sheet.appendRow(h);
  sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#2D7A5A').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function formatDate(d) {
  if (!d || !(d instanceof Date)) return String(d || '');
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${p(d.getMonth()+1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function checkToken(token) {
  return token === ACCESS_TOKEN;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}