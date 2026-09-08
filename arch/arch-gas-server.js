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
    if (action === 'assignStaff')    return jsonResponse(editRow(SHEET_REQUEST, body.id, { staff: body.staff, appVersion: body.appVersion }));
    if (action === 'deleteRequest')  return jsonResponse(deleteRow(SHEET_REQUEST,  body.id));

    if (action === 'addInfo')        return jsonResponse(addInfo(body.entry));
    if (action === 'editInfo')       return jsonResponse(editRow(SHEET_INFO, body.id, body.patch));
    if (action === 'deleteInfo')     return jsonResponse(deleteRow(SHEET_INFO,     body.id));

    if (action === 'addStaff')       return jsonResponse(addStaff(body.entry));
    if (action === 'editStaff')      return jsonResponse(editRow(SHEET_STAFF, body.id, body.patch));
    if (action === 'deleteStaff')    return jsonResponse(deleteRow(SHEET_STAFF, body.id));
    if (action === 'bulkUpdateStaffOrder') return jsonResponse(bulkUpdateStaffOrder(body.updates));

    if (action === 'migrateAddVersionColumn') return jsonResponse(migrateAddVersionColumn());

    // 分析レポートのテスト送信用（本運用の週次/月次自動送信は別途トリガー設定）
    if (action === 'sendTestReport') return jsonResponse(sendReportEmail(body.to, body.days || 7));
    // 月次レポート（月間＋通算の2通）を今すぐ手動で送る（動作確認用。本番は毎月1日にトリガーで自動実行）
    if (action === 'sendMonthlyReportsNow') return jsonResponse(sendMonthlyReports());
    // setupMonthlyTriggerで登録したトリガーが実在するか確認する（読み取り専用）
    if (action === 'listTriggers') return jsonResponse(listReportTriggers());

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
  // 実際のシートの列順は id/floor/name/dest/start/return/dateStr/updatedAt
  // （過去に「start」列を手動追加した経緯があり、この並びが正）。
  // 以前はここが return→start の順で書き込んでおり、新規登録のたびに
  // 出発・帰着が列レベルですり替わって保存される不具合の真因だった（2026年9月修正）。
  const sheet = getOrCreateSheet(SHEET_SCHEDULE, ['id','floor','name','dest','start','return','dateStr','updatedAt','appVersion']);
  appendRowAsText(sheet, [
    entry.id, entry.floor, entry.name, entry.dest,
    entry.start || '', entry.return,
    entry.dateStr, entry.updatedAt, entry.appVersion || ''
  ]);
  return { ok: true };
}

// ============================================================
//  ご依頼（request）
//  新規登録時はstaff（担当者）を空で保存する運用
// ============================================================
function addRequest(entry) {
  const sheet = getOrCreateSheet(SHEET_REQUEST, ['id','client','time','purpose','staff','dateStr','updatedAt','appVersion']);
  appendRowAsText(sheet, [
    entry.id, entry.client, entry.time,
    entry.purpose, entry.staff || '',
    entry.dateStr, entry.updatedAt, entry.appVersion || ''
  ]);
  return { ok: true };
}

// ============================================================
//  お知らせ（information）
// ============================================================
function addInfo(entry) {
  const sheet = getOrCreateSheet(SHEET_INFO, ['id','text','createdAt','appVersion']);
  appendRowAsText(sheet, [entry.id, entry.text, entry.createdAt, entry.appVersion || '']);
  return { ok: true };
}

// ============================================================
//  従業員マスタ（staff）
//  列: id, last(苗字), first(名前), floor(フロア), order(表示順)
// ============================================================
function addStaff(entry) {
  const sheet = getOrCreateSheet(SHEET_STAFF, ['id','last','first','floor','order','appVersion']);
  appendRowAsText(sheet, [entry.id, entry.last, entry.first || '', entry.floor, entry.order, entry.appVersion || '']);
  return { ok: true };
}

// 表示順の一括更新（並べ替え時、複数行のorderをまとめて書き換える）
// updates: [{id, order}, ...]
function bulkUpdateStaffOrder(updates) {
  const sheet = getOrCreateSheet(SHEET_STAFF, ['id','last','first','floor','order','appVersion']);
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
//  移行：既存シートに appVersion 列を追加する（1回実行すれば十分、既に列が
//  ある場合は何もしない）。バグ調査時に「どのバージョンで登録されたデータか」
//  を突き止められるようにするため、2026年9月に追加した。
// ============================================================
function migrateAddVersionColumn() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  [SHEET_SCHEDULE, SHEET_REQUEST, SHEET_INFO, SHEET_STAFF].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf('appVersion') !== -1) return;
    const newCol = lastCol + 1;
    const cell = sheet.getRange(1, newCol);
    cell.setNumberFormat('@');
    cell.setValue('appVersion');
  });
  return { ok: true };
}

// ============================================================
//  活動レポート（週次/月次で経営層向けに送るための集計＋メール送信）
//  2026年9月、実績が溜まってきたことを受けてプロトタイプとして追加。
//  行先欄は自由入力で1回の外出で複数箇所を回ることがあるため、
//  「、」「，」「・」「→」「全角/半角スペース」で区切って個別の行先として
//  分割集計する。ただし「休み」等の非行先ワードや日付の断片は除外する。
// ============================================================
const REPORT_EXCLUDE_WORDS = ['休み','帰宅','直行','在席','外出中','テレワーク','同行','会議室','リハ','リハーサル'];

function splitDestForReport(dest) {
  if (!dest) return [];
  return dest.split(/[、，・→\s]+/)
    .map(s => s.trim())
    .filter(Boolean)
    .filter(w => !/^\d{1,2}\/\d{1,2}$/.test(w)) // "9/7" のような日付断片を除外
    .filter(w => !REPORT_EXCLUDE_WORDS.some(ex => w.indexOf(ex) !== -1));
}

function parseDateStrToDate(s) {
  const [y, m, d] = String(s).split('-').map(Number);
  return new Date(y, m - 1, d);
}

// 集計本体（日付範囲指定版）。fromDate/toDateは両端を含む。
// pendingScope: 'range' なら入電がこの期間内のもの、'all' なら全期間の未対応を対象にする
function buildReportDataForRange(fromDate, toDate, pendingScope) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const schedules = getSheetData(ss, SHEET_SCHEDULE);
  const requests  = getSheetData(ss, SHEET_REQUEST);

  const today = new Date(); today.setHours(0,0,0,0);

  const targetSchedules = schedules.filter(s => {
    if (!s.dateStr) return false;
    const dt = parseDateStrToDate(s.dateStr);
    return dt >= fromDate && dt <= toDate;
  });

  // 人別集計：件数・合計外出時間（時刻指定ありの分のみ集計）
  const byPerson = {};
  targetSchedules.forEach(s => {
    if (!byPerson[s.name]) byPerson[s.name] = { count: 0, totalMin: 0, timedCount: 0 };
    byPerson[s.name].count++;
    if (/^\d{2}:\d{2}$/.test(s.start) && /^\d{2}:\d{2}$/.test(s.return)) {
      const [sh, sm] = s.start.split(':').map(Number);
      const [rh, rm] = s.return.split(':').map(Number);
      const dur = (rh * 60 + rm) - (sh * 60 + sm);
      if (dur > 0) { byPerson[s.name].totalMin += dur; byPerson[s.name].timedCount++; }
    }
  });
  const personRows = Object.keys(byPerson).map(name => {
    const p = byPerson[name];
    return {
      name, count: p.count,
      avgMin: p.timedCount ? Math.round(p.totalMin / p.timedCount) : null,
      totalMin: p.totalMin,
    };
  }).sort((a, b) => b.count - a.count);

  // 行先別集計（分割済み）
  const byDest = {};
  targetSchedules.forEach(s => {
    splitDestForReport(s.dest).forEach(d => { byDest[d] = (byDest[d] || 0) + 1; });
  });
  const destRows = Object.entries(byDest).map(([dest, count]) => ({ dest, count }))
    .sort((a, b) => b.count - a.count).slice(0, 15);

  // 曜日別件数
  const DOW = ['日','月','火','水','木','金','土'];
  const byDow = [0,0,0,0,0,0,0];
  targetSchedules.forEach(s => { byDow[parseDateStrToDate(s.dateStr).getDay()]++; });

  // ご依頼：未対応（担当者未定）一覧。滞留日数も併記
  // pendingScope='range'なら対象期間に入電したもの限定、'all'なら全期間の未対応をすべて出す
  const pendingRequests = requests.filter(r => {
      if (r.staff) return false;
      if (pendingScope !== 'range') return true;
      if (!r.dateStr) return false;
      const dt = parseDateStrToDate(r.dateStr);
      return dt >= fromDate && dt <= toDate;
    })
    .map(r => {
      const dt = r.dateStr ? parseDateStrToDate(r.dateStr) : null;
      const daysAgo = dt ? Math.round((today - dt) / 86400000) : null;
      return { client: r.client, time: r.time, purpose: r.purpose, dateStr: r.dateStr, daysAgo };
    })
    .sort((a, b) => (b.daysAgo || 0) - (a.daysAgo || 0));

  const fmtDate = (d) => `${d.getFullYear()}-${d.getMonth()+1}-${d.getDate()}`;
  return {
    fromStr: fmtDate(fromDate), toStr: fmtDate(toDate),
    totalCount: targetSchedules.length,
    personRows, destRows, byDow, DOW, pendingRequests,
  };
}

// 直近n日版（テスト送信用に残す）
function buildReportData(days) {
  const today = new Date(); today.setHours(0,0,0,0);
  const cutoff = new Date(today.getTime() - (days - 1) * 86400000);
  return buildReportDataForRange(cutoff, today, 'all');
}

// 指定した年月（1〜12）の月間レポート用データ
function buildMonthlyReportData(year, month) {
  const from = new Date(year, month - 1, 1);
  const to = new Date(year, month, 0); // 月末日
  return buildReportDataForRange(from, to, 'range');
}

// 運用開始日〜今日までの通算レポート用データ
function buildAllTimeReportData() {
  const from = new Date(2026, 7, 18); // 2026-08-18 本番運用開始日
  const today = new Date(); today.setHours(0,0,0,0);
  return buildReportDataForRange(from, today, 'all');
}

function buildReportHtml(data, title, pendingSectionLabel) {
  title = title || '行動予定表 活動レポート';
  pendingSectionLabel = pendingSectionLabel || '■ ご依頼：未対応の滞留状況';
  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const th = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;font-size:13px;color:#666;"';
  const td = 'style="padding:6px 10px;border-bottom:1px solid #eee;font-size:14px;"';
  const tdR = 'style="padding:6px 10px;border-bottom:1px solid #eee;font-size:14px;text-align:right;"';

  const personTable = data.personRows.map(p => `
    <tr>
      <td ${td}>${esc(p.name)}</td>
      <td ${tdR}>${p.count}件</td>
      <td ${tdR}>${p.avgMin != null ? p.avgMin + '分' : '－'}</td>
    </tr>`).join('');

  const destTable = data.destRows.map(d => `
    <tr><td ${td}>${esc(d.dest)}</td><td ${tdR}>${d.count}件</td></tr>`).join('');

  const dowTable = data.DOW.map((dow, i) => `
    <tr><td ${td}>${dow}曜日</td><td ${tdR}>${data.byDow[i]}件</td></tr>`).join('');

  const pendingTable = data.pendingRequests.length === 0
    ? `<p style="font-size:14px;color:#666;">未対応のご依頼はありません。</p>`
    : `<table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
        <tr><th ${th}>クライアント</th><th ${th}>要件</th><th ${th}>入電日</th><th ${th}>経過日数</th></tr>
        ${data.pendingRequests.map(r => `
          <tr>
            <td ${td}>${esc(r.client)}</td>
            <td ${td}>${esc(r.purpose)}</td>
            <td ${td}>${esc(r.dateStr)}</td>
            <td ${tdR} ${r.daysAgo >= 3 ? 'color:#c00;font-weight:bold;' : ''}>${r.daysAgo != null ? r.daysAgo + '日' : '－'}</td>
          </tr>`).join('')}
      </table>`;

  const spanDays = Math.round((new Date(data.toStr) - new Date(data.fromStr)) / 86400000) + 1;
  const dowCaveat = spanDays < 90
    ? `<p style="color:#999;font-size:12px;">※期間が短いため、曜日別の傾向はまだ参考値としてご覧ください。</p>`
    : '';

  return `
  <div style="font-family:'Hiragino Sans','Noto Sans JP',sans-serif;color:#222;max-width:640px;">
    <h2 style="margin-bottom:4px;">${esc(title)}</h2>
    <p style="color:#666;font-size:13px;margin-top:0;">対象期間：${data.fromStr} 〜 ${data.toStr}／行動予定 合計${data.totalCount}件</p>

    <h3 style="margin-bottom:6px;">${esc(pendingSectionLabel)}</h3>
    ${pendingTable}

    <h3 style="margin-bottom:6px;margin-top:20px;">■ 人別 外出件数・平均外出時間</h3>
    <table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
      <tr><th ${th}>氏名</th><th ${th} style="text-align:right;">外出件数</th><th ${th} style="text-align:right;">平均外出時間</th></tr>
      ${personTable}
    </table>

    <h3 style="margin-bottom:6px;margin-top:20px;">■ 行先別 頻度（上位15件・複数訪問先は分割集計）</h3>
    <table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
      <tr><th ${th}>行先</th><th ${th} style="text-align:right;">件数</th></tr>
      ${destTable}
    </table>

    <h3 style="margin-bottom:6px;margin-top:20px;">■ 曜日別 外出件数（参考値）</h3>
    <table style="border-collapse:collapse;width:100%;margin-bottom:8px;">
      ${dowTable}
    </table>
    ${dowCaveat}

    <p style="color:#999;font-size:11px;margin-top:24px;">このメールはConnectSuite 行動予定表システムから自動生成されました。</p>
  </div>`;
}

// 初回のみ：メール送信権限を承認するための手動実行用関数。
// スクリプトエディタでこの関数を選んで「実行」ボタンを押すと、Googleの権限確認画面が
// 出るので承認する。承認後、自分宛てに確認メールが届けば以後は自動送信も動くようになる。
function authorizeMailPermission() {
  const me = Session.getActiveUser().getEmail();
  MailApp.sendEmail(me, '【権限確認】行動予定表レポート機能', 'このメールが届いていれば、メール送信の権限承認は完了です。');
}

function sendReportEmail(toAddress, days) {
  const data = buildReportData(days || 7);
  const html = buildReportHtml(data);
  MailApp.sendEmail({
    to: toAddress,
    subject: `【テスト】行動予定表 活動レポート（直近${days || 7}日）`,
    htmlBody: html,
  });
  return { ok: true, totalCount: data.totalCount };
}

// 月次レポートの送付先。今はまこと様のみ、今後増やす場合はここにカンマ区切りで追加する
const REPORT_RECIPIENTS = ['makoto@aoki-prt.co.jp'];

// 「月間レポート」（先月分）と「通算レポート」（運用開始日〜今日）の2通を送る。
// 毎月1日にトリガーで自動実行する想定（setupMonthlyTriggerで設定）。
function sendMonthlyReports() {
  const now = new Date();
  // 実行日の前月を対象にする（1日に実行される想定なので、前月が「先月分」になる）
  const targetMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const y = targetMonthDate.getFullYear(), m = targetMonthDate.getMonth() + 1;

  const monthlyData = buildMonthlyReportData(y, m);
  const monthlyHtml = buildReportHtml(monthlyData, `月間レポート ${y}年${m}月分`, `■ ご依頼：${m}月受付分で未対応のもの`);

  const allTimeData = buildAllTimeReportData();
  const allTimeHtml = buildReportHtml(allTimeData, '通算レポート', '■ ご依頼：現時点で未対応のもの（全期間）');

  const to = REPORT_RECIPIENTS.join(',');
  MailApp.sendEmail({ to, subject: `月間レポート ${y}年${m}月分`, htmlBody: monthlyHtml });
  MailApp.sendEmail({ to, subject: `通算レポート（${allTimeData.fromStr} 〜 ${allTimeData.toStr}）`, htmlBody: allTimeHtml });
  return { ok: true, month: `${y}-${m}`, monthlyCount: monthlyData.totalCount, allTimeCount: allTimeData.totalCount };
}

// 初回のみ：「毎月1日 朝9時」にsendMonthlyReportsを自動実行するトリガーを設定する。
// スクリプトエディタでこの関数を選んで「実行」すると、日付ベーストリガーの権限確認が出るので承認する。
// 既に同名のトリガーがある場合は一旦削除してから作り直す（重複登録防止）。
function listReportTriggers() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendMonthlyReports')
    .map(t => ({
      handlerFunction: t.getHandlerFunction(),
      eventType: String(t.getEventType()),
    }));
  return { ok: true, count: triggers.length, triggers };
}

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReports') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendMonthlyReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  return { ok: true };
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
