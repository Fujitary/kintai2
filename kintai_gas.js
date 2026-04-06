// ============================================================
// 勤怠管理 GAS スクリプト v5
// NPO法人manabo-de 勤怠管理アプリ向け
//
// 【更新内容】
//   - create_spreadsheet: テンプレートからスプレッドシート新規作成
//   - create_month_sheet: 翌月シートを既存SSに追加
//   - add_record: 出勤・退勤時刻をフォーマット済みシートに書き込む
//   - delete_record: 対象行をクリア
//
// 【デプロイ設定】
//   実行ユーザー: 自分（管理者アカウント）
//   アクセス: 全員
//   ※ Drive APIを「サービス」に追加してください
// ============================================================

// ▼ テンプレートが入っているスプレッドシートのID
const TEMPLATE_SPREADSHEET_ID = '1PDWsLsetoqx7BuCdMpCRM2iSs5CH69CQvoK3FW6-HzA';

// ▼ テンプレートのシートタブ名
const TEMPLATE_SHEET_NAME = 'テンプレート';

// 列の定数（1始まり）
const COL_DATE     = 1;  // A: 月日
const COL_CLOCKIN  = 3;  // C: 出勤時間
const COL_CLOCKOUT = 4;  // D: 退勤時間
const COL_MEMO     = 12; // L: 具体的な作業内容
const COL_RECORD_ID = 13; // M: 記録ID（非表示・アプリ管理用）

// データ開始行（10行目 = 1日）
const DATA_START_ROW = 10;

// ============================================================
// POST ディスパッチャ
// ============================================================
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'create_spreadsheet') return createSpreadsheet(data);
    if (action === 'create_month_sheet') return createMonthSheet(data);
    if (action === 'add_record')         return addRecord(data.record);
    if (action === 'delete_record')      return deleteRecord(data.recordId, data.userName);

    return jsonResponse({ ok: false, error: '不明なアクション: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// GET（動作確認用）
// ============================================================
function doGet(e) {
  return jsonResponse({ ok: true, message: '勤怠GAS v5 稼働中' });
}

// ============================================================
// スプレッドシートを新規作成
//
// payload = {
//   action: 'create_spreadsheet',
//   userName: '田中太郎',
//   role: 'ユースワーカー',      // 役職（任意）
//   year: 2026,
//   month: 4
// }
// ============================================================
function createSpreadsheet(data) {
  const userName = data.userName || 'スタッフ';
  const role     = data.role     || '';
  const year     = data.year     || new Date().getFullYear();
  const month    = data.month    || (new Date().getMonth() + 1);

  // テンプレートスプレッドシートを取得
  const templateSS    = SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID);
  const templateSheet = templateSS.getSheetByName(TEMPLATE_SHEET_NAME);

  if (!templateSheet) {
    return jsonResponse({ ok: false, error: `テンプレートシート「${TEMPLATE_SHEET_NAME}」が見つかりません` });
  }

  // 新しいスプレッドシートを作成
  const newTitle = `業務日誌（${userName}）`;
  const newSS    = SpreadsheetApp.create(newTitle);
  const newSSId  = newSS.getId();

  // テンプレートシートを新SSにコピー
  const copiedSheet = templateSheet.copyTo(newSS);

  // シートタブ名を「YYYY年MM月」にリネーム
  const sheetTabName = `${year}年${String(month).padStart(2, '0')}月`;
  copiedSheet.setName(sheetTabName);

  // デフォルトで作られる「シート1」を削除
  const defaultSheet = newSS.getSheetByName('シート1');
  if (defaultSheet) newSS.deleteSheet(defaultSheet);

  // ヘッダー情報を書き込む
  writeSheetHeader(copiedSheet, userName, role, year, month);

  // スプレッドシートのURLを返す
  const url = `https://docs.google.com/spreadsheets/d/${newSSId}/edit`;

  return jsonResponse({
    ok: true,
    spreadsheetId:  newSSId,
    spreadsheetUrl: url,
    sheetName:      sheetTabName,
    message:        `スプレッドシートを作成しました: ${newTitle}`
  });
}

// ============================================================
// 既存スプレッドシートに翌月シートを追加
//
// payload = {
//   action: 'create_month_sheet',
//   spreadsheetId: '1xxx...',
//   userName: '田中太郎',
//   role: 'ユースワーカー',
//   year: 2026,
//   month: 5
// }
// ============================================================
function createMonthSheet(data) {
  const spreadsheetId = data.spreadsheetId;
  const userName      = data.userName || 'スタッフ';
  const role          = data.role     || '';
  const year          = data.year     || new Date().getFullYear();
  const month         = data.month    || (new Date().getMonth() + 1);
  const sheetTabName  = `${year}年${String(month).padStart(2, '0')}月`;

  // 対象スプレッドシートを開く
  const ss = SpreadsheetApp.openById(spreadsheetId);

  // 既に同月シートがあればそのまま返す
  const existing = ss.getSheetByName(sheetTabName);
  if (existing) {
    return jsonResponse({ ok: true, sheetName: sheetTabName, message: '既に存在します' });
  }

  // テンプレートをコピー
  const templateSS    = SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID);
  const templateSheet = templateSS.getSheetByName(TEMPLATE_SHEET_NAME);

  if (!templateSheet) {
    return jsonResponse({ ok: false, error: `テンプレートシート「${TEMPLATE_SHEET_NAME}」が見つかりません` });
  }

  const copiedSheet = templateSheet.copyTo(ss);
  copiedSheet.setName(sheetTabName);

  // ヘッダー情報を書き込む
  writeSheetHeader(copiedSheet, userName, role, year, month);

  return jsonResponse({ ok: true, sheetName: sheetTabName, message: `${sheetTabName} を作成しました` });
}

// ============================================================
// シートのヘッダー情報を書き込む
//   A3: 年月（例: 26年 4月分）
//   C5: 役職
//   C6: 氏名
// ============================================================
function writeSheetHeader(sheet, userName, role, year, month) {
  // 西暦の下2桁で表記（例: 2026 → 26）
  const shortYear = String(year).slice(-2);
  sheet.getRange('A3').setValue(`${shortYear}年 ${month}月分`);
  if (role)     sheet.getRange('C5').setValue(role);
  if (userName) sheet.getRange('C6').setValue(userName);
}

// ============================================================
// 打刻記録を書き込む
//
// record = {
//   id, userName, spreadsheetId,
//   projectName, memo,
//   start (ms timestamp), end (ms timestamp), duration (ms)
// }
// ============================================================
function addRecord(record) {
  // スプレッドシートIDが指定されている場合はそちらを使う
  const ssId = record.spreadsheetId || null;
  if (!ssId) {
    return jsonResponse({ ok: false, error: 'spreadsheetId が指定されていません' });
  }

  const ss        = SpreadsheetApp.openById(ssId);
  const startDate = new Date(record.start);
  const endDate   = new Date(record.end);

  // 対象月のシートタブを取得（なければ自動作成）
  const year         = startDate.getFullYear();
  const month        = startDate.getMonth() + 1;
  const sheetTabName = `${year}年${String(month).padStart(2, '0')}月`;
  let sheet          = ss.getSheetByName(sheetTabName);

  if (!sheet) {
    // 月シートがなければテンプレートからコピーして作成
    const result = createMonthSheet({
      spreadsheetId: ssId,
      userName:      record.userName || '',
      role:          record.role     || '',
      year,
      month
    });
    if (!result) return jsonResponse({ ok: false, error: '月シートの作成に失敗しました' });
    sheet = ss.getSheetByName(sheetTabName);
  }

  if (!sheet) {
    return jsonResponse({ ok: false, error: `シート「${sheetTabName}」が見つかりません` });
  }

  // 日付から行番号を計算（1日=10行目）
  const day = startDate.getDate();
  const row = DATA_START_ROW + day - 1;

  // 出勤・退勤時刻を書き込む
  const clockIn  = Utilities.formatDate(startDate, 'Asia/Tokyo', 'HH:mm');
  const clockOut = Utilities.formatDate(endDate,   'Asia/Tokyo', 'HH:mm');

  sheet.getRange(row, COL_CLOCKIN).setValue(clockIn);
  sheet.getRange(row, COL_CLOCKOUT).setValue(clockOut);

  // L列: プロジェクト名 + メモ
  const memoCell = sheet.getRange(row, COL_MEMO);
  const existing = String(memoCell.getValue() || '');
  const newMemo  = buildMemo(record.projectName, record.memo);
  // 同日に複数回打刻した場合は「 / 」区切りで追記
  memoCell.setValue(existing ? existing + ' / ' + newMemo : newMemo);

  // M列: 記録ID（削除時に使用）
  const idCell     = sheet.getRange(row, COL_RECORD_ID);
  const existingId = String(idCell.getValue() || '');
  idCell.setValue(existingId ? existingId + ',' + record.id : record.id);

  return jsonResponse({
    ok:      true,
    message: `書き込み完了: ${Utilities.formatDate(startDate, 'Asia/Tokyo', 'M月d日')} ${clockIn}〜${clockOut}`,
    row,
    sheet:   sheetTabName
  });
}

// ============================================================
// 記録を削除（対象行のC・D・L・M列をクリア）
// ============================================================
function deleteRecord(recordId, spreadsheetId) {
  if (!spreadsheetId) {
    return jsonResponse({ ok: false, error: 'spreadsheetId が指定されていません' });
  }

  const ss     = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) continue;

    const rowCount = lastRow - DATA_START_ROW + 1;
    const idValues = sheet.getRange(DATA_START_ROW, COL_RECORD_ID, rowCount, 1).getValues();

    for (let i = 0; i < idValues.length; i++) {
      const cellIds = String(idValues[i][0] || '');
      if (cellIds.includes(recordId)) {
        const row = DATA_START_ROW + i;
        // 記録IDを除去（複数記録がある場合は対象IDだけ除く）
        const remaining = cellIds.split(',').filter(id => id !== recordId).join(',');
        if (remaining) {
          // まだ他の記録が残っている → IDだけ更新（時刻・メモは保持）
          sheet.getRange(row, COL_RECORD_ID).setValue(remaining);
        } else {
          // この行の記録が全部消える → 行をクリア
          sheet.getRange(row, COL_CLOCKIN).clearContent();
          sheet.getRange(row, COL_CLOCKOUT).clearContent();
          sheet.getRange(row, COL_MEMO).clearContent();
          sheet.getRange(row, COL_RECORD_ID).clearContent();
        }
        return jsonResponse({ ok: true, message: '削除しました' });
      }
    }
  }

  return jsonResponse({ ok: false, error: '記録が見つかりません（ID: ' + recordId + '）' });
}

// ============================================================
// メモ文字列を組み立て
// 例: 「里山合宿：余野小学校での打ち合わせ」
// ============================================================
function buildMemo(projectName, memo) {
  const parts = [];
  if (projectName) parts.push(projectName);
  if (memo)        parts.push(memo);
  return parts.join('：');
}

// ============================================================
// ユーティリティ
// ============================================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}
