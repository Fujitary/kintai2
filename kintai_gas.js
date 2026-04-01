// ============================================================
// 勤怠管理 GAS スクリプト
// Google Apps Script に貼り付けて「ウェブアプリとしてデプロイ」してください
// ============================================================

// ▼ スプレッドシートのID（URLの /d/xxxxx/edit の xxxxx 部分）
const SPREADSHEET_ID = 'ここにスプレッドシートIDを貼り付け';

// ============================================================
// POST リクエスト受信（アプリ → スプレッドシート）
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'add_record') {
      return addRecord(data.record);
    } else if (action === 'delete_record') {
      return deleteRecord(data.recordId);
    } else if (action === 'sync_projects') {
      return syncProjects(data.projects);
    }

    return jsonResponse({ ok: false, error: '不明なアクション' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// GET リクエスト受信（スプレッドシート → アプリ）
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'get_records') {
      const year  = parseInt(e.parameter.year);
      const month = parseInt(e.parameter.month);
      return getRecords(year, month);
    } else if (action === 'get_projects') {
      return getProjects();
    }

    return jsonResponse({ ok: false, error: '不明なアクション' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// 記録を追加（月別シートに1行追記）
// ============================================================
function addRecord(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = getSheetName(new Date(record.start));
  const sheet = getOrCreateSheet(ss, sheetName);

  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['記録ID', 'プロジェクトID', 'プロジェクト名', '日付', '開始時刻', '終了時刻', '時間(h)', 'メモ']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#E8F0FE');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 80);
    sheet.setColumnWidth(8, 160);
  }

  const startDate = new Date(record.start);
  const endDate   = new Date(record.end);

  sheet.appendRow([
    record.id,
    record.projectId,
    record.projectName,
    Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
    Utilities.formatDate(startDate, 'Asia/Tokyo', 'HH:mm'),
    Utilities.formatDate(endDate,   'Asia/Tokyo', 'HH:mm'),
    Math.round((record.duration / 3600000) * 100) / 100,
    record.memo || ''
  ]);

  // サマリーシートを更新
  updateSummary(ss);

  return jsonResponse({ ok: true, message: '記録しました' });
}

// ============================================================
// 記録を削除（記録IDで行を検索して削除）
// ============================================================
function deleteRecord(recordId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (sheet.getName() === 'サマリー' || sheet.getName() === 'プロジェクト') continue;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === recordId) {
        sheet.deleteRow(i + 1);
        updateSummary(ss);
        return jsonResponse({ ok: true, message: '削除しました' });
      }
    }
  }

  return jsonResponse({ ok: false, error: '記録が見つかりません' });
}

// ============================================================
// 月の記録を取得してアプリへ返す
// ============================================================
function getRecords(year, month) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = `${year}年${String(month).padStart(2, '0')}月`;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonResponse({ ok: true, records: [] });
  }

  const data = sheet.getDataRange().getValues();
  const records = data.slice(1).map(row => ({
    id:          row[0],
    projectId:   row[1],
    projectName: row[2],
    date:        row[3],
    startTime:   row[4],
    endTime:     row[5],
    hours:       row[6],
    memo:        row[7]
  }));

  return jsonResponse({ ok: true, records });
}

// ============================================================
// プロジェクト一覧を同期
// ============================================================
function syncProjects(projects) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('プロジェクト');
  if (!sheet) {
    sheet = ss.insertSheet('プロジェクト', 0);
  }
  sheet.clearContents();
  sheet.appendRow(['プロジェクトID', 'プロジェクト名', 'カラー']);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
  projects.forEach(p => sheet.appendRow([p.id, p.name, p.color]));
  return jsonResponse({ ok: true });
}

// ============================================================
// プロジェクト一覧を取得
// ============================================================
function getProjects() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('プロジェクト');
  if (!sheet || sheet.getLastRow() <= 1) return jsonResponse({ ok: true, projects: [] });
  const data = sheet.getDataRange().getValues();
  const projects = data.slice(1).map(row => ({ id: row[0], name: row[1], color: row[2] }));
  return jsonResponse({ ok: true, projects });
}

// ============================================================
// サマリーシートを自動更新
// ============================================================
function updateSummary(ss) {
  let summary = ss.getSheetByName('サマリー');
  if (!summary) {
    summary = ss.insertSheet('サマリー', 0);
  }
  summary.clearContents();
  summary.appendRow(['月', 'プロジェクト', '合計時間(h)', '作業日数']);
  summary.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#E8F0FE');

  const sheets = ss.getSheets();
  const rows = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === 'サマリー' || name === 'プロジェクト') return;
    if (sheet.getLastRow() <= 1) return;

    const data = sheet.getDataRange().getValues().slice(1);
    const byProject = {};
    const days = new Set();

    data.forEach(row => {
      const pname = row[2];
      const hours = parseFloat(row[6]) || 0;
      const date  = row[3];
      byProject[pname] = (byProject[pname] || 0) + hours;
      if (date) days.add(String(date));
    });

    Object.entries(byProject).forEach(([pname, hours]) => {
      rows.push([name, pname, Math.round(hours * 100) / 100, days.size]);
    });
  });

  if (rows.length > 0) {
    summary.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  // 列幅調整
  summary.setColumnWidth(1, 110);
  summary.setColumnWidth(2, 140);
  summary.setColumnWidth(3, 110);
  summary.setColumnWidth(4, 80);
}

// ============================================================
// ユーティリティ
// ============================================================
function getSheetName(date) {
  return `${date.getFullYear()}年${String(date.getMonth() + 1).padStart(2, '0')}月`;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}
