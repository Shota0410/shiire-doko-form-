// ===== セットアップ手順 =====
// 1. Google スプレッドシートを新規作成する
// 2. 拡張機能 > Apps Script を開く
// 3. このファイルの内容を全てコピー&ペーストする
// 4. 「プロジェクトの設定」>「スクリプトプロパティ」で以下を追加:
//    ADMIN_KEY : (任意の管理者パスワード文字列)
// 5. 「デプロイ」>「新しいデプロイ」>「種類: ウェブアプリ」を選択
//    実行ユーザー: 自分, アクセスできるユーザー: 全員 に設定してデプロイ
// 6. 表示されたWeb App URLをindex.htmlの「GAS_URL」変数に貼り付ける
// 7. index.htmlと同じADMIN_KEYを管理ページのパスワードとして使用する
// ==========================

// ===== スプレッドシート・シート取得ユーティリティ =====

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function initSheets() {
  var ss = getSpreadsheet();

  var scheduleHeaders = ['ID', '日付', '時刻', '場所・集合場所', '定員', '説明・備考', '作成日時', '担当者'];
  var applicationHeaders = ['申込ID', 'スケジュールID', '申込日時', 'お名前', 'Instagram ID', 'LINE ID', '電話番号', 'メールアドレス', '備考'];

  var scheduleSheet = getOrCreateSheet(ss, 'スケジュール', scheduleHeaders);
  var applicationSheet = getOrCreateSheet(ss, '申込者', applicationHeaders);

  return { scheduleSheet: scheduleSheet, applicationSheet: applicationSheet };
}

// ===== UUID生成 =====

function generateUUID() {
  return Utilities.getUuid();
}

// ===== CORS対応レスポンス生成 =====

function createResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ===== 管理者認証 =====

function checkAdminKey(key) {
  var adminKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!adminKey) {
    // ADMIN_KEYが未設定の場合はアクセス拒否
    return false;
  }
  return key === adminKey;
}

// ===== スケジュール一覧取得 =====

function getSchedules() {
  var sheets = initSheets();
  var scheduleSheet = sheets.scheduleSheet;
  var applicationSheet = sheets.applicationSheet;

  var scheduleData = scheduleSheet.getDataRange().getValues();
  var applicationData = applicationSheet.getDataRange().getValues();

  if (scheduleData.length <= 1) {
    return { success: true, data: [] };
  }

  // 申込数カウント用マップ
  var countMap = {};
  for (var i = 1; i < applicationData.length; i++) {
    var sid = applicationData[i][1];
    if (sid) {
      countMap[sid] = (countMap[sid] || 0) + 1;
    }
  }

  var schedules = [];
  for (var j = 1; j < scheduleData.length; j++) {
    var row = scheduleData[j];
    var id = row[0];
    if (!id) continue;

    var capacity = Number(row[4]) || 0;
    var applied = countMap[id] || 0;
    var remaining = capacity - applied;

    schedules.push({
      id: id,
      date: row[1],
      time: row[2],
      location: row[3],
      capacity: capacity,
      description: row[5],
      createdAt: row[6],
      staff: row[7] || '',
      applied: applied,
      remaining: remaining < 0 ? 0 : remaining
    });
  }

  // 日付昇順でソート
  schedules.sort(function(a, b) {
    var da = new Date(String(a.date).replace(/\//g, '-') + 'T' + (a.time || '00:00'));
    var db = new Date(String(b.date).replace(/\//g, '-') + 'T' + (b.time || '00:00'));
    return da - db;
  });

  return { success: true, data: schedules };
}

// ===== 申込者一覧取得（管理者専用） =====

function getApplications(adminKey, scheduleId) {
  if (!checkAdminKey(adminKey)) {
    return { success: false, message: '認証エラー: 管理者キーが正しくありません' };
  }

  var sheets = initSheets();
  var applicationSheet = sheets.applicationSheet;
  var scheduleSheet = sheets.scheduleSheet;

  var applicationData = applicationSheet.getDataRange().getValues();
  var scheduleData = scheduleSheet.getDataRange().getValues();

  // スケジュールIDと名前のマップ
  var scheduleMap = {};
  for (var i = 1; i < scheduleData.length; i++) {
    var sRow = scheduleData[i];
    if (sRow[0]) {
      scheduleMap[sRow[0]] = sRow[1] + ' ' + sRow[2] + ' ' + sRow[3];
    }
  }

  var applications = [];
  for (var j = 1; j < applicationData.length; j++) {
    var row = applicationData[j];
    if (!row[0]) continue;

    // scheduleIdフィルタ
    if (scheduleId && row[1] !== scheduleId) continue;

    applications.push({
      id: row[0],
      scheduleId: row[1],
      scheduleName: scheduleMap[row[1]] || row[1],
      appliedAt: row[2],
      name: row[3],
      instagram: row[4],
      line: row[5],
      phone: row[6],
      email: row[7],
      note: row[8]
    });
  }

  // 申込日時降順
  applications.sort(function(a, b) {
    return new Date(b.appliedAt) - new Date(a.appliedAt);
  });

  return { success: true, data: applications };
}

// ===== 申し込み登録 =====

function apply(params) {
  var scheduleId = params.scheduleId;
  var name = params.name;

  if (!scheduleId || !name) {
    return { success: false, message: 'スケジュールIDとお名前は必須です' };
  }

  var sheets = initSheets();
  var scheduleSheet = sheets.scheduleSheet;
  var applicationSheet = sheets.applicationSheet;

  // スケジュール存在確認
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var targetSchedule = null;
  for (var i = 1; i < scheduleData.length; i++) {
    if (scheduleData[i][0] === scheduleId) {
      targetSchedule = scheduleData[i];
      break;
    }
  }

  if (!targetSchedule) {
    return { success: false, message: '指定されたスケジュールが見つかりません' };
  }

  // 定員チェック
  var capacity = Number(targetSchedule[4]) || 0;
  var applicationData = applicationSheet.getDataRange().getValues();
  var appliedCount = 0;
  for (var j = 1; j < applicationData.length; j++) {
    if (applicationData[j][1] === scheduleId) {
      appliedCount++;
    }
  }

  if (capacity > 0 && appliedCount >= capacity) {
    return { success: false, message: 'このスケジュールは満席です' };
  }

  // 申し込み登録
  var now = new Date();
  var appliedAt = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var id = generateUUID();

  applicationSheet.appendRow([
    id,
    scheduleId,
    appliedAt,
    params.name || '',
    params.instagram || '',
    params.line || '',
    params.phone || '',
    params.email || '',
    params.note || ''
  ]);

  return { success: true, message: 'お申し込みを受け付けました', data: { id: id } };
}

// ===== 日程追加（管理者専用） =====

function addSchedule(params) {
  if (!checkAdminKey(params.adminKey)) {
    return { success: false, message: '認証エラー: 管理者キーが正しくありません' };
  }

  var date = params.date;
  var time = params.time;
  var location = params.location;
  var capacity = params.capacity;

  if (!date || !location || !capacity) {
    return { success: false, message: '日付・場所・定員は必須です' };
  }

  var sheets = initSheets();
  var scheduleSheet = sheets.scheduleSheet;

  var now = new Date();
  var createdAt = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var id = generateUUID();

  scheduleSheet.appendRow([
    id,
    date,
    time || '',
    location,
    Number(capacity) || 0,
    params.description || '',
    createdAt,
    params.staff || ''
  ]);

  return { success: true, message: 'スケジュールを追加しました', data: { id: id } };
}

// ===== 日程更新（管理者専用） =====

function updateSchedule(params) {
  if (!checkAdminKey(params.adminKey)) {
    return { success: false, message: '認証エラー: 管理者キーが正しくありません' };
  }

  var id = params.id;
  if (!id) {
    return { success: false, message: 'IDは必須です' };
  }

  var sheets = initSheets();
  var scheduleSheet = sheets.scheduleSheet;
  var scheduleData = scheduleSheet.getDataRange().getValues();

  var targetRow = -1;
  for (var i = 1; i < scheduleData.length; i++) {
    if (scheduleData[i][0] === id) {
      targetRow = i + 1; // シートの行番号（1始まり）
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: '指定されたスケジュールが見つかりません' };
  }

  // 更新（IDと作成日時は変更しない）
  scheduleSheet.getRange(targetRow, 2).setValue(params.date || scheduleData[targetRow - 1][1]);
  scheduleSheet.getRange(targetRow, 3).setValue(params.time !== undefined ? params.time : scheduleData[targetRow - 1][2]);
  scheduleSheet.getRange(targetRow, 4).setValue(params.location || scheduleData[targetRow - 1][3]);
  scheduleSheet.getRange(targetRow, 5).setValue(params.capacity !== undefined ? Number(params.capacity) : scheduleData[targetRow - 1][4]);
  scheduleSheet.getRange(targetRow, 6).setValue(params.description !== undefined ? params.description : scheduleData[targetRow - 1][5]);
  scheduleSheet.getRange(targetRow, 8).setValue(params.staff !== undefined ? params.staff : (scheduleData[targetRow - 1][7] || ''));

  return { success: true, message: 'スケジュールを更新しました' };
}

// ===== 日程削除（管理者専用） =====

function deleteSchedule(params) {
  if (!checkAdminKey(params.adminKey)) {
    return { success: false, message: '認証エラー: 管理者キーが正しくありません' };
  }

  var id = params.id;
  if (!id) {
    return { success: false, message: 'IDは必須です' };
  }

  var sheets = initSheets();
  var scheduleSheet = sheets.scheduleSheet;
  var applicationSheet = sheets.applicationSheet;

  // スケジュール行を削除（後ろから検索して削除）
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var deleted = false;
  for (var i = scheduleData.length - 1; i >= 1; i--) {
    if (scheduleData[i][0] === id) {
      scheduleSheet.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }

  if (!deleted) {
    return { success: false, message: '指定されたスケジュールが見つかりません' };
  }

  // 関連申込者を削除（後ろから検索して削除）
  var applicationData = applicationSheet.getDataRange().getValues();
  for (var j = applicationData.length - 1; j >= 1; j--) {
    if (applicationData[j][1] === id) {
      applicationSheet.deleteRow(j + 1);
    }
  }

  return { success: true, message: 'スケジュールと関連する申し込みを削除しました' };
}

// ===== doGet: GETリクエストハンドラ =====

function doGet(e) {
  var params = e.parameter || {};
  var action = params.action || '';

  var result;

  try {
    if (action === 'getSchedules') {
      result = getSchedules();
    } else if (action === 'getApplications') {
      result = getApplications(params.adminKey || '', params.scheduleId || '');
    } else {
      result = { success: false, message: '不明なアクションです: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'サーバーエラー: ' + err.toString() };
  }

  return createResponse(result);
}

// ===== doPost: POSTリクエストハンドラ =====

function doPost(e) {
  var params;

  try {
    // JSON形式のリクエストボディをパース
    var body = e.postData && e.postData.contents ? e.postData.contents : '{}';
    params = JSON.parse(body);
  } catch (err) {
    // JSONパース失敗時はフォームデータとして試みる
    params = e.parameter || {};
  }

  var action = params.action || '';
  var result;

  try {
    if (action === 'apply') {
      result = apply(params);
    } else if (action === 'addSchedule') {
      result = addSchedule(params);
    } else if (action === 'updateSchedule') {
      result = updateSchedule(params);
    } else if (action === 'deleteSchedule') {
      result = deleteSchedule(params);
    } else {
      result = { success: false, message: '不明なアクションです: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'サーバーエラー: ' + err.toString() };
  }

  return createResponse(result);
}
