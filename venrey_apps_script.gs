// ============================================================
// Venrey スケジュール 月次自動シート作成スクリプト
// ============================================================
//
// 【使い方】
//   1. Google スプレッドシートを開く
//   2. 拡張機能 → Apps Script
//   3. このファイルの内容を全て貼り付けて保存
//   4. 「setupMonthlyTrigger」を一度だけ手動実行（トリガー登録）
//   5. 今すぐ4月シートを作りたい場合は「createNewMonthSheetNow」を実行
//
// 【動作】
//   ・毎月1日0時に自動実行
//   ・前月シートをベースに新月シートを作成（スタッフ名・曜日をコピー）
//   ・前月シートに入力済みの「来月データ列」があれば自動で移行する
//   ・シート名は「4月」「5月」のような形式で作成される
//   ・main.py はシート名を自動検出するため GID 変更不要
//
// ============================================================

// 現在の3月シートで「来月（4月）データ」が始まる列番号（1始まり）
// 例: AI列 = 35
var NEXT_MONTH_DATA_START_COL = 35;

// 前月シートに来月データがある場合、移行後にその列を削除するか
// true: 削除する（シートをすっきりさせる）
// false: 削除しない（念のため残しておく）
var DELETE_NEXT_MONTH_COLS_FROM_SOURCE = false;

// ============================================================

/**
 * 初回セットアップ：毎月1日0時のトリガーを登録する
 * ★ 一度だけ手動で実行してください
 */
function setupMonthlyTrigger() {
  // 重複防止：既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'createNewMonthSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('createNewMonthSheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log('✅ 毎月1日0時の自動実行トリガーを登録しました');
  SpreadsheetApp.getUi().alert('✅ トリガー登録完了！\n\n毎月1日0時に自動でシートが作成されます。');
}

/**
 * 今すぐ手動で新月シートを作成したい場合に実行する
 * 初回（3月 → 4月）はこちらを実行してください
 */
function createNewMonthSheetNow() {
  createNewMonthSheet();
}

/**
 * 本体：新しい月のシートを作成する
 * 毎月1日0時に自動実行、または手動実行
 */
function createNewMonthSheet() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var now  = new Date();
  var year = now.getFullYear();
  var mon  = now.getMonth() + 1; // 1始まり（4 = 4月）

  var newSheetName  = mon + '月';
  var prevSheetName = (mon === 1 ? 12 : mon - 1) + '月';

  // 既に存在する場合はスキップ
  if (ss.getSheetByName(newSheetName)) {
    Logger.log(newSheetName + ' は既に存在します。スキップします。');
    SpreadsheetApp.getUi().alert(newSheetName + ' シートは既に存在します。');
    return;
  }

  // 前月シートを取得（なければアクティブシートを使用）
  var srcSheet = ss.getSheetByName(prevSheetName) || ss.getActiveSheet();
  Logger.log('コピー元シート: ' + srcSheet.getName());

  var srcLastRow = srcSheet.getLastRow();
  var srcLastCol = srcSheet.getLastColumn();
  var srcData    = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();

  // 今月の日数を計算
  var daysInMonth = new Date(year, mon, 0).getDate();

  // 新シートを作成（末尾に追加）
  var newSheet = ss.insertSheet(newSheetName);

  // ── 1行目：月名 + 日付番号 ──────────────────────────────
  // A1: "4月😊", B1: "", C1: ""（出勤日数列等）, D1以降: 1, 2, ..., 末日
  var headerRow = [];
  headerRow.push(mon + '月😊'); // A1
  headerRow.push('');            // B1
  headerRow.push('');            // C1
  for (var d = 1; d <= daysInMonth; d++) {
    headerRow.push(d);
  }
  newSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  // ── 2行目：曜日ヘッダー ────────────────────────────────
  // A列〜C列は前月シートの2行目をそのままコピー
  var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  var row2 = [];
  row2.push(srcData.length >= 2 ? srcData[1][0] : '名前');   // A2
  row2.push(srcData.length >= 2 ? srcData[1][1] : '確認日'); // B2
  row2.push(srcData.length >= 2 ? srcData[1][2] : '出勤日数'); // C2
  for (var d2 = 1; d2 <= daysInMonth; d2++) {
    var dow = new Date(year, mon - 1, d2).getDay();
    row2.push(dayNames[dow]);
  }
  newSheet.getRange(2, 1, 1, row2.length).setValues([row2]);

  // 土日に背景色をつける（見やすくする）
  for (var d3 = 1; d3 <= daysInMonth; d3++) {
    var dow2 = new Date(year, mon - 1, d3).getDay();
    var col  = 3 + d3; // D列 = 4番目 = col 4 (1始まり)
    if (dow2 === 0) { // 日曜
      newSheet.getRange(1, col, srcLastRow, 1).setBackground('#fce8e6');
    } else if (dow2 === 6) { // 土曜
      newSheet.getRange(1, col, srcLastRow, 1).setBackground('#e8f0fe');
    }
  }

  // ── 3行目以降：スタッフ名 + 来月分のシフトデータを移行 ──
  var nextMonthDataExists = (srcLastCol >= NEXT_MONTH_DATA_START_COL);
  var staffCount = 0;

  for (var r = 2; r < srcData.length; r++) {
    var nameVal = String(srcData[r][0]).trim();
    if (!nameVal || nameVal === '' || nameVal === 'null' || nameVal === 'undefined') continue;

    var rowNum = r + 1; // スプレッドシートの行番号（1始まり）

    // A列：スタッフ名（区切り行含む）
    newSheet.getRange(rowNum, 1).setValue(srcData[r][0]);
    // B列・C列は空にする（確認日・出勤日数はリセット）
    newSheet.getRange(rowNum, 2).setValue('');
    newSheet.getRange(rowNum, 3).setValue('');

    // 来月のシフトデータが前月シートに存在する場合、移行する
    if (nextMonthDataExists) {
      for (var d4 = 0; d4 < daysInMonth; d4++) {
        var srcColIdx  = NEXT_MONTH_DATA_START_COL - 1 + d4; // 0始まり
        var destCol    = 4 + d4; // D列 = 4番目（1始まり）
        if (srcColIdx < srcData[r].length) {
          var cellVal = srcData[r][srcColIdx];
          if (cellVal !== null && cellVal !== '' && String(cellVal).trim() !== '') {
            newSheet.getRange(rowNum, destCol).setValue(cellVal);
          }
        }
      }
    }

    staffCount++;
  }

  // ── 来月データ列を前月シートから削除（設定による）──
  if (DELETE_NEXT_MONTH_COLS_FROM_SOURCE && nextMonthDataExists) {
    var colsToDelete = srcLastCol - NEXT_MONTH_DATA_START_COL + 1;
    if (colsToDelete > 0) {
      srcSheet.deleteColumns(NEXT_MONTH_DATA_START_COL, colsToDelete);
      Logger.log('前月シートの来月列を削除しました（' + colsToDelete + '列）');
    }
  }

  // ── 列幅を調整 ──
  newSheet.setColumnWidth(1, 120); // A列：名前
  newSheet.setColumnWidth(2, 70);  // B列：確認日
  newSheet.setColumnWidth(3, 60);  // C列：出勤日数
  for (var c = 4; c <= 3 + daysInMonth; c++) {
    newSheet.setColumnWidth(c, 75); // 日付列
  }

  // ── 完了ログ ──
  var msg = '✅ ' + newSheetName + ' シートを作成しました！\n\n'
    + '・スタッフ数: ' + staffCount + '人\n'
    + '・日数: ' + daysInMonth + '日\n'
    + (nextMonthDataExists ? '・前月シートから来月データを移行しました\n' : '')
    + '\nVenrey自動更新は「' + newSheetName + '」シートを自動で読み込みます。';

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}
