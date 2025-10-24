/**
 * 複数列でシートを並べ替える関数
 */
function multiColumnSort() {
  
  // ▼ 設定 ▼
  // -------------------------------------------
  // 対象のシート名 (空欄 "" の場合は現在開いているシート)
  var sheetName = ""; 
  
  // ソートの優先順位 (複数指定可)
  var sortCriteria = [
    { column: 4, ascending: false },  // 優先度1: 4列目 (D列: 完了/不要) を 降順 (false)
    { column: 6, ascending: true },  // 優先度2: 6列目 (F列: 目標日) を 昇順 (true)
    { column: 3, ascending: true },  // 優先度3: 3列目 (C列: 案件名) を 昇順 (true)
    { column: 5, ascending: true }  // 優先度4: 5列目 (E列: 期限) を 昇順 (true)
  ];
  
  // ヘッダー (見出し行) の行数
  var headerRows = 5;
  // -------------------------------------------
  // ▲ 設定 ▲

  
  try {
    // シートの取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet;
    if (sheetName === "") {
      sheet = spreadsheet.getActiveSheet();
    } else {
      sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        SpreadsheetApp.getUi().alert('エラー: シート "' + sheetName + '" が見つかりません。');
        return;
      }
    }

    // データ範囲の取得 (ヘッダーを除く)
    var startRow = headerRows + 1;
    if (sheet.getLastRow() < startRow) {
      // データがない場合は何もしない
      return; 
    }
    var range = sheet.getRange(startRow, 1, sheet.getLastRow() - headerRows, sheet.getLastColumn());

    // ソートの実行
    range.sort(sortCriteria);
    
    //SpreadsheetApp.getUi().alert('ソートが完了しました。'); // 完了メッセージ (任意)

  } catch (e) {
    SpreadsheetApp.getUi().alert('ソート中にエラーが発生しました: ' + e.message);
  }
}
