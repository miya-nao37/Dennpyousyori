/**
 * 特定行の指定範囲を、最終行の次にコピーします。
 */

function copyColumnToLast() {
  // --- 設定項目 ---
  const SOURCE_ROW = 3;       // コピー元の行番号を指定します (例: 2行目)
  const START_COLUMN = 1;     // コピー範囲の開始列を指定します (A列=1, B列=2, ...)
  const END_COLUMN = 9;       // コピー範囲の終了列を指定します (E列=5, ...)
  // --- 設定はここまで ---

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('KessaiSheet'); // 現在アクティブなシートを対象とします

  // 1. データが入力されている最終行を取得し、その次の行をコピー先として設定
  const lastRow = sheet.getLastRow();
  const destinationRow = lastRow + 1;

  // 2. コピーする列の数を計算
  const numColumns = END_COLUMN - START_COLUMN + 1;
  
  // 3. コピー元の範囲を取得
  const sourceRange = sheet.getRange(SOURCE_ROW, START_COLUMN, 1, numColumns);
  
  // 4. コピー先の範囲（開始セル）を取得
  const destinationRange = sheet.getRange(destinationRow, START_COLUMN);
  
  // 5. コピーを実行
  sourceRange.copyTo(destinationRange);

  // 6. 先頭列のセルに日付を入力
  const today = new Date();
  const month = today.getFullYear().toString() + '/0' + today.getMonth().toString();

  //下は結局使わない
  //const month = Utilities.formatDate(today, 'JST', 'yyyy/MM');

  sheet.getRange(destinationRow,1).setValue(month);
  sheet.getRange(destinationRow,1).setHorizontalAlignment("center");

  console.log(`${SOURCE_ROW}行目の${START_COLUMN}列目から${END_COLUMN}列目を、${destinationRow}行目にコピーしました。`);

  // 7. フラグの更新
  setFlag1('false');
  setFlag2('false');
  Logger.log("フラグをリセットしました。");
}