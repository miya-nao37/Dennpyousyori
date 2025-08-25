/**
 * 指定された行の偶数列がすべて記入されているか判定し、ログに出力します。
 * 記入されていればtrue、1つでも空欄があればfalseを返します。
 * @param {number} rowNumber - 判定したい行の番号（例：1）
 * @return {boolean} - 偶数列がすべて記入されていればtrue、そうでなければfalse
 */
function checkEvenColumns(rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dennpyousyori');
  
  // 指定された行のデータが入力されている最後の列番号を取得
  const lastColumn = sheet.getLastColumn();
  
  // 偶数列に空欄がないかチェックするためのフラグを初期化
  let currentFlag = getFlag1();
  
  // 2列目から始まり、2つずつ増やしながらループする（偶数列を順に処理）
  for (let col = 4; col <= lastColumn; col += 2) {
    // セルの値を取得
    const cellValue = sheet.getRange(rowNumber, col).getValue();
    
    // 値が空欄かどうかを確認（空文字列や数値の0も空欄と見なす場合は条件を適宜修正）
    if (cellValue === '') {
      setFlag1('false');
      break; // 1つでも空欄が見つかったらループを終了
    }
  }
  
  // 最終的な結果をログに出力
  Logger.log(currentFlag);
  
  return currentFlag;
}
