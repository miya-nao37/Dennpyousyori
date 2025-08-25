    /** 
    if (dayOfMonth >= 20) {
      if (isAllChecked) {
        if(currentFlag !== 'true') {
          const subject = '【完了報告】決済処理がすべて完了しています';
          const body = `お疲れ様です。\n\nスプレッドシートの決済項目がすべてチェックされ、処理が完了したことを確認しました。\n\nご対応ありがとうございました。`;
          MailApp.sendEmail(TO_EMAIL_ADDRESS, subject, body);
          console.log('すべてのチェックボックスがTRUEのため、完了メールを送信しました。');
          setFlag('true');
        } else {
          console.log('すべてのチェックボックスはTRUEですが、今月は既に完了メールを送信済みのため、何もしません。');
        }
      } else {
        const subject = '【要対応】決済処理に未完了の項目があります';
        const body = `お疲れ様です。\n\n決済処理に未完了の項目があります。\nスプレッドシートをご確認の上、チェックをお願いいたします。\n\nシートがすべてチェックされるまで、このリマインダーは毎日送信されます。\n\n▼スプレッドシートURL\n${spreadsheet.getUrl()}`;
        MailApp.sendEmail(TO_EMAIL_ADDRESS, subject, body);
        console.log(`未チェックの項目があり、かつ20日以降のため、リマインドメールを送信しました。`);
      }
    } else {
      console.log(`まだ20日ではないため、メールは送信しません。(本日: ${dayOfMonth}日)`);
    }
    */

/**
 * function checkStringsAndOutput() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 検索したい文字列が格納されている範囲（例：A1:A3）
  const targetRange = sheet.getRange("A1:A3");
  const targetValues = targetRange.getValues();

  // 検索対象の表の範囲（例：B1:D5）
  const searchRange = sheet.getRange("B1:D5");
  const searchValues = searchRange.getValues();

  // 結果を出力するセル（例：E1）
  const outputCell = sheet.getRange("E1");

  // ターゲットの文字列を一次元配列に変換
  const flatTargetValues = targetValues.flat();

  let isFound = false;

  // 検索対象の表をループ
  for (let i = 0; i < searchValues.length; i++) {
    for (let j = 0; j < searchValues[i].length; j++) {
      const cellValue = searchValues[i][j];

      // セルが文字列であることを確認
      if (typeof cellValue === 'string') {
        // ターゲットの文字列をループして、セルの値に含まれているかチェック
        for (const targetText of flatTargetValues) {
          if (typeof targetText === 'string' && cellValue.includes(targetText)) {
            isFound = true;
            break; // 見つかったら内部ループを終了
          }
        }
      }
      if (isFound) {
        break; // 見つかったら外部ループも終了
      }
    }
    if (isFound) {
      break;
    }
  }

  // 結果をE1セルに書き込む
  outputCell.setValue(isFound);
}
 */