function checkAndSendMailToShinseisya() {
  // --- 設定項目 ---
  const TO_MAIL_ADDRESS = 'naoki_miyake@asahi.co.jp'; // リマインドメールの送信先
  // --- 設定はここまで ---

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Dennpyousyori'); // シート名から操作するシートを指定
    const lastRow = sheet.getLastRow(); // データが入力されている最終行を取得
    const range = sheet.getRange(lastRow,2,1,4); //チェックする範囲を取得

    const today = new Date();
    const dayOfMonth = today.getDate();

    // m×n行列形式の配列をflatで1次元配列にして、セルすべてがtrueかどうかを確認
    const isAllChecked = range.getValues().flat().every(cell => cell === true); 
    const currentFlag = getFlag1();

    if (isAllChecked) {
      if (currentFlag !== 'true') {
        const subject = '【完了報告】決済処理が完了しています';
        const body = `お疲れ様です。\n\nスプレッドシートの決済項目がすべてチェックされ、処理が完了したことを確認しました。\n\nご対応ありがとうございました。`;
        MailApp.sendEmail(TO_MAIL_ADDRESS, subject, body);
        console.log('すべてのチェックボックスがTRUEのため、完了メールを送信しました。');
        setFlag1('true');
      } else {
        console.log('すべてのチェックボックスはTRUEですが、今月は既に完了メールを送信済みのため、何もしません。');
      }
    } else {
      if (dayOfMonth >= 25) {
        const subject = '【要対応】決済処理に未完了の項目があります';
        const body = `お疲れ様です。\n\n決済処理に未完了の項目があります。\nスプレッドシートをご確認の上、チェックをお願いいたします。\n\nシートがすべてチェックされるまで、このリマインダーは毎日送信されます。\n\n▼スプレッドシートURL\n${spreadsheet.getUrl()}`;
        MailApp.sendEmail(TO_MAIL_ADDRESS, subject, body);
        console.log(`未チェックの項目があり、かつ25日以降のため、リマインドメールを送信しました。`);
      } else {
        console.log(`未チェックの項目がありますが、まだ25日ではないため、メールは送信しません。(本日: ${dayOfMonth}日)`);
      }
    }

  } catch (e) {
    // エラーが発生した場合、catchブロックが実行される

    // エラー内容をログに出力
    Logger.log("エラーが発生しました: " + e.message);

    // エラー通知メールを送信
    MailApp.sendEmail(
      TO_MAIL_ADDRESS, // 通知を受け取りたいメールアドレス
      'GAS スクリプト実行エラー', // メール件名
      'スクリプト実行中にエラーが発生しました。\n\n' +
      '関数名: checkAndSendMailToShinseisya\n' +
      'エラーメッセージ: ' + e.message + '\n' +
      'スタックトレース: ' + e.stack
    );
  }
}
