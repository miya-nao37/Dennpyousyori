// ===============================================================
// 定数設定エリア
// ===============================================================
const APPLICATION_SHEET_NAME = '申請';
const SETTINGS_SHEET_NAME = '設定';


// ===============================================================
// メニュー追加機能 (分岐版)
// ===============================================================
/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 * 「プルダウンを更新」にサブメニューを持たせます。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 決済管理メニュー')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('プルダウンを更新')
      .addItem('今月から反映', 'updateDropdownsForThisMonth')
      .addItem('次月から反映', 'updateDropdownsForNextMonth'))
    .addItem('決済項目を追加', 'addNewItemColumn')
    .addSeparator()
    .addItem('リマインドメールを送信（手動）', 'checkApprovalsAndSendReminders')
    .addToUi();
}


// ===============================================================
// トリガー設定用メイン関数
// ===============================================================
/**
 * 毎日定時に実行されるトリガー用の関数です。
 */
function dailyTrigger() {
  console.log('日次トリガーを開始します。');
  checkApprovalsAndSendReminders();
  console.log('日次トリガーが正常に終了しました。');
}


// ===============================================================
// リマインドメールを関係者全員に送信する
// ===============================================================
/**
 * 未入力の項目がある場合、その項目に関わる関係者（申請者・承認者リストの全員）にリマインドメールを送信します。
 */
function checkApprovalsAndSendReminders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet) throw new Error('申請シートが見つかりません。');
    
    const config = getSettingsConfig();
    if (Object.keys(config).length === 0) {
        console.warn('設定シートが空か、正しく読み込めませんでした。');
        return;
    }

    const today = new Date();
    const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const targetYearMonth = Utilities.formatDate(lastMonth, 'JST', 'yyyy/MM');
    const dateValues = appSheet.getRange('A4:A' + appSheet.getLastRow()).getValues();
    const targetRow = dateValues.findIndex(row => row[0] instanceof Date && Utilities.formatDate(row[0], 'JST', 'yyyy/MM') === targetYearMonth) + 4;

    if (targetRow < 4) {
      console.log(`チェック対象の年月（${targetYearMonth}）の行が見つかりませんでした。`);
      return;
    }

    const headers = appSheet.getRange(1, 1, 1, appSheet.getLastColumn()).getValues()[0];
    const targetRowValues = appSheet.getRange(targetRow, 1, 1, appSheet.getLastColumn()).getValues()[0];
    const reminders = {}; 

    for (let col = 3; col < headers.length; col += 2) {
      const itemName = headers[col];
      const itemConfig = config[itemName];
      
      if (!itemConfig || today.getDate() < itemConfig.reminderDay) continue;

      const applicantValue = targetRowValues[col];
      const approverValue = targetRowValues[col + 1];
      
      if (!applicantValue && itemConfig.applicants.emails.length > 0) {
        const missingItemString = `  - ${itemName} (申請者)`;
        const emailsToNotify = itemConfig.applicants.emails;
        emailsToNotify.forEach(email => {
          if (!reminders[email]) reminders[email] = [];
          if (!reminders[email].includes(missingItemString)) reminders[email].push(missingItemString);
        });
      }
      
      if (!approverValue && itemConfig.approvers.emails.length > 0) {
        const missingItemString = `  - ${itemName} (承認者)`;
        const emailsToNotify = itemConfig.approvers.emails;
        emailsToNotify.forEach(email => {
          if (!reminders[email]) reminders[email] = [];
          if (!reminders[email].includes(missingItemString)) reminders[email].push(missingItemString);
        });
      }
    }
    
    for (const email in reminders) {
      const missingItems = reminders[email];
      if (missingItems.length > 0) {
        const subject = `【要対応】${targetYearMonth}度 決済処理の申請・承認のお願い`;
        const body = `
各位

${targetYearMonth}度分 決済処理について、ご担当の以下項目で未完了のタスクがあります。
内容をご確認の上、ご対応をお願いいたします。

▼ 未完了の項目
${missingItems.join('\n')}

▼ 詳細は下記のスプレッドシートをご確認ください。
${ss.getUrl()}

※このメールはシステムにより関係者全員に自動送信されています。
        `;
        MailApp.sendEmail(email, subject, body.trim());
        console.log(`リマインドメールを ${email} に送信しました。未完了項目: ${missingItems.length}件`);
      }
    }

  } catch (e) {
    console.error(`リマインダー処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
};