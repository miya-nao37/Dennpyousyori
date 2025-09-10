// ===============================================================
// 定数設定エリア
// ===============================================================
const APPLICATION_SHEET_NAME = '申請';
const SETTINGS_SHEET_NAME = '設定';


// ===============================================================
// メニュー追加機能
// ===============================================================
/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 決済管理Menu')
    .addItem('テンプレート行のプルダウンを更新', 'updateTemplateRowDropdowns')
    .addItem('決済項目を追加', 'addNewItemColumn')
    .addSeparator()
    .addItem('リマインドメールを送信（手動）', 'checkApprovalsAndSendReminders')
    .addToUi();
}


// ===============================================================
// リマインドメールを担当者ごとに送信する (項目別開始日に対応)
// ===============================================================
/**
 * 申請・承認がされていない項目をチェックし、担当者ごとにリマインドメールを送信します。
 * 項目ごとに設定されたリマインド開始日を考慮します。
 */
function checkApprovalsAndSendReminders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet) throw new Error('申請シートが見つかりません。');
    
    const config = getSettingsConfig();
    if (Object.keys(config).length === 0) {
        console.warn('設定シートが空か、正しく読み込めませんでした。リマインド処理を中断します。');
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
      
      if (!itemConfig || !itemConfig.email || today.getDate() < itemConfig.reminderDay) {
        continue;
      }

      const applicant = targetRowValues[col];
      const approver = targetRowValues[col + 1];
      const email = itemConfig.email;

      if (!reminders[email]) {
        reminders[email] = [];
      }
      if (!applicant) {
        reminders[email].push(`  - ${itemName} (申請者)`);
      }
      if (!approver) {
        reminders[email].push(`  - ${itemName} (承認者)`);
      }
    }
    
    for (const email in reminders) {
      const missingItems = reminders[email];
      if (missingItems.length > 0) {
        const subject = `【要対応】${targetYearMonth}度 決済処理の申請・承認のお願い`;
        const body = `
お疲れ様です。

ご担当の${targetYearMonth}度分 決済処理について、以下の項目で申請または承認が完了していません。
内容をご確認の上、ご対応をお願いいたします。

▼ 未完了の項目
${missingItems.join('\n')}

▼ 詳細は下記のスプレッドシートをご確認ください。
${ss.getUrl()}

※このメールはシステムにより自動送信されています。
        `;
        MailApp.sendEmail(email, subject, body.trim());
        console.log(`リマインドメールを ${email} に送信しました。未完了項目: ${missingItems.length}件`);
      } else {
        console.log(`${email} 宛のタスクはすべて完了しています。メールは送信しませんでした。`);
      }
    }

  } catch (e) {
    console.error(`リマインダー処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
}

// ===============================================================
// ヘルパー機能: 設定シートの情報をオブジェクトとして取得
// ===============================================================
/**
 * 設定シートから情報を読み込み、使いやすいオブジェクト形式で返します。
 * @returns {object} 項目名をキーとした設定情報のオブジェクト
 */
function getSettingsConfig() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) return {};
  
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 6).getValues();
  const config = {};

  data.forEach(row => {
    const itemName = row[0];
    const email = row[2];
    const reminderDay = row[3]
    const applicants = row[4] ? row[4].toString().split(',').map(item => item.trim()) : [];
    const approvers = row[5] ? row[5].toString().split(',').map(item => item.trim()) : [];
    
    if (itemName) {
      config[itemName] = {
        email: email,
        reminderDay: reminderDay,
        applicants: applicants,
        approvers: approvers
      };
    }
  });
  return config;
}


