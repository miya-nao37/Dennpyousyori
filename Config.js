/**
 * 【REVISED】新しい10列構成の設定シートから情報を読み込み、オブジェクト形式で返します。
 * @returns {object} 項目名をキーとした設定情報のオブジェクト
 */
function getSettingsConfig() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return {};
  
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 10).getValues();
  const config = {};

  data.forEach(row => {
    const itemName = row[0]; // A列
    if (!itemName) return; // 決済項目名がなければスキップ

    const reminderDay = row[1]; // B列

    // C,D列から申請者名、E,F列からメールアドレスを収集（空欄は除外）
    const applicantNames = [row[2], row[3]].filter(String);
    const applicantEmails = [row[4], row[5]].filter(String);

    // G,H列から承認者名、I,J列からメールアドレスを収集（空欄は除外）
    const approverNames = [row[6], row[7]].filter(String);
    const approverEmails = [row[8], row[9]].filter(String);
    
    config[itemName] = {
      reminderDay: reminderDay,
      applicants: {
        names: applicantNames,
        emails: applicantEmails
      },
      approvers: {
        names: approverNames,
        emails: approverEmails
      }
    };
  });
  return config;
}