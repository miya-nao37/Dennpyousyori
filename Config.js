/**
 * カンマ区切りの文字列をトリムして配列に変換します。
 * @param {string} str - カンマ区切りの文字列
 * @returns {string[]}
 */
function splitAndTrim(str) {
    if (!str || typeof str !== 'string' || str.trim() === '') return [];
    return str.split(',').map(item => item.trim());
}

/**
 * 設定シートから情報を読み込み、使いやすいオブジェクト形式で返します。
 * @returns {object} 項目名をキーとした設定情報のオブジェクト
 */
function getSettingsConfig() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return {};
  
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 6).getValues();
  const config = {};

  data.forEach(row => {
    const itemName = row[0];        // A列: 決済項目名
    const reminderDay = row[1];       // B列: リマインド開始日
    const applicantNames = row[2];  // C列: 申請者リスト（氏名）
    const applicantEmails = row[3]; // D列: 申請者メールアドレス
    const approverNames = row[4];   // E列: 承認者リスト（氏名）
    const approverEmails = row[5];  // F列: 承認者メールアドレス
    
    if (itemName) {
      config[itemName] = {
        reminderDay: reminderDay,
        applicants: {
          names: splitAndTrim(applicantNames),
          emails: splitAndTrim(applicantEmails)
        },
        approvers: {
          names: splitAndTrim(approverNames),
          emails: splitAndTrim(approverEmails)
        }
      };
    }
  });
  return config;
};