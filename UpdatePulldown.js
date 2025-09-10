// ===============================================================
// テンプレート行のプルダウンを更新する
// ===============================================================
/**
 * 設定シートの最新情報に基づき、申請シートの3行目（テンプレート行）の
 * プルダウン（データ入力規則）を更新します。
 */
function updateTemplateRowDropdowns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet) throw new Error(`シート「${APPLICATION_SHEET_NAME}」が見つかりません。`);

    const config = getSettingsConfig();
    if (Object.keys(config).length === 0) {
      SpreadsheetApp.getUi().alert('設定シートが空か、正しく読み込めませんでした。');
      return;
    }

    const templateRow = 3; // 参照元は3行目
    const headers = appSheet.getRange(1, 1, 1, appSheet.getLastColumn()).getValues()[0];
    
    console.log('テンプレート行のプルダウン更新を開始します。');
    for (let col = 3; col < headers.length; col += 2) {
      const itemName = headers[col];
      const itemConfig = config[itemName];

      if (itemConfig) {
        // 申請者プルダウン
        if (itemConfig.applicants && itemConfig.applicants.length > 0) {
          const applicantCell = appSheet.getRange(templateRow, col + 1);
          const applicantRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.applicants).build();
          applicantCell.setDataValidation(applicantRule);
        }
        // 承認者プルダウン
        if (itemConfig.approvers && itemConfig.approvers.length > 0) {
          const approverCell = appSheet.getRange(templateRow, col + 2);
          const approverRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.approvers).build();
          approverCell.setDataValidation(approverRule);
        }
      } else {
        console.warn(`設定シートに項目「${itemName}」の設定が見つかりませんでした。`);
      }
    }
    SpreadsheetApp.getUi().alert('テンプレート行（3行目）のプルダウンを更新しました。');
  } catch (e) {
    console.error(`テンプレート行の更新中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}