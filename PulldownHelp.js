// ===============================================================
// ヘルパー機能: 補助的な役割を持つ関数群
// ===============================================================

/**
 * 【内部用】テンプレート行のプルダウンのみを更新します。
 */
function _internalUpdateTemplateRowOnly() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet) throw new Error(`シート「${APPLICATION_SHEET_NAME}」が見つかりません。`);

    const config = getSettingsConfig();
    if (Object.keys(config).length === 0) {
      console.warn('設定シートが空か、正しく読み込めませんでした。更新をスキップします。');
      return;
    }
    
    updateDropdownsForRow(appSheet, 3, config, false);
    console.log('テンプレート行のプルダウンを内部的に更新しました。');
}


/**
 * 指定された行のプルダウンを更新します。
 */
function updateDropdownsForRow(sheet, rowNum, config, skipNonEmpty) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (let col = 3; col < headers.length; col += 2) {
    const itemName = headers[col];
    const itemConfig = config[itemName];

    if (itemConfig) {
      const applicantCell = sheet.getRange(rowNum, col + 1);
      if (!(skipNonEmpty && applicantCell.getValue() !== '')) {
        if (itemConfig.applicants && itemConfig.applicants.names.length > 0) {
          const applicantRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.applicants.names).build();
          applicantCell.setDataValidation(applicantRule);
        }
      }

      const approverCell = sheet.getRange(rowNum, col + 2);
      if (!(skipNonEmpty && approverCell.getValue() !== '')) {
        if (itemConfig.approvers && itemConfig.approvers.names.length > 0) {
          const approverRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.approvers.names).build();
          approverCell.setDataValidation(approverRule);
        }
      }
    }
  }
}