// ===============================================================
// ヘルパー機能: 補助的な役割を持つ関数群
// ===============================================================

/**
 * 【NEW・内部用】テンプレート行のプルダウンのみを更新します。UIフィードバックなし。
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
    
    updateDropdownsForRow(appSheet, 3, config, false); // 3行目を全更新
    console.log('テンプレート行のプルダウンを内部的に更新しました。');
}


/**
 * 指定された行のプルダウン（データ入力規則）を更新します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシートオブジェクト
 * @param {number} rowNum - 更新する行番号
 * @param {object} config - 設定情報のオブジェクト
 * @param {boolean} skipNonEmpty - trueの場合、すでに入力済みのセルは更新をスキップする
 */
function updateDropdownsForRow(sheet, rowNum, config, skipNonEmpty) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (let col = 3; col < headers.length; col += 2) { // D列から2列ごとに処理
    const itemName = headers[col];
    const itemConfig = config[itemName];

    if (itemConfig) {
      // 申請者プルダウン
      const applicantCell = sheet.getRange(rowNum, col + 1);
      if (!(skipNonEmpty && applicantCell.getValue() !== '')) {
        if (itemConfig.applicants && itemConfig.applicants.names.length > 0) {
          const applicantRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.applicants.names).build();
          applicantCell.setDataValidation(applicantRule);
        }
      }

      // 承認者プルダウン
      const approverCell = sheet.getRange(rowNum, col + 2);
      if (!(skipNonEmpty && approverCell.getValue() !== '')) {
        if (itemConfig.approvers && itemConfig.approvers.names.length > 0) {
          const approverRule = SpreadsheetApp.newDataValidation().requireValueInList(itemConfig.approvers.names).build();
          approverCell.setDataValidation(approverRule);
        }
      }
    }
  }
};