// ===============================================================
// 機能2: 新しい決済項目をシートに追加する
// ===============================================================
/**
 * 新しい決済項目行を設定シートに追加し、申請シートに列を追加します。
 */
function addNewItemColumn() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const itemNameResponse = ui.prompt('新しい決済項目の追加', '追加する決済項目名を入力してください:', ui.ButtonSet.OK_CANCEL);
    if (itemNameResponse.getSelectedButton() !== ui.Button.OK || !itemNameResponse.getResponseText()) return ui.alert('処理をキャンセルしました。');
    const newItemName = itemNameResponse.getResponseText().trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet || !settingsSheet) throw new Error('必要なシートが見つかりません。');

    // 新しい項目用の空行を設定シートに追加
    settingsSheet.appendRow([newItemName, '', '', '', '', '', '', '', '', '']);
    console.log(`設定シートに「${newItemName}」の行を追加しました。`);

    // 追加した行に格子状の罫線を設定
    const newSettingsRow = settingsSheet.getLastRow();
    const settingsLastCol = settingsSheet.getLastColumn();
    settingsSheet.getRange(newSettingsRow, 1, 1, settingsLastCol).setBorder(true, true, true, true, true, true);

    const lastCol = appSheet.getLastColumn();
    appSheet.insertColumnsAfter(lastCol, 2);

    const header1Range = appSheet.getRange(1, lastCol + 1, 1, 2);
    header1Range.merge();
    header1Range.setValue(newItemName);
    header1Range.setHorizontalAlignment('center');

    const header2Range = appSheet.getRange(2, lastCol + 1, 1, 2);
    header2Range.setValues([['申請者', '承認者']]);

    appSheet.getRange(1, lastCol - 1, appSheet.getMaxRows(), 2).copyTo(appSheet.getRange(1, lastCol + 1, appSheet.getMaxRows(), 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    console.log(`申請シートに「${newItemName}」の列を追加しました。`);

    // 追加した列に格子状の罫線を設定
    const newColumnsRange = appSheet.getRange(1, lastCol + 1, appSheet.getMaxRows(), 2);
    newColumnsRange.setBorder(true, true, true, true, true, true);

    _internalUpdateTemplateRowOnly();

    ui.alert(`決済項目「${newItemName}」を追加しました。\n\n【重要】\n最後に「設定」シートを開き、追加された行の各項目（リマインド開始日、申請者、承認者、メールアドレス）を忘れずに入力してください。`);

  } catch (e) {
    console.error(`項目追加処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}