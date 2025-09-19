// ===============================================================
// 新しい決済項目をシートに追加する (罫線機能付き)
// ===============================================================
/**
 * 新しい決済項目行を設定シートに追加し、申請シートに列を追加します。
 * 追加された範囲に罫線を自動で設定します。
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

    // Action 1: 設定シートに行を追加し、罫線を引く
    const columnRange = settingsSheet.getRange("A" + "1");
    const dataRange = columnRange.getDataRegion(SpreadsheetApp.Dimension.ROWS);
    const aColumnLastRow = dataRange.getLastRow();

    //settingsSheet.appendRow([newItemName, '', '', '', '', '', '', '', '', '']);
    //const newSettingsRow = settingsSheet.getLastRow();
    settingsSheet.getRange(aColumnLastRow + 1, 1).setValue(newItemName);
    settingsSheet.getRange(aColumnLastRow + 1, 1, 1, 10).setBorder(true, true, true, true, true, true);
    console.log(`設定シートに「${newItemName}」の行を追加し、罫線を設定しました。`);

    // Action 2: 申請シートに列を追加し、ヘッダーと書式を設定
    const lastCol = appSheet.getLastColumn();
    appSheet.insertColumnsAfter(lastCol, 2);

    const header1Range = appSheet.getRange(1, lastCol + 1, 1, 2);
    header1Range.merge();
    header1Range.setValue(newItemName);
    header1Range.setHorizontalAlignment('center');

    const header2Range = appSheet.getRange(2, lastCol + 1, 1, 2);
    header2Range.setValues([['申請者', '承認者']]);

    // 書式をコピー
    appSheet.getRange(1, lastCol - 1, appSheet.getLastRow(), 2).copyTo(appSheet.getRange(1, lastCol + 1, appSheet.getLastRow(), 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    
    // 罫線を設定
    const newAppSheetRange = appSheet.getRange(1, lastCol + 1, appSheet.getLastRow(), 2);
    newAppSheetRange.setBorder(true, true, true, true, true, true);
    console.log(`申請シートに「${newItemName}」の列を追加し、罫線を設定しました。`);

    // Action 3: プルダウンを更新
    _internalUpdateTemplateRowOnly();

    ui.alert(`決済項目「${newItemName}」を追加しました。\n\n【重要】\n最後に「設定」シートを開き、追加された行の各項目（リマインド開始日、申請者、承認者、メールアドレス）を忘れずに入力してください。`);

  } catch (e) {
    console.error(`項目追加処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}