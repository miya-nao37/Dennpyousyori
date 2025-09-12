// ===============================================================
// 新しい決済項目をシートに追加する (氏名/メール分離版)
// ===============================================================
/**
 * ユーザーとの対話を通じて新しい決済項目情報を取得し、設定シートと申請シートの両方に自動で反映します。
 */
function addNewItemColumn() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const itemNameResponse = ui.prompt('ステップ1/3: 新しい決済項目の追加', '追加する決済項目名を入力してください:', ui.ButtonSet.OK_CANCEL);
    if (itemNameResponse.getSelectedButton() !== ui.Button.OK || !itemNameResponse.getResponseText()) return ui.alert('処理をキャンセルしました。');
    const newItemName = itemNameResponse.getResponseText().trim();

    const applicantsResponse = ui.prompt('ステップ2/3: 申請者リストの登録', `「${newItemName}」の申請者の氏名をカンマ区切りで入力してください。\n例: 田中 太郎, 鈴木 一郎`, ui.ButtonSet.OK_CANCEL);
    if (applicantsResponse.getSelectedButton() !== ui.Button.OK) return ui.alert('処理をキャンセルしました。');
    const applicantsList = applicantsResponse.getResponseText().trim();

    const approversResponse = ui.prompt('ステップ3/3: 承認者リストの登録', `「${newItemName}」の承認者の氏名をカンマ区切りで入力してください。\n例: 佐藤 花子`, ui.ButtonSet.OK_CANCEL);
    if (approversResponse.getSelectedButton() !== ui.Button.OK) return ui.alert('処理をキャンセルしました。');
    const approversList = approversResponse.getResponseText().trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet || !settingsSheet) throw new Error('必要なシートが見つかりません。');

    settingsSheet.appendRow([newItemName, '', applicantsList, '', approversList, '']);
    console.log(`設定シートに「${newItemName}」の行を追加しました。`);

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

    _internalUpdateTemplateRowOnly(); // テンプレート行を更新

    ui.alert(`決済項目「${newItemName}」を追加しました。\n\n【重要】\n最後に「設定」シートを開き、追加された行の「リマインド開始日」と、対応する「メールアドレス」の列を忘れずに入力してください。`);

  } catch (e) {
    console.error(`項目追加処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
};