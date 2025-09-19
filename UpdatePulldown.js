// ===============================================================
// 機能1: プルダウン更新 (分岐機能)
// ===============================================================

/**
 * 【メニュー用】今月からプルダウンを更新します。
 */
function updateDropdownsForThisMonth() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName(APPLICATION_SHEET_NAME);
    if (!appSheet) throw new Error(`シート「${APPLICATION_SHEET_NAME}」が見つかりません。`);

    const config = getSettingsConfig();
    if (Object.keys(config).length === 0) {
      ui.alert('設定シートが空か、正しく読み込めませんでした。');
      return;
    }

    updateDropdownsForRow(appSheet, 3, config, false);

    const latestRow = appSheet.getLastRow();
    if (latestRow >= 4) {
       updateDropdownsForRow(appSheet, latestRow, config, true);
       ui.alert(`テンプレート行（3行目）と最新月（${latestRow}行目）の未入力プルダウンを更新しました。`);
    } else {
       ui.alert('テンプレート行（3行目）のプルダウンを更新しました。\n（データ行が存在しないため、最新月の更新はスキップされました）');
    }
  } catch (e) {
    console.error(`「今月から反映」処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
};

/**
 * 【メニュー用】次月からプルダウンを更新します。
 */
function updateDropdownsForNextMonth() {
  try {
    _internalUpdateTemplateRowOnly();
    SpreadsheetApp.getUi().alert('テンプレート行（3行目）のプルダウンを更新しました。');
  } catch (e) {
    console.error(`「次月から反映」処理中にエラーが発生しました: ${e.message}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
};