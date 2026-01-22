// デバッグ用でeventオブジェクトの作成
function debug() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  const Row = 3;
  const range = sheet.getRange(Row, CONFIG.COLUMNS.START); // 開始日時のセルを仮の編集対象とする

  // 本来トリガーから渡されるeventを疑似的に作成
  const debugEvent = {
    range: range,
    source: SpreadsheetApp.getActiveSpreadsheet()
  };

  console.log("テスト実行開始…");
  syncToCalendar(debugEvent);
  console.log("テストとカレンダーを確認")
}