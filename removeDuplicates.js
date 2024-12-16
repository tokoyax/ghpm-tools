//
// DailyStatusから重複行を削除するやつ
//
function removeDuplicatesKeepOldest() {
  // スプレッドシートのIDとシート名を指定
  const sheetName = "DailyStatus_backup_rm_dup";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet ${sheetName} not found`);
    return;
  }
  
  // データを取得
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // ヘッダーを除外
  const uniqueData = [];
  const seenRecords = new Map();

  // "Import DateTime" の列番号を特定
  const importDateTimeIndex = headers.indexOf("Import DateTime");

  if (importDateTimeIndex === -1) {
    Logger.log("Column 'Import DateTime' not found");
    return;
  }

  // 各行をチェックし、重複を排除
  data.forEach(row => {
    const key = row
      .filter((_, index) => index !== importDateTimeIndex) // "Import DateTime" 列を除く
      .join("|"); // ユニークなキーとして文字列化
    const currentDate = new Date(row[importDateTimeIndex]); // "Import DateTime" を日時として扱う

    if (!seenRecords.has(key)) {
      seenRecords.set(key, row); // 初回出現の行を保存
    } else {
      const existingRow = seenRecords.get(key);
      const existingDate = new Date(existingRow[importDateTimeIndex]);
      if (currentDate < existingDate) {
        seenRecords.set(key, row); // より古い行で更新
      }
    }
  });

  // Map からユニークなデータを抽出
  seenRecords.forEach(value => uniqueData.push(value));

  // シートをクリアして新しいデータを書き込み
  sheet.clearContents();
  sheet.appendRow(headers);
  sheet.getRange(2, 1, uniqueData.length, headers.length).setValues(uniqueData);

  Logger.log("Duplicate rows removed successfully, oldest rows retained");
}
