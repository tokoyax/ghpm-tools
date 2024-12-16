//====================
// 同じステータスにn日以上留まっているIssueをリストアップする
//====================
function listStaleIssues() {
  const settings = getSettingsFromSheet();

  const INPUT_SHEET_NAME = settings['日次履歴データ取得設定']['日次履歴データ出力シート名'] || 'DailyStatus';
  const OUTPUT_SHEET_NAME = settings['WIPチェック設定']['出力先シート名'] || 'StaleIssues';
  const WIP_MAX_DAYS = Number(settings['WIPチェック設定']['WIP最大日数']) || 5; // WIP最大日数
  const CHECK_STATUSES = (settings['WIPチェック設定']['チェック対象ステータス'] || '')
    .split(',')
    .map(status => status.trim())
    .filter(Boolean); // チェック対象ステータス
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INPUT_SHEET_NAME);
  if (!sheet) {
    Logger.log(`シート "${INPUT_SHEET_NAME}" が見つかりません。`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // ヘッダー行
  const results = [['Issue Number', 'Title', 'URL', 'Status', 'WIP days', 'Created At']]; // 出力ヘッダー

  // ヘッダーインデックスを取得
  const issueNumberIndex = headers.indexOf("Issue Number");
  const titleIndex = headers.indexOf("Title");
  const urlIndex = headers.indexOf("Issue URL");
  const statusIndex = headers.indexOf("Status");
  const createdAtIndex = headers.indexOf("Created At");
  const importDateIndex = headers.indexOf("Import DateTime");

  if (
    issueNumberIndex === -1 || titleIndex === -1 || urlIndex === -1 ||
    statusIndex === -1 || createdAtIndex === -1 || importDateIndex === -1
  ) {
    Logger.log("必要な列が見つかりません。設定を確認してください。");
    return;
  }

  const now = new Date();

  // データ行をチェック
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusIndex];
    const importDate = new Date(row[importDateIndex]);

    if (CHECK_STATUSES.includes(status)) {
      const wipDays = Math.floor((now - importDate) / (1000 * 60 * 60 * 24)); // 経過日数を計算
      if (wipDays >= WIP_MAX_DAYS) {
        results.push([
          row[issueNumberIndex],
          row[titleIndex],
          row[urlIndex],
          status,
          wipDays,
          row[createdAtIndex]
        ]);
      }
    }
  }

  // 結果を出力
  let outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTPUT_SHEET_NAME);
  if (!outputSheet) {
    outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(OUTPUT_SHEET_NAME);
  } else {
    outputSheet.clear(); // 既存データをクリア
  }

  if (results.length > 1) {
    outputSheet.getRange(1, 1, results.length, results[0].length).setValues(results);
    Logger.log(`WIP最大日数を超えたIssueをリストアップしました (${results.length - 1} 件)。`);
  } else {
    Logger.log("WIP最大日数を超えたIssueは見つかりませんでした。");
  }
}

//====================
// 設定値を取得する関数
//====================
function getSettingsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) {
    throw new Error('設定シートが存在しません。先に設定シートを作成してください。');
  }

  const data = sheet.getDataRange().getValues();
  const settings = {};

  data.slice(1).forEach(row => {
    const [category, key, value] = row;
    if (!settings[category]) {
      settings[category] = {};
    }
    settings[category][key] = value || '';
  });

  return settings;
}
