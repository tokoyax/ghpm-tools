//====================
// WIPチェック: StaleIssuesの生成
//====================

// メイン実行関数
function listStaleIssues() {
  const settings = loadSettings(); // Settingsシートから設定をロード
  const dailyStatusData = getDailyStatusData(settings.dailyStatusSheetName);
  const staleIssues = identifyStaleIssues(
    dailyStatusData,
    settings.checkStatuses,
    settings.maxWipDays
  );
  writeStaleIssuesToSheet(staleIssues, settings.staleIssuesSheetName);
}

//====================
// 設定の読み込み
//====================
function loadSettings() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!settingsSheet) throw new Error('Settingsシートが存在しません。');

  const data = settingsSheet.getDataRange().getValues();
  const settings = {};

  data.slice(1).forEach(([category, key, value]) => {
    if (!settings[category]) settings[category] = {};
    settings[category][key] = value;
  });

  return {
    dailyStatusSheetName: settings['日次履歴データ取得設定']['日次履歴データ出力シート名'],
    staleIssuesSheetName: settings['WIPチェック設定']['出力先シート名'],
    maxWipDays: parseInt(settings['WIPチェック設定']['WIP最大日数'], 10),
    checkStatuses: settings['WIPチェック設定']['チェック対象ステータス'].split(',').map(status => status.trim()),
  };
}

//====================
// DailyStatusデータの取得
//====================
function getDailyStatusData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`${sheetName} シートが存在しません。`);

  const data = sheet.getDataRange().getValues();
  return data.slice(1); // ヘッダーを除外
}

//====================
// Issueの滞留日数を計算
//====================
function identifyStaleIssues(dailyStatusData, checkStatuses, maxWipDays) {
  const issueMap = groupBy(dailyStatusData, row => row[1]); // Issue Number (Index 1)
  const today = new Date();

  return Object.values(issueMap).flatMap(issueRows => {
    const statusChanges = issueRows.map(row => ({
      date: new Date(row[0]), // Import DateTime (Index 0)
      status: row[3],         // Status (Index 3)
      createdAt: row[5],      // Created At (Index 5)
      url: row[10],           // Issue URL (Index 10)
      title: row[2],          // Title (Index 2)
    }));

    const latestStatus = statusChanges[statusChanges.length - 1];
    const filteredStatus = checkStatuses.includes(latestStatus.status);

    if (!filteredStatus) return []; // チェック対象ステータスでない場合は除外

    const earliestDate = getEarliestDateForStatus(statusChanges, latestStatus.status);
    const wipDays = calculateWipDays(today, earliestDate);

    if (wipDays > maxWipDays) {
      return [{
        issueNumber: issueRows[0][1],
        title: latestStatus.title,
        url: latestStatus.url,
        status: latestStatus.status,
        wipDays: wipDays,
        createdAt: latestStatus.createdAt,
      }];
    }
    return [];
  });
}

//====================
// 最も古いステータス変更日を取得
//====================
function getEarliestDateForStatus(statusChanges, targetStatus) {
  return statusChanges
    .filter(change => change.status === targetStatus)
    .reduce((earliest, current) => (current.date < earliest ? current.date : earliest), new Date());
}

//====================
// WIP日数の計算
//====================
function calculateWipDays(today, startDate) {
  const msInDay = 1000 * 60 * 60 * 24;
  return Math.floor((today - startDate) / msInDay);
}

//====================
// 結果をStaleIssuesシートに出力
//====================
function writeStaleIssuesToSheet(issues, sheetName) {
  const sheet = getOrCreateSheet(sheetName);

  // ヘッダー行を書き込む
  sheet.clear();
  sheet.appendRow(['Issue Number', 'Title', 'URL', 'Status', 'WIP days', 'Created At']);

  // データを書き込む
  issues.forEach(issue => {
    sheet.appendRow([
      issue.issueNumber,
      issue.title,
      issue.url,
      issue.status,
      issue.wipDays,
      issue.createdAt,
    ]);
  });
}

//====================
// シートが存在しない場合は作成
//====================
function getOrCreateSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
}

//====================
// 配列をキーごとにグループ化
//====================
function groupBy(array, keySelector) {
  return array.reduce((groups, item) => {
    const key = keySelector(item);
    if (!groups[key]) groups[key] = [];
    groups[key].push(item);
    return groups;
  }, {});
}
