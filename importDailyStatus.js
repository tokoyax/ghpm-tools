//====================
// 初期設定ダイアログ表示
//====================
function showInitialSettingsDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('InitialSettings')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '初期設定');
}

//====================
// 設定シート作成
//====================
function createSettingsSheet() {
  const sheetName = 'Settings';
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (sheet) {
    SpreadsheetApp.getUi().alert('設定シートが既に存在します。');
    return;
  }

  sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);

  const settingsData = [
    ['カテゴリー', '項目', '値', '説明'],
    ['基本設定', 'リポジトリのオーナー', '', 'GitHubリポジトリのオーナー名'],
    ['基本設定', 'リポジトリ名', '', 'GitHubリポジトリ名'],
    ['日次履歴データ取得設定', '日次履歴データ出力シート名', 'DailyStatus', 'データの保存先シート名'],
    ['日次履歴データ取得設定', 'Githubから取得するIssueのラベル(複数指定可能)', '', '取得対象のラベル'],
    ['日次履歴データ取得設定', 'Githubから取得除外するIssueのラベル(複数指定可能)', '', '除外対象のラベル'],
    ['サイクルタイムデータ生成設定', 'サイクルタイムデータ出力シート名', 'CycleTime', 'サイクルタイムデータの保存先'],
    ['サイクルタイムデータ生成設定', '生成開始日', '', 'データ生成開始日 (YYYY/MM/DD)'],
    ['サイクルタイムデータ生成設定', '生成終了日', '', 'データ生成終了日 (YYYY/MM/DD)'],
    ['サイクルタイムデータ生成設定', 'ステータス From', '', '開始ステータス'],
    ['サイクルタイムデータ生成設定', 'ステータス To', '', '終了ステータス'],
    ['サイクルタイムデータ生成設定', '含めるラベル', '', '取得対象のラベル'],
    ['サイクルタイムデータ生成設定', '除外するラベル', '', '除外対象のラベル'],
    ['コントロールチャートデータ生成設定', 'コントロールチャートデータ出力シート名', 'ControlChart', 'チャートデータの保存先'],
    ['コントロールチャートデータ生成設定', 'EMAのスムージング係数', '0.2', 'スムージング係数 (0.1～0.5を推奨)']
  ];

  sheet.getRange(1, 1, settingsData.length, settingsData[0].length).setValues(settingsData);

  // ヘッダーのフォーマット
  sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#f4f4f4');
  sheet.setColumnWidths(1, 4, 200);

  SpreadsheetApp.getUi().alert('設定シートを作成しました。');
}

//====================
// 設定値の取得
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

//====================
// GitHubトークン取得
//====================
function getGitHubToken() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('GITHUB_TOKEN') || ''; // 空の場合に空文字列を返す
}

//====================
// GitHubトークン保存
//====================
function saveGitHubToken(token) {
  const userProperties = PropertiesService.getUserProperties();
  if (token) {
    userProperties.setProperty('GITHUB_TOKEN', token);
    SpreadsheetApp.getUi().alert('GitHubトークンが保存されました。');
  } else {
    userProperties.deleteProperty('GITHUB_TOKEN'); // 空のトークンの場合は削除
    SpreadsheetApp.getUi().alert('GitHubトークンが削除されました。');
  }
}

//====================
// メインの実行関数
//====================
function fetchDailyIssueStatus() {
  const settings = getSettingsFromSheet();

  const token = getGitHubToken();
  const owner = settings['基本設定']['リポジトリのオーナー'];
  const repo = settings['基本設定']['リポジトリ名'];
  const includeLabels = (settings['日次履歴データ取得設定']['Githubから取得するIssueのラベル(複数指定可能)'] || '').split(',').filter(Boolean);
  const excludeLabels = (settings['日次履歴データ取得設定']['Githubから取得除外するIssueのラベル(複数指定可能)'] || '').split(',').filter(Boolean);

  let allIssues = [];
  let hasNextPage = true;
  let endCursor = null;
  const today = getCurrentDateTimeFormatted();

  while (hasNextPage) {
    const { issues, pageInfo } = fetchIssues(owner, repo, token, endCursor);
    allIssues = allIssues.concat(filterIssues(issues, includeLabels, excludeLabels));
    hasNextPage = pageInfo.hasNextPage;
    endCursor = pageInfo.endCursor;
  }

  writeIssuesToSheet(allIssues, today);
}

//====================
// GitHub Issuesを取得するクエリの実行関数
//====================
function fetchIssues(owner, repo, token, afterCursor = null) {
  const query = `
    query ($owner: String!, $repo: String!, $after: String) {
      repository(owner: $owner, name: $repo) {
        issues(first: 100, after: $after) {
          pageInfo {
            hasNextPage
            endCursor
          }
          nodes {
            number
            title
            createdAt
            closedAt
            labels(first: 10) {
              nodes { name }
            }
            state
            url
            projectItems(first: 1) {
              nodes {
                sprint: fieldValueByName(name: "Sprint") {
                  ... on ProjectV2ItemFieldIterationValue { title }
                }
                status: fieldValueByName(name: "Status") {
                  ... on ProjectV2ItemFieldSingleSelectValue { name }
                }
              }
            }
            repository { name url }
          }
        }
      }
    }
  `;

  const variables = { owner, repo, after: afterCursor };
  const options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({ query, variables }),
  };

  const response = UrlFetchApp.fetch('https://api.github.com/graphql', options);
  const data = JSON.parse(response.getContentText());

  if (data.errors) {
    throw new Error("GraphQL query error.");
  }

  return {
    issues: data.data.repository.issues.nodes,
    pageInfo: data.data.repository.issues.pageInfo,
  };
}

//====================
// Issueフィルタリング関数
//====================
function filterIssues(issues, includeLabels, excludeLabels) {
  return issues.filter(issue => {
    const labels = issue.labels.nodes.map(label => label.name);
    return includeLabels.some(label => labels.includes(label)) &&
           !excludeLabels.some(label => labels.includes(label));
  });
}


//====================
// Issue書き込み関数
//====================
function writeIssuesToSheet(issues, today) {
  const settings = getSettingsFromSheet();
  const sheetName = settings['日次履歴データ取得設定']['日次履歴データ出力シート名'] || 'DailyStatus';
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }

  // ヘッダーがない場合は作成
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Import DateTime", "Issue Number", "Title", "Status", "Labels", 
      "Created At", "Closed At", "Sprint", "Repository", "State", "Issue URL"
    ]);
  }

  // 最新状態をMap化（Issue Number → 最新行データ）
  const lastRow = sheet.getLastRow();
  const latestDataMap = new Map(); // Issue Number をキーとした最新データのマップ
  if (lastRow > 1) {
    const allRows = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // 全データを取得
    allRows.forEach(row => {
      const importDateTime = new Date(row[0]); // Import DateTime
      const issueNumber = row[1]; // Issue Number
      if (issueNumber) {
        // 既存データと比較して最新のデータを保持
        if (!latestDataMap.has(issueNumber) || new Date(latestDataMap.get(issueNumber).row[0]) < importDateTime) {
          latestDataMap.set(issueNumber, { row });
        }
      }
    });
  }

  // データを比較して変更があるもののみ記録
  issues.forEach(issue => {
    const issueNumber = issue.number;
    const title = issue.title;
    const projectStatus = issue.projectItems.nodes.length > 0 && issue.projectItems.nodes[0].status
      ? issue.projectItems.nodes[0].status.name
      : "No Status";
    const sprint = issue.projectItems.nodes.length > 0 && issue.projectItems.nodes[0].sprint
      ? issue.projectItems.nodes[0].sprint.title
      : "No Sprint";
    const createdAt = dateFormat(issue.createdAt);
    const closedAt = issue.closedAt ? dateFormat(issue.closedAt) : "";
    const labels = issue.labels.nodes.map(label => label.name).join(", ");
    const repositoryName = issue.repository.name;
    const state = issue.state; // "OPEN" or "CLOSED"
    const issueUrl = issue.url;

    // 新しい行データを構成
    const newRow = [
      today, issueNumber, title, projectStatus, labels,
      createdAt, closedAt, sprint, repositoryName, state, issueUrl
    ];
    const newRowWithoutImportDateTime = newRow.slice(1).join("|"); // Import DateTime を除いたデータ

    // 最新状態を比較
    const existingData = latestDataMap.get(issueNumber); // 最新行データ
    if (existingData) {
      const existingRowWithoutImportDateTime = existingData.row.slice(1).join("|"); // Import DateTime を除く
      if (existingRowWithoutImportDateTime === newRowWithoutImportDateTime) {
        return; // 完全に一致していれば記録しない
      }
    }

    // 変更がある場合のみ新しいデータを記録
    sheet.appendRow(newRow);
  });
}

//====================
// date format関数
//====================
function dateFormat(utcDateStr) {
  const date = new Date(utcDateStr);
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
}

//====================
// JSTの現在時刻を取得
//====================
function getCurrentDateTimeFormatted() {
  const now = new Date();
  return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}
