function addEstimateColumnToDailyStatus() {
  const sheetName = 'DailyStatus';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('DailyStatusシートが存在しません。');
    return;
  }

  let header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // "Estimate" カラムが存在しない場合は追加
  if (!header.includes('Estimate')) {
    const estimateColumnIndex = header.indexOf('Labels') + 2;
    sheet.insertColumnAfter(estimateColumnIndex - 1);
    sheet.getRange(1, estimateColumnIndex).setValue('Estimate');
    header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // 再取得
  }

  const settings = getSettingsFromSheet();
  const token = getGitHubToken();
  const owner = settings['基本設定']['リポジトリのオーナー'];
  const repo = settings['基本設定']['リポジトリ名'];
  const targetSprint = 'Sprint 51';
  const issues = fetchAllIssuesWithSprint(owner, repo, token, targetSprint);

  const sprintColumnIndex = header.indexOf('Sprint') + 1; // Sprint列のインデックス
  const estimateColumnIndex = header.indexOf('Estimate') + 1;

  // 既存の行データを取得
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, header.length).getValues();

  issues.forEach(issue => {
    const issueNumber = issue.number;
    const estimate = issue.projectItems.nodes[0]?.estimate?.number || ""; // 数値型として扱う

    rows.forEach((row, i) => {
      const sheetIssueNumber = row[0]; // "Issue Number"列の値
      const sheetSprint = row[sprintColumnIndex - 1]; // "Sprint"列の値

      if (sheetIssueNumber == issueNumber && sheetSprint === targetSprint) {
        sheet.getRange(i + 2, estimateColumnIndex).setValue(estimate); // 該当セルにEstimate値を設定
        Logger.log(`Updated Issue Number: ${issueNumber}, Sprint: ${sheetSprint}, Estimate: ${estimate}`);
      }
    });
  });

  Logger.log('Estimateカラムの追加とデータ更新が完了しました。');
}

function fetchAllIssuesWithSprint(owner, repo, token, sprintName) {
  let allIssues = [];
  let hasNextPage = true;
  let endCursor = null;

  while (hasNextPage) {
    const { issues, pageInfo } = fetchIssuesPage(owner, repo, token, endCursor);
    if (!issues) {
      Logger.log('Issues not found in the current page.');
      break;
    }
    allIssues = allIssues.concat(issues);
    hasNextPage = pageInfo.hasNextPage;
    endCursor = pageInfo.endCursor;
  }

  // デバッグ用ログ出力
  allIssues.forEach(issue => {
    const sprintTitle = issue.projectItems.nodes[0]?.sprint?.title || "N/A";
    Logger.log(`Issue Number: ${issue.number}, Sprint: ${sprintTitle}`);
  });

  // 指定したSprintでフィルタリング
  return allIssues.filter(issue => {
    const sprint = issue.projectItems.nodes[0]?.sprint?.title || "";
    return sprint === sprintName;
  });
}

function fetchIssuesPage(owner, repo, token, afterCursor = null) {
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
            projectItems(first: 1) {
              nodes {
                sprint: fieldValueByName(name: "Sprint") {
                  ... on ProjectV2ItemFieldIterationValue {
                    title
                  }
                }
                estimate: fieldValueByName(name: "Estimate") {
                  ... on ProjectV2ItemFieldNumberValue {
                    number
                  }
                }
              }
            }
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
    Logger.log(`GraphQL errors: ${JSON.stringify(data.errors)}`);
    throw new Error("GraphQL query error.");
  }

  Logger.log(`Fetched ${data.data.repository.issues.nodes.length} issues in current page.`);
  return {
    issues: data.data.repository.issues.nodes,
    pageInfo: data.data.repository.issues.pageInfo,
  };
}
