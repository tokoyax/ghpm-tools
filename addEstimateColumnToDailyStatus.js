function addEstimateColumnToDailyStatus() {
  const sheetName = 'DailyStatus';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('DailyStatusシートが存在しません。');
    return;
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!header.includes('Estimate')) {
    const estimateColumnIndex = header.indexOf('Labels') + 2;
    sheet.insertColumnAfter(estimateColumnIndex - 1);
    sheet.getRange(1, estimateColumnIndex).setValue('Estimate');
  }

  const settings = getSettingsFromSheet();
  const token = getGitHubToken();
  const owner = settings['基本設定']['リポジトリのオーナー'];
  const repo = settings['基本設定']['リポジトリ名'];
  const issues = fetchAllIssuesWithSprint(owner, repo, token, 'Sprint 51');

  issues.forEach(issue => {
    const issueNumber = issue.number;
    const estimate = issue.projectItems.nodes.length > 0 && issue.projectItems.nodes[0].estimate
      ? issue.projectItems.nodes[0].estimate.name
      : "";

    const rows = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] == issueNumber) {
        sheet.getRange(i + 2, header.indexOf('Estimate') + 1).setValue(estimate);
        break;
      }
    }
  });

  Logger.log('Estimateカラムの追加とデータ更新が完了しました。');
}

function fetchAllIssuesWithSprint(owner, repo, token, sprintName) {
  let allIssues = [];
  let hasNextPage = true;
  let endCursor = null;

  while (hasNextPage) {
    const { issues, pageInfo } = fetchIssuesPage(owner, repo, token, endCursor);
    allIssues = allIssues.concat(issues);
    hasNextPage = pageInfo.hasNextPage;
    endCursor = pageInfo.endCursor;
  }

  // フィルタリング（スプリント名で絞り込み）
  return allIssues.filter(issue => {
    const sprint = issue.projectItems.nodes.length > 0 && issue.projectItems.nodes[0].sprint
      ? issue.projectItems.nodes[0].sprint.title
      : "";
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
                  ... on ProjectV2ItemFieldIterationValue { title }
                }
                estimate: fieldValueByName(name: "Estimate") {
                  ... on ProjectV2ItemFieldSingleSelectValue { name }
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

  return {
    issues: data.data.repository.issues.nodes,
    pageInfo: data.data.repository.issues.pageInfo,
  };
}
