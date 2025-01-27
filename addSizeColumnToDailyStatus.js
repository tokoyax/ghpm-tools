function addSizeColumnToDailyStatus() {
  const sheetName = 'DailyStatus';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('DailyStatusシートが存在しません。');
    return;
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (header.includes("Size")) {
    SpreadsheetApp.getUi().alert('Sizeカラムは既に存在します。');
    return;
  }

  const sizeColumnIndex = header.indexOf("Labels") + 2;
  sheet.insertColumnAfter(sizeColumnIndex - 1);
  sheet.getRange(1, sizeColumnIndex).setValue("Size");

  const settings = getSettingsFromSheet();
  const token = getGitHubToken();
  const owner = settings['基本設定']['リポジトリのオーナー'];
  const repo = settings['基本設定']['リポジトリ名'];
  const issues = fetchIssuesWithSprint(owner, repo, token, "Sprint51");

  issues.forEach(issue => {
    const issueNumber = issue.number;
    const size = issue.projectItems.nodes.length > 0 && issue.projectItems.nodes[0].size
      ? issue.projectItems.nodes[0].size.name : "";

    const rows = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === issueNumber) {
        sheet.getRange(i + 2, sizeColumnIndex).setValue(size);
        break;
      }
    }
  });

  SpreadsheetApp.getUi().alert('Sizeカラムを追加し、データを更新しました。');
}

function fetchIssuesWithSprint(owner, repo, token, sprintName) {
  const query = `
    query ($owner: String!, $repo: String!, $sprintName: String!) {
      repository(owner: $owner, name: $repo) {
        issues(first: 100) {
          nodes {
            number
            projectItems(first: 1) {
              nodes {
                sprint: fieldValueByName(name: $sprintName) {
                  ... on ProjectV2ItemFieldIterationValue { title }
                }
                size: fieldValueByName(name: "Size") {
                  ... on ProjectV2ItemFieldSingleSelectValue { name }
                }
              }
            }
          }
        }
      }
    }
  `;
  const variables = { owner, repo, sprintName };
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

  return data.data.repository.issues.nodes.filter(issue => {
    const sprint = issue.projectItems.nodes[0]?.sprint?.title || "";
    return sprint === sprintName;
  });
}
