// 定数定義
const HEADERS = {
  SPRINT: ['Sprint Name', 'Sprint Start Date', 'Sprint End Date', 'Total Story Points', 'Total Issue Count'],
  BURNDOWN: ['Date', 'Completed Story Points', 'Remaining Story Points', 'Completed Issue Count', 'Remaining Issue Count', 'Is Working Day'],
  COMPLETION_LOG: ['Date', 'Issue Number', 'Issue Title', 'Status', 'Story Points', 'Issue URL']
};

const STYLE = {
  HEADER_BACKGROUND: '#f4f4f4'
};

function generateBurndownChartWithCompletionLog() {
  const settings = getSettingsFromSheet();
  const chartSheetName = settings['バーンダウンチャートデータ生成設定']['バーンダウンチャートデータ出力シート名'] || 'BurndownChart';
  const sprintName = settings['バーンダウンチャートデータ生成設定']['対象Sprint名'];
  const completedStatuses = (settings['バーンダウンチャートデータ生成設定']['完了ステータス'] || "").split(',').map(s => s.trim());
  const token = getGitHubToken();
  const owner = settings['基本設定']['リポジトリのオーナー'];
  const repo = settings['基本設定']['リポジトリ名'];

  if (!sprintName || completedStatuses.length === 0) {
    Logger.log('対象スプリント名または完了ステータスが設定されていません。');
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(chartSheetName);

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(chartSheetName);
    const sprintHeader = [['Sprint Name'], ['Sprint Start Date'], ['Sprint End Date'], ['Total Story Points'], ['Total Issue Count']];
    const burndownHeader = ['Date', 'Completed Story Points', 'Remaining Story Points', 'Completed Issue Count', 'Remaining Issue Count', 'Is Working Day'];
    sheet.getRange(1, 1, sprintHeader.length, 1).setValues(sprintHeader);
    sheet.getRange(7, 1, 1, burndownHeader.length).setValues([burndownHeader]);
    sheet.getRange('A1:A5').setFontWeight('bold').setBackground('#f4f4f4');
    sheet.getRange('A7:F7').setFontWeight('bold').setBackground('#f4f4f4');
  }

  const sprintDates = fetchSprintDates(owner, repo, sprintName, token);
  if (!sprintDates) {
    Logger.log('スプリントの日付が取得できませんでした。');
    return;
  }

  sheet.getRange('B2').setValue(sprintDates.startDate);
  sheet.getRange('B3').setValue(sprintDates.endDate);

  const dailyStatusSheetName = settings['日次履歴データ取得設定']['日次履歴データ出力シート名'] || 'DailyStatus';
  const dailyStatusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dailyStatusSheetName);
  if (!dailyStatusSheet) {
    Logger.log('DailyStatusシートが見つかりません。');
    return;
  }

  const dataRange = dailyStatusSheet.getDataRange().getValues();
  const header = dataRange[0];
  const sprintIndex = header.indexOf('Sprint');
  const estimateIndex = header.indexOf('Estimate');
  const statusIndex = header.indexOf('Status');
  const dateIndex = header.indexOf('Import DateTime');
  const issueNumberIndex = header.indexOf('Issue Number');
  const titleIndex = header.indexOf('Title');

  if ([sprintIndex, estimateIndex, statusIndex, dateIndex, issueNumberIndex, titleIndex].some(idx => idx === -1)) {
    Logger.log('DailyStatusシートのデータが不足しています。');
    return;
  }

  const sprintStartDate = new Date(sprintDates.startDate);
  const sprintEndDate = new Date(sprintDates.endDate);
  const totalDays = Math.ceil((sprintEndDate - sprintStartDate) / (1000 * 60 * 60 * 24)) + 1;

  let totalStoryPoints = 0;
  let totalIssues = 0;
  const dailyData = Array(totalDays).fill(null).map(() => ({ 
    completedStoryPoints: 0, 
    completedIssues: 0,
    completedIssueNumbers: new Set() // 完了した課題番号を追跡
  }));

  // まず、全課題の合計を計算
  const issueTracker = new Set();
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const sprint = row[sprintIndex];
    const estimate = row[estimateIndex] || 0;
    const issueNumber = row[issueNumberIndex];

    if (!issueNumber || sprint !== sprintName) continue;

    if (!issueTracker.has(issueNumber)) {
      totalStoryPoints += estimate;
      totalIssues += 1;
      issueTracker.add(issueNumber);
    }
  }

  // 完了状態の計算を修正
  const completedIssues = new Map(); // 課題の最新状態を追跡
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const sprint = row[sprintIndex];
    const estimate = row[estimateIndex] || 0;
    const status = row[statusIndex];
    const importDate = new Date(row[dateIndex]);
    const issueNumber = row[issueNumberIndex];
    const title = row[titleIndex];

    if (!issueNumber || sprint !== sprintName || importDate < sprintStartDate || importDate > sprintEndDate) continue;

    // インポート日時から日付部分のみを取得
    const currentDate = new Date(importDate.getFullYear(), importDate.getMonth(), importDate.getDate());
    const dayIndex = Math.floor((currentDate - sprintStartDate) / (1000 * 60 * 60 * 24));

    // 完了判定
    const isCompleted = completedStatuses.some(completedStatus => 
      status.trim() === completedStatus.trim()
    );

    // その日の状態を記録（同じ日のデータは上書き）
    completedIssues.set(issueNumber, {
      isCompleted: isCompleted,
      dayIndex: dayIndex,
      estimate: estimate,
      title: title,
      status: status,
      importDate: currentDate
    });
  }

  // 日ごとの完了状態を計算
  for (let day = 0; day < totalDays; day++) {
    const currentDate = new Date(sprintStartDate);
    currentDate.setDate(currentDate.getDate() + day);
    
    // その日までに完了している課題を集計
    completedIssues.forEach((issue, number) => {
      // その日と同じかそれ以前の完了課題をカウント（日付の完全一致）
      const issueDate = new Date(issue.importDate);
      const targetDate = new Date(currentDate);
      
      // 日付を比較（年月日のみ）
      if (issue.isCompleted && 
          issueDate.getFullYear() === targetDate.getFullYear() &&
          issueDate.getMonth() === targetDate.getMonth() &&
          issueDate.getDate() <= targetDate.getDate()) {
        dailyData[day].completedIssueNumbers.add(number);
      }
    });

    // その日の完了課題数とストーリーポイントを計算
    dailyData[day].completedIssues = dailyData[day].completedIssueNumbers.size;
    dailyData[day].completedStoryPoints = Array.from(dailyData[day].completedIssueNumbers)
      .reduce((sum, num) => sum + (completedIssues.get(num)?.estimate || 0), 0);
  }

  // 出力データの生成
  let remainingStoryPoints = totalStoryPoints;
  let remainingIssues = totalIssues;
  const outputData = dailyData.map((data, index) => {
    const date = new Date(sprintStartDate);
    date.setDate(date.getDate() + index);
    const isWorkingDay = ![0, 6].includes(date.getDay());

    remainingStoryPoints = totalStoryPoints - data.completedStoryPoints;
    remainingIssues = totalIssues - data.completedIssues;

    return [
      Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd'),
      data.completedStoryPoints,
      remainingStoryPoints,
      data.completedIssues,
      remainingIssues,
      isWorkingDay,
    ];
  });

  // データ出力
  sheet.getRange(8, 1, outputData.length, outputData[0].length).setValues(outputData);

  // スプリント情報のヘッダーと値を設定
  const sprintHeader = [['Sprint Name'], ['Sprint Start Date'], ['Sprint End Date'], ['Total Story Points'], ['Total Issue Count']];
  sheet.getRange(1, 1, sprintHeader.length, 1).setValues(sprintHeader);
  sheet.getRange(1, 2).setValue(sprintName);
  sheet.getRange(2, 2).setValue(sprintDates.startDate);
  sheet.getRange(3, 2).setValue(sprintDates.endDate);
  sheet.getRange(4, 2).setValue(totalStoryPoints);
  sheet.getRange(5, 2).setValue(totalIssues);

  // ヘッダーの書式設定
  sheet.getRange('A1:A5').setFontWeight('bold').setBackground('#f4f4f4');

  // 完了ログのヘッダーを定義
  const completionLogHeader = ['Date', 'Issue Number', 'Issue Title', 'Status', 'Story Points', 'Issue URL'];
  
  // 完了ログの作成
  const completionLog = Array.from(completedIssues.entries())
    .filter(([_, issue]) => issue.isCompleted)
    .map(([number, issue]) => [
      Utilities.formatDate(issue.importDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
      number,
      issue.title,
      issue.status,
      issue.estimate,
      `https://github.com/${owner}/${repo}/issues/${number}`
    ]);

  // 完了ログの出力
  const completionLogStartRow = 8 + outputData.length + 2;
  sheet.getRange(completionLogStartRow, 1, 1, completionLogHeader.length)
    .setValues([completionLogHeader])
    .setFontWeight('bold')
    .setBackground('#f4f4f4');  // スプリントヘッダーと同じ背景色を設定
  if (completionLog.length > 0) {
    sheet.getRange(completionLogStartRow + 1, 1, completionLog.length, completionLog[0].length).setValues(completionLog);
  }

  // グラフの追加
  addBurndownChart(sheet, { dailyData, sprintName, sprintDates });

  Logger.log('バーンダウンチャート用データと完了課題ログを作成しました。');
}

function fetchSprintDates(owner, repo, sprintName, token) {
  const query = `
    query ($owner: String!, $repo: String!) {
      repository(owner: $owner, name: $repo) {
        projectsV2(first: 10) {
          nodes {
            title
            items(first: 100) {
              nodes {
                fieldValues(first: 10) {
                  nodes {
                    __typename
                    ... on ProjectV2ItemFieldIterationValue {
                      title
                      startDate
                      duration
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  `;

  const variables = { owner, repo };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ query, variables }),
  };

  const response = UrlFetchApp.fetch('https://api.github.com/graphql', options);
  const data = JSON.parse(response.getContentText());

  if (data.errors) {
    Logger.log(`GraphQL errors: ${JSON.stringify(data.errors)}`);
    return null;
  }

  const projects = data.data.repository.projectsV2.nodes;

  for (const project of projects) {
    for (const item of project.items.nodes) {
      for (const field of item.fieldValues.nodes) {
        if (field.__typename === "ProjectV2ItemFieldIterationValue" && field.title === sprintName) {
          const startDate = field.startDate;
          const duration = field.duration; // 日数として取得
          if (startDate && duration) {
            const endDate = new Date(new Date(startDate).getTime() + (duration - 1) * 24 * 60 * 60 * 1000); // duration - 1
            return { startDate, endDate: endDate.toISOString().split('T')[0] }; // 日付部分のみ返す
          }
        }
      }
    }
  }

  Logger.log('指定されたスプリント名のデータが見つかりませんでした。');
  return null;
}

function addBurndownChart(sheet, data) {
  // データ範囲を取得（日付と残りのストーリーポイントのみ）
  const dateRange = sheet.getRange(8, 1, data.dailyData.length, 1); // A列：日付
  const remainingPointsRange = sheet.getRange(8, 3, data.dailyData.length, 1); // C列：残りのストーリーポイント
  
  // 理想線のデータを作成
  const totalDays = data.dailyData.length;
  const startPoints = sheet.getRange('B4').getValue(); // Total Story Points from header
  
  // 稼働日数を計算（週末を除く）
  let workingDays = 0;
  const idealData = Array(totalDays).fill().map((_, index) => {
    const currentDate = new Date(data.sprintDates.startDate);
    currentDate.setDate(currentDate.getDate() + index);
    const isWeekend = [0, 6].includes(currentDate.getDay()); // 0=日曜, 6=土曜
    
    if (!isWeekend) {
      workingDays++;
    }
    return isWeekend;
  });

  // 1稼働日あたりの減少ポイントを計算
  const pointsPerWorkingDay = startPoints / workingDays;
  
  // 理想線データを生成
  let remainingPoints = startPoints;
  const idealBurndown = idealData.map((isWeekend, index) => {
    if (index === 0) return [startPoints]; // 初日
    if (isWeekend) return [remainingPoints]; // 週末は前日と同じ
    
    remainingPoints -= pointsPerWorkingDay;
    return [Math.max(0, remainingPoints)]; // マイナスにならないよう0で制限
  });
  
  // G列のヘッダーを設定
  sheet.getRange(7, 7).setValue('Ideal Burndown')
    .setFontWeight('bold')
    .setBackground(STYLE.HEADER_BACKGROUND);
  
  // G列に理想線データを配置
  const idealRange = sheet.getRange(8, 7, totalDays, 1);
  idealRange.setValues(idealBurndown);
  
  // グラフを作成
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dateRange)      // X軸の日付
    .addRange(remainingPointsRange)  // 実際の残りのストーリーポイント
    .addRange(idealRange)     // 理想線
    .setPosition(1, 8, 0, 0)  // 1行目、H列から配置
    .setOption('title', `${data.sprintName} Burndown Chart`)
    .setOption('series', {
      0: { // 実際の線（オレンジ）
        targetAxisIndex: 0,
        labelInLegend: 'Remaining Story Points',
        color: '#ed6c02',
        lineWidth: 3,
        pointSize: 0
      },
      1: { // 理想線（グレー点線）
        targetAxisIndex: 0,
        labelInLegend: 'Ideal Burndown',
        color: '#666666',
        lineWidth: 1,
        lineDashType: 'dot'  // 点線パターン（'solid', 'dot', 'mediumDash', 'mediumDashDot', 'longDash', 'longDashDot'から選択）
      }
    })
    .setOption('hAxis', {
      title: 'Date',
      format: 'yyyy/MM/dd'
    })
    .setOption('vAxis', {
      title: 'Story Points',
      minValue: 0
    })
    .setOption('legend', {
      position: 'top'
    })
    .build();

  // デバッグ用のログ出力
  Logger.log('Total Days:', totalDays);
  Logger.log('Working Days:', workingDays);
  Logger.log('Start Points:', startPoints);
  Logger.log('Points per Working Day:', pointsPerWorkingDay);
  Logger.log('Ideal Burndown:', idealBurndown);

  // 既存のグラフがあれば削除
  const charts = sheet.getCharts();
  charts.forEach(existingChart => sheet.removeChart(existingChart));
  
  // 新しいグラフを追加
  sheet.insertChart(chart);
}

function outputToSheet(sheet, data) {
  // ... existing outputToSheet code ...

  // グラフの追加
  addBurndownChart(sheet, data);
}