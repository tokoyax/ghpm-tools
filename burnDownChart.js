function generateBurndownChart() {
  const settings = getSettingsFromSheet();
  const chartSheetName = settings['バーンダウンチャートデータ生成設定']['バーンダウンチャートデータ出力シート名'] || 'BurndownChart';
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(chartSheetName);

  // BurndownChartシートが存在しない場合は作成
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(chartSheetName);
    Logger.log(`シート "${chartSheetName}" を新規作成しました。`);
    // ヘッダーを設定
    const header = [
      'sprint name', 'sprint start date', 'sprint end date',
      'total story points', 'total issue count',
      'date', 'completed story points', 'remaining story points',
      'completed issue count', 'remaining issue count', 'working day'
    ];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#f4f4f4');
  }

  const sprintName = sheet.getRange('B1').getValue(); // セルB1にスプリント名
  if (!sprintName) {
    Logger.log('スプリント名が設定されていません。');
    return;
  }

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
  const dateIndex = header.indexOf('Import DateTime');

  if (sprintIndex === -1 || estimateIndex === -1 || dateIndex === -1) {
    Logger.log('DailyStatusシートのデータが不足しています。');
    return;
  }

  const sprintStartDate = new Date(sheet.getRange('B2').getValue());
  const sprintEndDate = new Date(sheet.getRange('B3').getValue());

  // 集計用変数
  let totalStoryPoints = 0;
  let totalIssues = 0;
  const burndownData = {};

  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const sprint = row[sprintIndex];
    const estimate = row[estimateIndex] || 0;
    const date = new Date(row[dateIndex]);

    if (sprint === sprintName && date >= sprintStartDate && date <= sprintEndDate) {
      totalStoryPoints += estimate;
      totalIssues += 1;

      const dateKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
      if (!burndownData[dateKey]) {
        burndownData[dateKey] = { completedStoryPoints: 0, completedIssues: 0 };
      }
      burndownData[dateKey].completedStoryPoints += estimate;
      burndownData[dateKey].completedIssues += 1;
    }
  }

  // データをBurndownChartシートに書き込み
  const outputData = [];
  let remainingStoryPoints = totalStoryPoints;
  let remainingIssues = totalIssues;

  Object.keys(burndownData).sort().forEach(date => {
    const data = burndownData[date];
    remainingStoryPoints -= data.completedStoryPoints;
    remainingIssues -= data.completedIssues;
    const isWorkingDay = ![0, 6].includes(new Date(date).getDay()); // 土日判定

    outputData.push([
      date,
      data.completedStoryPoints,
      remainingStoryPoints,
      data.completedIssues,
      remainingIssues,
      isWorkingDay
    ]);
  });

  // 出力データを書き込み
  sheet.getRange(7, 1, outputData.length, outputData[0].length).setValues(outputData);

  // グラフの生成
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(7, 1, outputData.length, 5))
    .setPosition(2, 7, 0, 0)
    .setOption('title', `Burndown Chart - ${sprintName}`)
    .setOption('hAxis', { title: 'Date' })
    .setOption('vAxis', { title: 'Points/Issues' })
    .build();

  sheet.insertChart(chart);
  Logger.log('バーンダウンチャートが作成されました。');
}
