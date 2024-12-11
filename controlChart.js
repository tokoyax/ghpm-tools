// ========================
// 設定箇所（ユーザーが編集）
// ========================
const SETTINGS = {
  START_DATE: '2024/11/01',
  END_DATE: '2024/12/31',
  FROM_STATUS: 'In Progress',
  TO_STATUSES: ['Done', 'QA確認中', 'QA確認完了', 'リリース作業中 🚀', 'リリース完了 🌝', 'Closed 🚫'],
  //FROM_STATUS: 'Epic準備中',
  //TO_STATUSES: ['Epic Backlog'],
  INCLUDE_LABELS: ['UserStory', 'Epic', 'Task', 'BugFix'],
  CYCLETIME_SHEET: 'CycleTime',
  DAILYSTATUS_SHEET: 'DailyStatus',
  CONTROL_CHART_SHEET: 'ControlChart',
  EMA_ALPHA: 0.2 // EMAのスムージング係数
};

// ========================
// エントリーポイント
// ========================
function execute() {
  generateCycleTimeSheet();
  createControlChart();
}

// ========================
// サイクルタイムデータを生成する
// ========================
function generateCycleTimeSheet() {
  const dailyStatusSheet = getSheetByNameOrCreate(SETTINGS.DAILYSTATUS_SHEET);
  const cycleTimeSheet = getSheetByNameOrCreate(SETTINGS.CYCLETIME_SHEET);

  cycleTimeSheet.clear();
  cycleTimeSheet.appendRow([
    "From Status Date", "To Status Date", "Issue Number", "Title", 
    "From-To Status Time (秒)", "Label", "Created At", "Closed At", "Sprint", "Issue URL"
  ]);

  const dailyStatusData = dailyStatusSheet.getDataRange().getValues();
  const headers = dailyStatusData[0];
  const indices = {
    status: headers.indexOf("Status"),
    date: headers.indexOf("Import DateTime"),
    issueNumber: headers.indexOf("Issue Number"),
    title: headers.indexOf("Title"),
    labels: headers.indexOf("Labels"),
    createdAt: headers.indexOf("Created At"),
    closedAt: headers.indexOf("Closed At"),
    sprint: headers.indexOf("Sprint"),
    issueUrl: headers.indexOf("Issue URL")
  };

  let issueHistory = {};

  for (let i = 1; i < dailyStatusData.length; i++) {
    const row = dailyStatusData[i];
    const issueNumber = row[indices.issueNumber];
    const status = row[indices.status];
    const dateStr = row[indices.date];
    const date = new Date(dateStr);

    if (date < new Date(SETTINGS.START_DATE) || date > new Date(SETTINGS.END_DATE)) continue;

    if (!issueHistory[issueNumber]) issueHistory[issueNumber] = [];
    issueHistory[issueNumber].push({ status, date, row });
  }

  Object.keys(issueHistory).forEach(issueNumber => {
    const history = issueHistory[issueNumber];
    let firstFromStatusDate = null;
    let lastToStatusDate = null;
    let relevantRow = null;

    history.forEach(event => {
      if (event.status === SETTINGS.FROM_STATUS && firstFromStatusDate === null) {
        firstFromStatusDate = event.date;
      }
      if (SETTINGS.TO_STATUSES.includes(event.status)) {
        lastToStatusDate = event.date;
        relevantRow = event.row;
      }
    });

    if (firstFromStatusDate && lastToStatusDate) {
      const cycleTimeSeconds = (lastToStatusDate - firstFromStatusDate) / 1000;
      cycleTimeSheet.appendRow([
        Utilities.formatDate(firstFromStatusDate, "JST", "yyyy/MM/dd HH:mm:ss"),
        Utilities.formatDate(lastToStatusDate, "JST", "yyyy/MM/dd HH:mm:ss"),
        issueNumber,
        relevantRow[indices.title],
        cycleTimeSeconds,
        relevantRow[indices.labels],
        relevantRow[indices.createdAt],
        relevantRow[indices.closedAt],
        relevantRow[indices.sprint],
        relevantRow[indices.issueUrl]
      ]);
    }
  });
}

function createControlChart() {
  const cycleTimeSheet = getSheetByNameOrCreate(SETTINGS.CYCLETIME_SHEET);
  const chartSheet = getSheetByNameOrCreate(SETTINGS.CONTROL_CHART_SHEET);

  chartSheet.clear();
  chartSheet.appendRow([
    "Date", "Cycle Time (Days)", "Issue Number", "Issue Title", "Issue URL",
    "From Status Date", "To Status Date", "AVG", "EMA", "StandardDeviation upperBand", "StandardDeviation lowerBand"
  ]);

  const dataRange = cycleTimeSheet.getDataRange();
  const values = dataRange.getValues().slice(1);

  const allCycleTimes = values.map(row => row[4] / (60 * 60 * 24)); // 全Issueのサイクルタイム（日数）
  const overallAvg = calculateAverage(allCycleTimes); // 全体の平均

  let ema = null; // 初期EMAはnull
  let emaStdDev = calculateStandardDeviation(allCycleTimes, overallAvg); // 初期標準偏差を全体の標準偏差で設定

  const rollingCycleTimes = [];
  const dailyData = {};
  values.forEach(row => {
    const date = Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd"); // From Status Date
    const cycleTimeDays = row[4] / (60 * 60 * 24); // 秒を日数に変換
    const issueInfo = {
      issueNumber: row[2],
      title: row[3],
      url: row[9],
      cycleTime: cycleTimeDays,
      fromDateTime: row[0],
      toDateTime: row[1]
    };

    if (!dailyData[date]) {
      dailyData[date] = [];
    }
    dailyData[date].push(issueInfo);
  });

  const sortedDates = generateDateRange(new Date(SETTINGS.START_DATE), new Date(SETTINGS.END_DATE));

  sortedDates.forEach(dateObj => {
    const date = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd");
    const dailyIssues = dailyData[date] || [];
    const dailyTimes = dailyIssues.map(issue => issue.cycleTime);
    const dailyAvg = calculateAverage(dailyTimes);

    // データが存在する場合のみ追加
    if (dailyTimes.length > 0) {
      rollingCycleTimes.push(...dailyTimes);
    }

    // EMAの計算
    ema = ema === null
      ? dailyAvg // 初回EMAはその日の平均を基準とする
      : SETTINGS.EMA_ALPHA * dailyAvg + (1 - SETTINGS.EMA_ALPHA) * ema;

    // 移動標準偏差のEMAを計算
    const recentValues = rollingCycleTimes.slice(-5); // 直近5件のデータ
    const recentStdDev = recentValues.length > 1
      ? calculateStandardDeviation(recentValues, calculateAverage(recentValues))
      : emaStdDev; // データ不足の場合は直前のEMA標準偏差を使用
    emaStdDev = SETTINGS.EMA_ALPHA * recentStdDev + (1 - SETTINGS.EMA_ALPHA) * emaStdDev; // EMAで更新

    const upperBand = ema + 2 * emaStdDev;
    const lowerBand = Math.max(ema - 2 * emaStdDev, 0); // 下限は0以上

    if (dailyIssues.length > 0) {
      dailyIssues.forEach(issue => {
        chartSheet.appendRow([
          date, issue.cycleTime, issue.issueNumber, issue.title, issue.url,
          Utilities.formatDate(new Date(issue.fromDateTime), "JST", "yyyy/MM/dd HH:mm:ss"),
          Utilities.formatDate(new Date(issue.toDateTime), "JST", "yyyy/MM/dd HH:mm:ss"),
          overallAvg, ema, upperBand, lowerBand
        ]);
      });
    } else {
      chartSheet.appendRow([
        date, null, null, null, null, null, null,
        overallAvg, ema, upperBand, lowerBand
      ]);
    }
  });
}

// ========================
// ユーティリティ関数
// ========================
function getSheetByNameOrCreate(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function calculateAverage(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function calculateStandardDeviation(values, avg) {
  if (values.length <= 1) return 0; // 標準偏差の計算に十分なデータがない場合は0を返す
  const variance = values.reduce((sum, value) => sum + Math.pow(value - avg, 2), 0) / values.length;
  return Math.sqrt(variance);
}

function generateDateRange(startDate, endDate) {
  const dates = [];
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    dates.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return dates;
}