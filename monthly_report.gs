/**
 * Webサイト月次レポート自動化ツール
 *
 * 初回設定:
 * 1. Googleスプレッドシートで「拡張機能 > Apps Script」を開く
 * 2. このファイルの内容を貼り付ける
 * 3. Apps Script左メニュー「サービス」から次の高度なGoogleサービスを追加する
 *    - Google Analytics Data API: 識別子 AnalyticsData
 *    - Google Search Console API: 識別子 SearchConsole
 * 4. Google Cloud側でも同じAPIを有効化する
 * 5. setupMonthlyReportSheets() を1回実行して、設定シートなどを作成する
 *
 * GA4プロパティIDは数値だけを入力してください。例: 123456789
 * Search ConsoleのサイトURLは登録済みプロパティと完全一致させてください。
 * URLプレフィックスなら https://example.com/、ドメインプロパティなら sc-domain:example.com
 */

const SHEET_NAMES = {
  settings: '設定',
  summary: '月次サマリー',
  queries: '検索キーワード',
  pages: '人気ページ',
  report: 'レポート',
};

const SETTINGS = {
  siteName: 'サイト名',
  ga4PropertyId: 'GA4プロパティID',
  searchConsoleSiteUrl: 'Search ConsoleのサイトURL',
  inquirySheetName: '問い合わせ管理シート名',
  reportMonth: 'レポート対象月',
};

const SUMMARY_HEADERS = [
  '月',
  'ユーザー数',
  '新規ユーザー数',
  'セッション数',
  '検索表示回数',
  '検索クリック数',
  'CTR',
  '平均掲載順位',
  '問い合わせ数',
  '人気ページ',
  '改善メモ',
];

const QUERY_HEADERS = ['月', 'キーワード', '表示回数', 'クリック数', 'CTR', '平均掲載順位'];
const PAGE_HEADERS = ['月', 'ページURL', '閲覧数', 'メモ'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('月次レポート')
    .addItem('今月分を作成', 'createCurrentMonthReport')
    .addItem('前月分を作成', 'createPreviousMonthReport')
    .addSeparator()
    .addItem('初期シートを作成', 'setupMonthlyReportSheets')
    .addToUi();
}

function setupMonthlyReportSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSettingsSheet_(ss);
  setupSheet_(ss, SHEET_NAMES.summary, SUMMARY_HEADERS);
  setupSheet_(ss, SHEET_NAMES.queries, QUERY_HEADERS);
  setupSheet_(ss, SHEET_NAMES.pages, PAGE_HEADERS);
  setupReportSheet_(ss);

  SpreadsheetApp.getUi().alert('月次レポート用のシートを作成しました。設定シートを入力してください。');
}

function createCurrentMonthReport() {
  createMonthlyReport(formatMonth_(new Date()));
}

function createPreviousMonthReport() {
  const today = new Date();
  const previousMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  createMonthlyReport(formatMonth_(previousMonth));
}

/**
 * 月次レポートを作成します。
 * targetMonthを省略すると、設定シートの「レポート対象月」を使います。
 *
 * @param {string=} targetMonth yyyy-MM形式。例: 2026-06
 */
function createMonthlyReport(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRequiredSheets_(ss);

  const config = getConfig_();
  const month = targetMonth || config.reportMonth || formatMonth_(new Date());
  const period = getMonthPeriod_(month);

  const ga4Data = fetchGA4Data(config.ga4PropertyId, period.startDate, period.endDate);
  const searchConsoleData = fetchSearchConsoleData(
    config.searchConsoleSiteUrl,
    period.startDate,
    period.endDate
  );
  const inquiryCount = fetchInquiryCount(config.inquirySheetName, period.startDate, period.endDate);

  writeKeywordRows_(month, searchConsoleData.queries);
  writePageRows_(month, ga4Data.pages);
  writeMonthlySummary(month, ga4Data, searchConsoleData, inquiryCount);

  const reportText = generateReportText(month, config.siteName, ga4Data, searchConsoleData, inquiryCount);
  writeReportSheet_(month, config.siteName, reportText, ga4Data, searchConsoleData, inquiryCount);

  SpreadsheetApp.getUi().alert(month + ' の月次レポートを作成しました。');
}

/**
 * GA4からユーザー数、セッション数、人気ページを取得します。
 *
 * API設定:
 * Apps Scriptの「サービス」で Google Analytics Data API を追加してください。
 * 識別子は AnalyticsData のままにします。
 *
 * @param {string} propertyId GA4プロパティID。例: 123456789
 * @param {string} startDate yyyy-MM-dd
 * @param {string} endDate yyyy-MM-dd
 * @return {{users:number,newUsers:number,sessions:number,pages:Array<{url:string,views:number}>}}
 */
function fetchGA4Data(propertyId, startDate, endDate) {
  if (!propertyId) {
    throw new Error('設定シートにGA4プロパティIDを入力してください。');
  }

  const propertyName = 'properties/' + propertyId;

  const overviewRequest = {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    metrics: [{ name: 'totalUsers' }, { name: 'newUsers' }, { name: 'sessions' }],
  };
  const overviewResponse = AnalyticsData.Properties.runReport(overviewRequest, propertyName);
  const metrics = overviewResponse.rows && overviewResponse.rows.length
    ? overviewResponse.rows[0].metricValues
    : [];

  const pageRequest = {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: 'pagePathPlusQueryString' }],
    metrics: [{ name: 'screenPageViews' }],
    orderBys: [{ metric: { metricName: 'screenPageViews' }, desc: true }],
    limit: 10,
  };
  const pageResponse = AnalyticsData.Properties.runReport(pageRequest, propertyName);
  const pages = (pageResponse.rows || []).map(function(row) {
    return {
      url: row.dimensionValues[0].value,
      views: Number(row.metricValues[0].value || 0),
    };
  });

  return {
    users: Number(metrics[0] ? metrics[0].value : 0),
    newUsers: Number(metrics[1] ? metrics[1].value : 0),
    sessions: Number(metrics[2] ? metrics[2].value : 0),
    pages: pages,
  };
}

/**
 * Search Consoleから検索実績と検索キーワードを取得します。
 *
 * API設定:
 * Apps Scriptの「サービス」で Google Search Console API を追加してください。
 * 識別子は SearchConsole のままにします。
 *
 * @param {string} siteUrl Search Consoleのプロパティ。例: https://example.com/ または sc-domain:example.com
 * @param {string} startDate yyyy-MM-dd
 * @param {string} endDate yyyy-MM-dd
 * @return {{impressions:number,clicks:number,ctr:number,position:number,queries:Array}}
 */
function fetchSearchConsoleData(siteUrl, startDate, endDate) {
  if (!siteUrl) {
    throw new Error('設定シートにSearch ConsoleのサイトURLを入力してください。');
  }

  const totalRequest = {
    startDate: startDate,
    endDate: endDate,
    dimensions: [],
    rowLimit: 1,
  };
  const totalResponse = SearchConsole.Searchanalytics.query(totalRequest, siteUrl);
  const totalRow = totalResponse.rows && totalResponse.rows.length ? totalResponse.rows[0] : null;

  const queryRequest = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ['query'],
    rowLimit: 50,
    startRow: 0,
  };
  const queryResponse = SearchConsole.Searchanalytics.query(queryRequest, siteUrl);
  const queries = (queryResponse.rows || []).map(function(row) {
    return {
      keyword: row.keys[0],
      impressions: Number(row.impressions || 0),
      clicks: Number(row.clicks || 0),
      ctr: Number(row.ctr || 0),
      position: Number(row.position || 0),
    };
  });

  return {
    impressions: totalRow ? Number(totalRow.impressions || 0) : 0,
    clicks: totalRow ? Number(totalRow.clicks || 0) : 0,
    ctr: totalRow ? Number(totalRow.ctr || 0) : 0,
    position: totalRow ? Number(totalRow.position || 0) : 0,
    queries: queries,
  };
}

/**
 * 問い合わせ管理シートから対象月の問い合わせ数を数えます。
 * 1行目に「日付」「日時」「受付日」「送信日時」「created_at」のいずれかを含む列がある想定です。
 *
 * @param {string} inquirySheetName 問い合わせ管理シート名
 * @param {string} startDate yyyy-MM-dd
 * @param {string} endDate yyyy-MM-dd
 * @return {number}
 */
function fetchInquiryCount(inquirySheetName, startDate, endDate) {
  if (!inquirySheetName) {
    return 0;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(inquirySheetName);
  if (!sheet) {
    return 0;
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return 0;
  }

  const dateColumnIndex = findDateColumnIndex_(values[0]);
  if (dateColumnIndex === -1) {
    throw new Error('問い合わせ管理シートの1行目に日付列を作成してください。例: 日付、日時、受付日');
  }

  const start = parseDate_(startDate);
  const end = parseDate_(endDate);
  end.setHours(23, 59, 59, 999);

  return values.slice(1).filter(function(row) {
    const date = normalizeDateValue_(row[dateColumnIndex]);
    return date && date >= start && date <= end;
  }).length;
}

/**
 * 月次サマリーシートへ1行書き込みます。
 * 同じ月の行がある場合は上書きします。
 */
function writeMonthlySummary(month, ga4Data, searchConsoleData, inquiryCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.summary);
  const topPages = ga4Data.pages.slice(0, 3).map(function(page) {
    return page.url + '（' + page.views + '）';
  }).join('\n');

  const row = [
    month,
    ga4Data.users,
    ga4Data.newUsers,
    ga4Data.sessions,
    searchConsoleData.impressions,
    searchConsoleData.clicks,
    searchConsoleData.ctr,
    searchConsoleData.position,
    inquiryCount,
    topPages,
    '',
  ];

  upsertRowByMonth_(sheet, month, row);
  formatSummarySheet_(sheet);
}

/**
 * レポートシートに貼り付ける文章を作ります。
 */
function generateReportText(month, siteName, ga4Data, searchConsoleData, inquiryCount) {
  const topPage = ga4Data.pages.length ? ga4Data.pages[0].url : 'データなし';
  const topKeyword = searchConsoleData.queries.length ? searchConsoleData.queries[0].keyword : 'データなし';

  return [
    siteName + ' 月次レポート（' + month + '）',
    '',
    '■ アクセス状況',
    'ユーザー数: ' + ga4Data.users.toLocaleString(),
    '新規ユーザー数: ' + ga4Data.newUsers.toLocaleString(),
    'セッション数: ' + ga4Data.sessions.toLocaleString(),
    '最も閲覧されたページ: ' + topPage,
    '',
    '■ 検索流入状況',
    '検索表示回数: ' + searchConsoleData.impressions.toLocaleString(),
    '検索クリック数: ' + searchConsoleData.clicks.toLocaleString(),
    'CTR: ' + formatPercent_(searchConsoleData.ctr),
    '平均掲載順位: ' + round_(searchConsoleData.position, 1),
    '主な検索キーワード: ' + topKeyword,
    '',
    '■ 問い合わせ',
    '問い合わせ数: ' + inquiryCount.toLocaleString(),
    '',
    '■ 改善メモ',
    '・検索表示はあるがクリック率が低いキーワードは、タイトルや説明文の改善候補です。',
    '・閲覧数が多いページは、問い合わせ導線や関連ページへのリンクを確認してください。',
  ].join('\n');
}

function writeKeywordRows_(month, queries) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.queries);
  removeRowsByMonth_(sheet, month);

  if (!queries.length) {
    return;
  }

  const rows = queries.map(function(query) {
    return [month, query.keyword, query.impressions, query.clicks, query.ctr, query.position];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, QUERY_HEADERS.length).setValues(rows);
  formatQuerySheet_(sheet);
}

function writePageRows_(month, pages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.pages);
  removeRowsByMonth_(sheet, month);

  if (!pages.length) {
    return;
  }

  const rows = pages.map(function(page) {
    return [month, page.url, page.views, ''];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PAGE_HEADERS.length).setValues(rows);
  formatPageSheet_(sheet);
}

function writeReportSheet_(month, siteName, reportText, ga4Data, searchConsoleData, inquiryCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.report);
  sheet.clear();

  sheet.getRange('A1').setValue(siteName + ' 月次レポート');
  sheet.getRange('A2').setValue(month);
  sheet.getRange('A4').setValue('アクセス');
  sheet.getRange('A5:B7').setValues([
    ['ユーザー数', ga4Data.users],
    ['新規ユーザー数', ga4Data.newUsers],
    ['セッション数', ga4Data.sessions],
  ]);
  sheet.getRange('D4').setValue('検索');
  sheet.getRange('D5:E8').setValues([
    ['表示回数', searchConsoleData.impressions],
    ['クリック数', searchConsoleData.clicks],
    ['CTR', searchConsoleData.ctr],
    ['平均掲載順位', searchConsoleData.position],
  ]);
  sheet.getRange('A10').setValue('問い合わせ数');
  sheet.getRange('B10').setValue(inquiryCount);
  sheet.getRange('A12').setValue('レポート本文');
  sheet.getRange('A13:E25').merge().setValue(reportText);

  sheet.getRange('A1:E1').merge().setFontSize(18).setFontWeight('bold');
  sheet.getRange('A4:B4').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('D4:E4').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A12:E12').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A13:E25').setWrap(true).setVerticalAlignment('top');
  sheet.getRange('E7').setNumberFormat('0.00%');
  sheet.getRange('B5:B10').setNumberFormat('#,##0');
  sheet.getRange('E5:E6').setNumberFormat('#,##0');
  sheet.getRange('E8').setNumberFormat('0.0');
  sheet.setColumnWidths(1, 5, 140);
  sheet.setRowHeights(13, 13, 28);
}

function setupSettingsSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.settings) || ss.insertSheet(SHEET_NAMES.settings);

  if (sheet.getLastRow() === 0) {
    const rows = [
      ['項目', '値', '入力例・メモ'],
      [SETTINGS.siteName, '', '例: 京都はじまるサポート'],
      [SETTINGS.ga4PropertyId, '', '例: 123456789。GA4管理画面のプロパティID'],
      [SETTINGS.searchConsoleSiteUrl, '', '例: https://example.com/ または sc-domain:example.com'],
      [SETTINGS.inquirySheetName, '問い合わせ', '問い合わせ一覧が入っているシート名'],
      [SETTINGS.reportMonth, formatMonth_(new Date()), 'yyyy-MM形式。例: 2026-06'],
    ];
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  }

  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#d9ead3');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 260);
  sheet.setColumnWidths(3, 1, 420);
}

function setupSheet_(ss, sheetName, headers) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function setupReportSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.report) || ss.insertSheet(SHEET_NAMES.report);
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1').setValue('レポート出力用シート');
    sheet.getRange('A2').setValue('月次レポート作成後、PDF出力用の文章がここに作成されます。');
  }
  sheet.setColumnWidths(1, 5, 140);
}

function ensureRequiredSheets_(ss) {
  if (!ss.getSheetByName(SHEET_NAMES.settings)) {
    setupMonthlyReportSheets();
  }
  if (!ss.getSheetByName(SHEET_NAMES.summary)) setupSheet_(ss, SHEET_NAMES.summary, SUMMARY_HEADERS);
  if (!ss.getSheetByName(SHEET_NAMES.queries)) setupSheet_(ss, SHEET_NAMES.queries, QUERY_HEADERS);
  if (!ss.getSheetByName(SHEET_NAMES.pages)) setupSheet_(ss, SHEET_NAMES.pages, PAGE_HEADERS);
  if (!ss.getSheetByName(SHEET_NAMES.report)) setupReportSheet_(ss);
}

function getConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.settings);
  if (!sheet) {
    throw new Error('設定シートがありません。setupMonthlyReportSheets() を実行してください。');
  }

  const values = sheet.getDataRange().getValues();
  const config = {};

  values.slice(1).forEach(function(row) {
    const key = row[0];
    const value = row[1];
    if (key === SETTINGS.siteName) config.siteName = value || 'Webサイト';
    if (key === SETTINGS.ga4PropertyId) config.ga4PropertyId = String(value || '').trim();
    if (key === SETTINGS.searchConsoleSiteUrl) config.searchConsoleSiteUrl = String(value || '').trim();
    if (key === SETTINGS.inquirySheetName) config.inquirySheetName = String(value || '').trim();
    if (key === SETTINGS.reportMonth) config.reportMonth = normalizeMonth_(value);
  });

  return config;
}

function upsertRowByMonth_(sheet, month, row) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const months = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < months.length; i++) {
      if (String(months[i][0]) === month) {
        sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function removeRowsByMonth_(sheet, month) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const months = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = months.length - 1; i >= 0; i--) {
    if (String(months[i][0]) === month) {
      sheet.deleteRow(i + 2);
    }
  }
}

function findDateColumnIndex_(headers) {
  const candidates = ['日付', '日時', '受付日', '送信日時', 'created_at', 'created', 'date'];
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase();
    if (candidates.some(function(candidate) {
      return header.indexOf(candidate.toLowerCase()) !== -1;
    })) {
      return i;
    }
  }
  return -1;
}

function getMonthPeriod_(month) {
  const parts = month.split('-');
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  const start = new Date(year, monthIndex, 1);
  const end = new Date(year, monthIndex + 1, 0);

  return {
    startDate: formatDate_(start),
    endDate: formatDate_(end),
  };
}

function normalizeMonth_(value) {
  if (value instanceof Date) {
    return formatMonth_(value);
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}$/.test(text)) {
    return text;
  }
  if (/^\d{4}\/\d{1,2}$/.test(text)) {
    const parts = text.split('/');
    return parts[0] + '-' + ('0' + parts[1]).slice(-2);
  }
  return text;
}

function normalizeDateValue_(value) {
  if (value instanceof Date) {
    return value;
  }
  if (!value) {
    return null;
  }
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function parseDate_(value) {
  const parts = value.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function formatMonth_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatPercent_(value) {
  return round_(value * 100, 2) + '%';
}

function round_(value, digits) {
  const base = Math.pow(10, digits);
  return Math.round(Number(value || 0) * base) / base;
}

function formatSummarySheet_(sheet) {
  sheet.getRange('G:G').setNumberFormat('0.00%');
  sheet.getRange('H:H').setNumberFormat('0.0');
  sheet.getRange('B:F').setNumberFormat('#,##0');
  sheet.getRange('I:I').setNumberFormat('#,##0');
  sheet.getRange('J:K').setWrap(true);
  sheet.autoResizeColumns(1, SUMMARY_HEADERS.length);
}

function formatQuerySheet_(sheet) {
  sheet.getRange('E:E').setNumberFormat('0.00%');
  sheet.getRange('F:F').setNumberFormat('0.0');
  sheet.getRange('C:D').setNumberFormat('#,##0');
  sheet.autoResizeColumns(1, QUERY_HEADERS.length);
}

function formatPageSheet_(sheet) {
  sheet.getRange('C:C').setNumberFormat('#,##0');
  sheet.getRange('B:D').setWrap(true);
  sheet.autoResizeColumns(1, PAGE_HEADERS.length);
}
