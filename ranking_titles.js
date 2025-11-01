/**
 * ランキング名量産機能
 * 恋愛関係のランキング名を50個生成
 */

/* ===== ランキング名を50個生成 ===== */
function generateRankingTitles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // プロンプト生成
  const prompt = getRankingTitlesPrompt();

  const startTime = new Date();
  const response = callGemini(apiKey, prompt);
  const endTime = new Date();

  // JSONパース
  const parsedData = parseRankingTitles(response);
  if (!parsedData) {
    throw new Error('JSONのパースに失敗しました。Geminiの応答を確認してください。');
  }

  // シートに出力（追記）
  outputRankingTitlesToSheet(sheet, parsedData);

  // ログ出力（G列 = 7列目、36行目以降）
  addLogForRankingTitles(sheet, 'ランキング名生成', prompt, response, startTime, endTime);

  SpreadsheetApp.getUi().alert(`完了：${parsedData.titles.length}個のランキング名を追記しました。`);
}

/* ===== ランキング名を50個生成（ネガティブ→ポジティブ寄り添い型） ===== */
function generateRankingTitlesNegativeToPositive() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // プロンプト生成（ネガティブ→ポジティブ版）
  const prompt = getRankingTitlesNegativeToPositivePrompt();

  const startTime = new Date();
  const response = callGemini(apiKey, prompt);
  const endTime = new Date();

  // JSONパース
  const parsedData = parseRankingTitles(response);
  if (!parsedData) {
    throw new Error('JSONのパースに失敗しました。Geminiの応答を確認してください。');
  }

  // シートに出力（追記）
  outputRankingTitlesToSheet(sheet, parsedData);

  // ログ出力（G列 = 7列目、36行目以降）
  addLogForRankingTitles(sheet, 'ランキング名生成（ネガポジ型）', prompt, response, startTime, endTime);

  SpreadsheetApp.getUi().alert(`完了：${parsedData.titles.length}個のランキング名（ネガティブ→ポジティブ型）を追記しました。`);
}

/* ===== シートへの出力（追記） ===== */
function outputRankingTitlesToSheet(sheet, data) {
  // D列で最後の空でない行を探す（追記位置を特定）
  const maxRows = sheet.getMaxRows();
  let appendRow = 5; // デフォルトは5行目から

  for (let i = 5; i <= maxRows; i++) {
    const cellValue = sheet.getRange(i, 4).getValue();
    if (!cellValue || cellValue === '') {
      appendRow = i;
      break;
    }
  }

  // 50個のランキング名を追記
  data.titles.forEach((item, index) => {
    const currentRow = appendRow + index;

    // D列: 番号
    sheet.getRange(currentRow, 4).setValue(item.number).setHorizontalAlignment('center');

    // E列: ランキング名
    sheet.getRange(currentRow, 5).setValue(item.title).setWrap(true).setVerticalAlignment('middle');

    // F列: 説明
    sheet.getRange(currentRow, 6).setValue(item.description).setWrap(true).setVerticalAlignment('top');

    // 行の高さを調整
    sheet.setRowHeight(currentRow, 60);
  });
}

/* ===== ランキング名専用ログ出力（G列 = 7列目） ===== */
function addLogForRankingTitles(sheet, stepName, request, response, startTime, endTime) {
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestSummary = `[${stepName}]\n実行時間: ${duration}秒\n\nプロンプト:\n${request.substring(0, 500)}${request.length > 500 ? '...' : ''}`;
  const responseSummary = `レスポンス:\n${response.substring(0, 500)}${response.length > 500 ? '...' : ''}`;

  const logColumn = 7;  // G列
  const logStartRow = 36;
  let logRow = logStartRow;
  const maxRows = sheet.getMaxRows();

  // ログ列で最後の空でない行を探す
  for (let i = logStartRow; i <= maxRows; i++) {
    const cellValue = sheet.getRange(i, logColumn).getValue();
    if (!cellValue || cellValue === '') {
      logRow = i;
      break;
    }
  }

  sheet.getRange(logRow, logColumn, 1, 3).setValues([[timestamp, requestSummary, responseSummary]]);
}

/* ===== シート初期化 ===== */
function initializeRankingTitlesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // シートをクリア（1行目以外）
  const lastRow = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
  }

  // ヘッダー設定
  sheet.getRange('A1').setValue('ランキング名量産（恋愛系）');
  sheet.getRange('A2').setValue('※50個のランキング名を生成します');
  sheet.getRange('D1').setValue('出力エリア →');

  // 列幅設定
  sheet.setColumnWidth(1, 200);  // A列
  sheet.setColumnWidth(2, 30);   // B列: 空白
  sheet.setColumnWidth(3, 30);   // C列: 空白
  sheet.setColumnWidth(4, 60);   // D列: 番号
  sheet.setColumnWidth(5, 400);  // E列: ランキング名
  sheet.setColumnWidth(6, 500);  // F列: 説明
  sheet.setColumnWidth(7, 150);  // G列: ログ（タイムスタンプ）
  sheet.setColumnWidth(8, 350);  // H列: ログ（リクエスト）
  sheet.setColumnWidth(9, 350);  // I列: ログ（レスポンス）

  // ヘッダー行（4行目）
  sheet.getRange('D4').setValue('番号').setFontWeight('bold').setBackground('#e6d7ff').setHorizontalAlignment('center');
  sheet.getRange('E4').setValue('ランキング名').setFontWeight('bold').setBackground('#e6d7ff').setHorizontalAlignment('center');
  sheet.getRange('F4').setValue('説明').setFontWeight('bold').setBackground('#e6d7ff').setHorizontalAlignment('center');

  // ログヘッダー（35行目）
  sheet.getRange('G35').setValue('📊 実行ログ').setFontWeight('bold').setBackground('#c27ba0').setFontColor('#ffffff');
  sheet.getRange('H35').setValue('リクエスト').setFontWeight('bold').setBackground('#c27ba0').setFontColor('#ffffff');
  sheet.getRange('I35').setValue('レスポンス').setFontWeight('bold').setBackground('#c27ba0').setFontColor('#ffffff');

  SpreadsheetApp.getUi().alert('シートを初期化しました！\n「ランキング名を生成」を実行してください。');
}
