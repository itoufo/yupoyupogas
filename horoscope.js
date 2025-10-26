/**
 * 今日の星座占い機能
 */

/* ===== まとめて実行 ===== */
function generateTodayHoroscope() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // A2セルから日付を取得（空なら今日の日付）
  let date = String(sheet.getRange('A2').getValue() || '').trim();
  if (!date) {
    const today = new Date();
    date = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
    sheet.getRange('A2').setValue(date);
  }

  // 既存出力クリア（D5以降）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 4, Math.max(1, lastRow - 4), 20).clearContent();
  }

  // 実行
  executeHoroscope(sheet, apiKey, date);

  SpreadsheetApp.getUi().alert('完了：D5以降に今日の星座占いを出力しました。');
}

/* ===== 実行処理（共通） ===== */
function executeHoroscope(sheet, apiKey, date) {
  // B5からプロンプト取得、空なら初期化
  let prompt = String(sheet.getRange('B5').getValue() || '').trim();
  if (!prompt) {
    prompt = getTodayHoroscopePrompt('{{date}}');
    sheet.getRange('B5').setValue(prompt);
  }

  // テンプレート変数を置換
  prompt = prompt.replace(/\{\{date\}\}/g, date);

  const startTime = new Date();
  const response = callGemini(apiKey, prompt);
  const endTime = new Date();

  // JSONパース
  const parsedData = parseHoroscopeData(response);
  if (!parsedData) {
    throw new Error('JSONのパースに失敗しました。GPTの応答を確認してください。');
  }

  // シートに出力
  outputHoroscopeToSheet(sheet, parsedData, date);

  // ログ出力
  addLog(sheet, '今日の星座占い', prompt, response, startTime, endTime);
}

/* ===== シート初期化 ===== */
function initializeHoroscopeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // シートをクリア（1行目以外）
  const lastRow = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
  }

  // A2セルに今日の日付を設定
  const today = new Date();
  const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
  sheet.getRange('A2').setValue(dateStr);

  // ヘッダー設定
  sheet.getRange('A1').setValue('日付');
  sheet.getRange('A2').setValue(dateStr);
  sheet.getRange('B1').setValue('プロンプト（編集可能）');
  sheet.getRange('D1').setValue('出力エリア →');

  // 列幅設定
  sheet.setColumnWidth(1, 150);  // A列: 日付
  sheet.setColumnWidth(2, 500);  // B列: プロンプト
  sheet.setColumnWidth(3, 30);   // C列: 空白
  sheet.setColumnWidth(4, 150);  // D列: 出力開始

  // E列〜P列: ランキング表示用（12列分）
  for (let i = 5; i <= 16; i++) {
    sheet.setColumnWidth(i, 120);
  }

  SpreadsheetApp.getUi().alert('シートを初期化しました！\nA2に日付を入力して「今日の星座占いを生成」を実行してください。');
}

/* ===== シートへの出力 ===== */
function outputHoroscopeToSheet(sheet, data, date) {
  // 出力開始行
  let currentRow = 5;

  // タイトル
  sheet.getRange(currentRow, 4, 1, 4)
    .merge()
    .setValue(`${date} 星座占い`)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e6f2ff');
  currentRow += 2;

  // 今日の概要
  sheet.getRange(currentRow, 4)
    .setValue('今日はどんな日？')
    .setFontWeight('bold')
    .setBackground('#fff2cc')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(currentRow, 5, 1, 4)
    .merge()
    .setValue(data.today_overview.description)
    .setWrap(true)
    .setVerticalAlignment('top');
  currentRow++;

  // パワースポット
  sheet.getRange(currentRow, 4)
    .setValue('今日のパワースポット')
    .setFontWeight('bold')
    .setBackground('#fff2cc')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(currentRow, 5, 1, 4)
    .merge()
    .setValue(data.today_overview.power_spot)
    .setWrap(true)
    .setVerticalAlignment('middle');
  currentRow += 3;

  // ランキング表示（横並び12列）
  sheet.getRange(currentRow, 5)
    .setValue('星座占いランキング')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  // ヘッダー行（順位）
  const headerData = [];
  for (let i = 1; i <= 12; i++) {
    headerData.push(`${i}位`);
  }
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([headerData])
    .setFontWeight('bold')
    .setBackground('#ffd966')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // 星座名行
  const zodiacData = data.rankings.map(item => item.zodiac);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([zodiacData])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(11);
  currentRow++;

  // ラッキーカラー行
  sheet.getRange(currentRow, 4)
    .setValue('ラッキーカラー')
    .setFontWeight('bold')
    .setBackground('#e6f2ff')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const colorData = data.rankings.map(item => item.lucky_color);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([colorData])
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // ラッキーフード行
  sheet.getRange(currentRow, 4)
    .setValue('ラッキーフード')
    .setFontWeight('bold')
    .setBackground('#e6f2ff')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const foodData = data.rankings.map(item => item.lucky_food);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([foodData])
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // ラッキーアクション行
  sheet.getRange(currentRow, 4)
    .setValue('ラッキーアクション')
    .setFontWeight('bold')
    .setBackground('#e6f2ff')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const actionData = data.rankings.map(item => item.lucky_action);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([actionData])
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow += 2;

  // 恋愛運（短文）
  sheet.getRange(currentRow, 4)
    .setValue('恋愛運（短）')
    .setFontWeight('bold')
    .setBackground('#ffe6f0')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const loveShortData = data.rankings.map(item => item.love_short);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([loveShortData])
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(9);
  currentRow++;

  // 恋愛運（長文）
  sheet.getRange(currentRow, 4)
    .setValue('恋愛運（長）')
    .setFontWeight('bold')
    .setBackground('#ffe6f0')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const loveLongData = data.rankings.map(item => item.love_long);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([loveLongData])
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(9);
  currentRow += 2;

  // 総合運（短文）
  sheet.getRange(currentRow, 4)
    .setValue('総合運（短）')
    .setFontWeight('bold')
    .setBackground('#e6ffe6')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const overallShortData = data.rankings.map(item => item.overall_short);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([overallShortData])
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(9);
  currentRow++;

  // 総合運（長文）
  sheet.getRange(currentRow, 4)
    .setValue('総合運（長）')
    .setFontWeight('bold')
    .setBackground('#e6ffe6')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const overallLongData = data.rankings.map(item => item.overall_long);
  sheet.getRange(currentRow, 5, 1, 12)
    .setValues([overallLongData])
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(9);

  // 行の高さ調整
  sheet.setRowHeight(currentRow - 1, 50);  // 恋愛運（長）
  sheet.setRowHeight(currentRow, 50);      // 総合運（長）
}
