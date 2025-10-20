/**
 * ランキング30機能
 */

/* ===== まとめて実行（STEP1+2） ===== */
function generateRankingContent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:星座or誕生月）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  const type = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (type !== '星座' && type !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }

  // 既存出力クリア（D5:E34, F5:G以降）
  sheet.getRange('D5:E34').clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    // F列（6）からG列（7）までクリア
    sheet.getRange(5, 6, Math.max(1, lastRow - 4), 2).clearContent();
  }

  // STEP1実行
  const designText = executeRankingStep1(sheet, apiKey, theme, type);

  // STEP2実行
  executeRankingStep2(sheet, apiKey, theme, type, designText);

  SpreadsheetApp.getUi().alert('完了：D5:E34にランキング設計、F5以降にランキング30位＋キャプションを出力しました。');
}

/* ===== STEP1のみ実行 ===== */
function generateRankingStep1Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:星座or誕生月）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  const type = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (type !== '星座' && type !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }

  // STEP1出力エリアのみクリア
  sheet.getRange('D5:E34').clearContent();

  // STEP1実行
  executeRankingStep1(sheet, apiKey, theme, type);

  SpreadsheetApp.getUi().alert('STEP1完了：D5:E34 にランキング設計を出力しました。');
}

/* ===== STEP2のみ実行 ===== */
function generateRankingStep2Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:星座or誕生月）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  const type = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (type !== '星座' && type !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }

  // STEP1の出力を取得
  const designText = String(sheet.getRange('D5').getValue() || '').trim();
  if (!designText) {
    SpreadsheetApp.getUi().alert('先にSTEP1を実行してください。D5:E34 にランキング設計が必要です。');
    return;
  }

  // STEP2出力エリアのみクリア（F5:G以降）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    // F列（6）からG列（7）までクリア
    sheet.getRange(5, 6, Math.max(1, lastRow - 4), 2).clearContent();
  }

  // STEP2実行
  executeRankingStep2(sheet, apiKey, theme, type, designText);

  SpreadsheetApp.getUi().alert('STEP2完了：F5以降にランキング30位＋キャプションを出力しました。');
}

/* ===== STEP1実行（共通処理） ===== */
function executeRankingStep1(sheet, apiKey, theme, type) {
  // B5:B34からプロンプト取得、空なら初期化
  let promptDesign = String(sheet.getRange('B5').getValue() || '').trim();
  if (!promptDesign) {
    promptDesign = getRankingDesignPrompt(theme, type);
    sheet.getRange('B5').setValue(promptDesign);
  } else {
    // テンプレート変数を置換
    promptDesign = promptDesign
      .replace(/\{\{theme\}\}/g, theme)
      .replace(/\{\{type\}\}/g, type);
  }

  const startTime = new Date();
  const designText = callGPT5(apiKey, promptDesign);
  const endTime = new Date();

  // D5:E34に出力
  sheet.getRange('D5').setValue(designText);

  // ログ出力
  addLog(sheet, 'ランキングSTEP1: ランキング設計', promptDesign, designText, startTime, endTime);

  return designText;
}

/* ===== STEP2実行（共通処理） ===== */
function executeRankingStep2(sheet, apiKey, theme, type, designText) {
  // C5:C34からプロンプト取得、空なら初期化
  let promptRanking = String(sheet.getRange('C5').getValue() || '').trim();
  if (!promptRanking) {
    promptRanking = getRankingContentsPrompt(theme, type, designText);
    sheet.getRange('C5').setValue(promptRanking);
  } else {
    // テンプレート変数を置換
    promptRanking = promptRanking
      .replace(/\{\{theme\}\}/g, theme)
      .replace(/\{\{type\}\}/g, type)
      .replace(/\{\{designText\}\}/g, designText);
  }

  const startTime = new Date();
  const rankingJson = callGPT5(apiKey, promptRanking);
  const endTime = new Date();

  const parsedData = parseRankingContents(rankingJson);
  if (!parsedData) {
    SpreadsheetApp.getUi().alert('ランキング生成に失敗しました。');
    return;
  }

  // F5以降に出力
  let currentRow = 5;

  // ヘッダー行
  sheet.getRange(currentRow, 6).setValue('順位')
       .setFontWeight('bold')
       .setBackground('#ffd966')
       .setHorizontalAlignment('center');
  sheet.getRange(currentRow, 7).setValue('組み合わせ＋説明')
       .setFontWeight('bold')
       .setBackground('#ffd966')
       .setHorizontalAlignment('center');
  currentRow++;

  // ランキング内容
  parsedData.rankings.forEach((item) => {
    sheet.getRange(currentRow, 6).setValue(item.rank)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.getRange(currentRow, 7).setValue(`${item.combination}\n${item.description}`)
         .setWrap(true)
         .setVerticalAlignment('top');
    currentRow++;
  });

  // 空行
  currentRow++;

  // キャプション出力
  sheet.getRange(currentRow, 6, 1, 2).merge()
       .setValue('【Instagramキャプション】')
       .setFontWeight('bold')
       .setBackground('#b6d7a8')
       .setHorizontalAlignment('center');
  currentRow++;

  sheet.getRange(currentRow, 6, 1, 2).merge()
       .setValue(parsedData.instagram_caption)
       .setWrap(true)
       .setVerticalAlignment('top');

  // ログ出力
  addLog(sheet, 'ランキングSTEP2: ランキング30生成', promptRanking, rankingJson, startTime, endTime);
}

/* ===== ランキングシート初期化 ===== */
function initializeRankingSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 既存の結合を解除
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxCols).breakApart();

  // 既存内容を全てクリア（2行目以降）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  }

  // ヘッダー設定（1行目）
  sheet.getRange('A1').setValue('📝 入力');
  sheet.getRange('B1').setValue('📋 STEP1プロンプト');
  sheet.getRange('C1').setValue('📋 STEP2プロンプト');
  sheet.getRange('D1').setValue('✨ STEP1出力');
  sheet.getRange('F1:G1').merge().setValue('🏆 STEP2出力');

  // ログヘッダー（35行目）
  sheet.getRange('H35').setValue('📊 実行ログ');
  sheet.getRange('I35').setValue('リクエスト');
  sheet.getRange('J35').setValue('レスポンス');

  // 入力エリア（2-3行目）
  sheet.getRange('A2').setValue('ランキングテーマを入力（例：2025年の恋愛運）');
  sheet.getRange('A3').setValue('星座 or 誕生月を選択');

  // A3にドロップダウンを設定
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['星座', '誕生月'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A3').setDataValidation(rule);

  // サブヘッダー（4行目）
  sheet.getRange('B4').setValue('▼ STEP1プロンプト本文');
  sheet.getRange('C4').setValue('▼ STEP2プロンプト本文');
  sheet.getRange('D4').setValue('▼ STEP1出力本文');
  sheet.getRange('F4:G4').merge().setValue('▼ STEP2出力（ランキング30位）');

  // デフォルトプロンプトを配置（5行目から縦30行結合）
  const defaultPrompt1 = getRankingDesignPrompt('{{theme}}', '{{type}}');
  sheet.getRange('B5').setValue(defaultPrompt1);
  sheet.getRange('B5:B34').merge();

  const defaultPrompt2 = getRankingContentsPrompt('{{theme}}', '{{type}}', '{{designText}}');
  sheet.getRange('C5').setValue(defaultPrompt2);
  sheet.getRange('C5:C34').merge();

  // STEP1出力エリアを結合（D5:E34）
  sheet.getRange('D5:E34').merge();

  // フォーマット適用
  formatRankingSheet(sheet);

  SpreadsheetApp.getUi().alert('ランキングシートを初期化しました。A2にテーマ、A3に「星座」または「誕生月」を選択してください。');
}

/* ===== ランキングシートフォーマッティング ===== */
function formatRankingSheet(sheet) {
  // ヘッダー行（1行目）をボールド＋背景色
  const headerRange = sheet.getRange('A1:J1');
  headerRange.setFontWeight('bold')
             .setBackground('#e69138')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  // サブヘッダー行（4行目）をボールド＋背景色
  const subHeaderRange = sheet.getRange('A4:J4');
  subHeaderRange.setFontWeight('bold')
                .setBackground('#f6b26b')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');

  // 入力エリア（A2:A3）
  sheet.getRange('A2:A3').setBackground('#fff2cc');

  // プロンプトエリア（B5:B34, C5:C34）
  sheet.getRange('B5:B34').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');
  sheet.getRange('C5:C34').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP1出力（D5:E34）
  sheet.getRange('D5:E34').setBackground('#cfe2f3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP2出力（F5:G以降）
  sheet.getRange('F:G').setBackground('#fce5cd').setWrap(true);

  // ログヘッダー行（35行目）
  sheet.getRange('H35:J35').setFontWeight('bold')
                           .setBackground('#c27ba0')
                           .setFontColor('#ffffff')
                           .setHorizontalAlignment('center');

  // ログエリア（36行目以降）
  sheet.getRange('H36:J').setBackground('#ead1dc').setWrap(true);

  // 列幅調整
  sheet.setColumnWidth(1, 200);  // A列（入力）
  sheet.setColumnWidth(2, 450);  // B列（STEP1プロンプト）
  sheet.setColumnWidth(3, 450);  // C列（STEP2プロンプト）
  sheet.setColumnWidth(4, 350);  // D列（STEP1出力）
  sheet.setColumnWidth(5, 50);   // E列（結合用）
  sheet.setColumnWidth(6, 80);   // F列（順位）
  sheet.setColumnWidth(7, 300);  // G列（組み合わせ＋説明）
  sheet.setColumnWidth(8, 150);  // H列（タイムスタンプ）
  sheet.setColumnWidth(9, 350);  // I列（リクエスト）
  sheet.setColumnWidth(10, 350); // J列（レスポンス）

  // 行の高さ調整
  sheet.setRowHeight(1, 40);  // ヘッダー行
  sheet.setRowHeight(4, 35);  // サブヘッダー行
  sheet.setRowHeights(5, 30, 60); // 5-34行目（結合セル用）
  sheet.setRowHeight(35, 35); // ログヘッダー行
}
