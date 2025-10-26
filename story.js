/**
 * 7分割ストーリー機能
 */

/* ===== まとめて実行（STEP1+2） ===== */
function generateFortuneProStoryAndRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:手法）
  const theme  = String(sheet.getRange('A2').getValue() || '').trim();
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme)  { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください。'); return; }
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // 既存出力クリア（D5:E34, F5:M以降、ログは残す）
  sheet.getRange('D5:E34').clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) sheet.getRange(5, 6, Math.max(1, lastRow - 4), 8).clearContent(); // F5〜M

  // STEP1実行
  const storyText = executeStep1(sheet, apiKey, theme, method);

  // STEP2実行
  const postsCount = executeStep2(sheet, apiKey, method, storyText);

  SpreadsheetApp.getUi().alert(
    `完了：D5:E34に設計、F5:M${postsCount + 4} に ${postsCount} 本のストーリー＋IGキャプションを出力しました。`
  );
}

/* ===== STEP1のみ実行 ===== */
function generateStep1Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:手法）
  const theme  = String(sheet.getRange('A2').getValue() || '').trim();
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme)  { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください。'); return; }
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // STEP1出力エリアのみクリア
  sheet.getRange('D5:E34').clearContent();

  // STEP1実行
  executeStep1(sheet, apiKey, theme, method);

  SpreadsheetApp.getUi().alert('STEP1完了：D5:E34 にストーリー設計を出力しました。');
}

/* ===== STEP2のみ実行 ===== */
function generateStep2Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A3:手法）
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // STEP1の出力を取得
  const storyText = String(sheet.getRange('D5').getValue() || '').trim();
  if (!storyText) {
    SpreadsheetApp.getUi().alert('先にSTEP1を実行してください。D5:E34 にストーリー設計が必要です。');
    return;
  }

  // STEP2出力エリアのみクリア
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) sheet.getRange(5, 6, Math.max(1, lastRow - 4), 8).clearContent();

  // STEP2実行
  const postsCount = executeStep2(sheet, apiKey, method, storyText);

  SpreadsheetApp.getUi().alert(
    `STEP2完了：F5:M${postsCount + 4} に ${postsCount} 本のストーリー＋IGキャプションを出力しました。`
  );
}

/* ===== STEP1実行（共通処理） ===== */
function executeStep1(sheet, apiKey, theme, method) {
  // B5:B34からプロンプト取得、空なら初期化
  let promptStory = String(sheet.getRange('B5').getValue() || '').trim();
  if (!promptStory) {
    promptStory = getStoryDesignPrompt(theme, method);
    sheet.getRange('B5').setValue(promptStory);
  } else {
    // テンプレート変数を置換
    promptStory = promptStory.replace(/\{\{theme\}\}/g, theme).replace(/\{\{method\}\}/g, method);
  }

  const startTime = new Date();
  const storyText = callGemini(apiKey, promptStory);
  const endTime = new Date();

  // D5:E34に出力
  sheet.getRange('D5').setValue(storyText);

  // ログ出力
  addLog(sheet, 'STEP1: ストーリー設計', promptStory, storyText, startTime, endTime);

  return storyText;
}

/* ===== STEP2実行（共通処理） ===== */
function executeStep2(sheet, apiKey, method, storyText) {
  // C5:C34からプロンプト取得、空なら初期化
  let promptRows = String(sheet.getRange('C5').getValue() || '').trim();
  if (!promptRows) {
    promptRows = getRowsGenerationPrompt(method, storyText);
    sheet.getRange('C5').setValue(promptRows);
  } else {
    // テンプレート変数を置換
    promptRows = promptRows
      .replace(/\{\{method\}\}/g, method)
      .replace(/\{\{storyText\}\}/g, storyText);
  }

  const startTime = new Date();
  const rowsJson = callGemini(apiKey, promptRows);
  const endTime = new Date();

  const posts = parsePostsObjectsWithCaption(rowsJson);
  if (!posts || posts.length === 0) {
    SpreadsheetApp.getUi().alert('投稿生成に失敗しました。');
    return 0;
  }

  // F5〜M = 8列（title, l1a, l1b, l2a, l2b, l3a, l3b, ig_caption）
  const values = posts.map(p => [
    (p.title || '').trim(),
    (p.l1a || '').trim(),
    (p.l1b || '').trim(),
    (p.l2a || '').trim(),
    (p.l2b || '').trim(),
    (p.l3a || '').trim(),
    (p.l3b || '').trim(),
    (p.ig_caption || '').trim()
  ]);
  sheet.getRange(5, 6, values.length, 8).setValues(values);

  // ログ出力
  addLog(sheet, 'STEP2: 7分割生成', promptRows, rowsJson, startTime, endTime);

  return values.length;
}

/* ===== シート初期化（ヘッダー＋プロンプト配置） ===== */
function initializeSheet() {
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
  sheet.getRange('F1').setValue('タイトル');
  sheet.getRange('G1').setValue('L1A');
  sheet.getRange('H1').setValue('L1B');
  sheet.getRange('I1').setValue('L2A');
  sheet.getRange('J1').setValue('L2B');
  sheet.getRange('K1').setValue('L3A');
  sheet.getRange('L1').setValue('L3B');
  sheet.getRange('M1').setValue('IGキャプション');

  // ログヘッダー（35行目）
  sheet.getRange('N35').setValue('📊 実行ログ');
  sheet.getRange('O35').setValue('リクエスト');
  sheet.getRange('P35').setValue('レスポンス');

  // 入力エリア（2-3行目）
  sheet.getRange('A2').setValue('テーマを入力');
  sheet.getRange('A3').setValue('占い手法を入力');

  // サブヘッダー（4行目）
  sheet.getRange('B4').setValue('▼ STEP1プロンプト本文');
  sheet.getRange('C4').setValue('▼ STEP2プロンプト本文');
  sheet.getRange('D4').setValue('▼ STEP1出力本文');
  sheet.getRange('F4').setValue('▼ STEP2出力');

  // デフォルトプロンプトを配置（5行目から縦30行結合）
  const defaultPrompt1 = getStoryDesignPrompt('{{theme}}', '{{method}}');
  sheet.getRange('B5').setValue(defaultPrompt1);
  sheet.getRange('B5:B34').merge();

  const defaultPrompt2 = getRowsGenerationPrompt('{{method}}', '{{storyText}}');
  sheet.getRange('C5').setValue(defaultPrompt2);
  sheet.getRange('C5:C34').merge();

  // STEP1出力エリアを結合（D5:E34）
  sheet.getRange('D5:E34').merge();

  // フォーマット適用
  formatSheet(sheet);

  SpreadsheetApp.getUi().alert('シートを初期化しました。A2にテーマ、A3に占い手法を入力してください。');
}

/* ===== セルフォーマッティング ===== */
function formatSheet(sheet) {
  // ヘッダー行（1行目）をボールド＋背景色
  const headerRange = sheet.getRange('A1:P1');
  headerRange.setFontWeight('bold')
             .setBackground('#4a86e8')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  // サブヘッダー行（4行目）をボールド＋背景色
  const subHeaderRange = sheet.getRange('A4:P4');
  subHeaderRange.setFontWeight('bold')
                .setBackground('#6d9eeb')
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

  // STEP2出力（F5:M以降）
  sheet.getRange('F5:M').setBackground('#f4cccc').setWrap(true);

  // ログヘッダー行（35行目）
  sheet.getRange('N35:P35').setFontWeight('bold')
                           .setBackground('#c27ba0')
                           .setFontColor('#ffffff')
                           .setHorizontalAlignment('center');

  // ログエリア（36行目以降）
  sheet.getRange('N36:P').setBackground('#ead1dc').setWrap(true);

  // 列幅調整
  sheet.setColumnWidth(1, 150);  // A列（入力）
  sheet.setColumnWidth(2, 450);  // B列（STEP1プロンプト）
  sheet.setColumnWidth(3, 450);  // C列（STEP2プロンプト）
  sheet.setColumnWidth(4, 450);  // D列（STEP1出力）
  sheet.setColumnWidth(5, 50);   // E列（結合用）
  sheet.setColumnWidths(6, 8, 130); // F-M列（STEP2出力）
  sheet.setColumnWidth(14, 150); // N列（タイムスタンプ）
  sheet.setColumnWidth(15, 350); // O列（リクエスト）
  sheet.setColumnWidth(16, 350); // P列（レスポンス）

  // 行の高さ調整
  sheet.setRowHeight(1, 40);  // ヘッダー行
  sheet.setRowHeight(4, 35);  // サブヘッダー行
  sheet.setRowHeights(5, 30, 60); // 5-34行目（結合セル用）
  sheet.setRowHeight(35, 35); // ログヘッダー行
}
