function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔮 GPT占い投稿生成')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📖 7分割ストーリー')
      .addItem('STEP1+2：まとめて実行（ストーリー設計→7分割）', 'generateFortuneProStoryAndRows')
      .addSeparator()
      .addItem('STEP1のみ：ストーリー設計（GPT-5）', 'generateStep1Only')
      .addItem('STEP2のみ：7分割＋IGキャプ生成（GPT-5）', 'generateStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initializeSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⭐ 12星座別コンテンツ')
      .addItem('STEP1+2：まとめて実行（サブテーマ生成→12星座）', 'generate12ZodiacContent')
      .addSeparator()
      .addItem('STEP1のみ：サブテーマ生成（GPT-5）', 'generate12ZodiacStep1Only')
      .addItem('STEP2のみ：12星座＋IGキャプ生成（GPT-5）', 'generate12ZodiacStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initialize12ZodiacSheet'))
    .addToUi();
}

function generateFortuneProStoryAndRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:手法）
  const theme  = String(sheet.getRange('A2').getValue() || '').trim();
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme)  { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください。'); return; }
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // 既存出力クリア（D5:E14, F5:M以降、ログは残す）
  sheet.getRange('D5:E14').clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) sheet.getRange(5, 6, Math.max(1, lastRow - 4), 8).clearContent(); // F5〜M

  // STEP1実行
  const storyText = executeStep1(sheet, apiKey, theme, method);

  // STEP2実行
  const postsCount = executeStep2(sheet, apiKey, method, storyText);

  SpreadsheetApp.getUi().alert(
    `完了：D5:E14に設計、F5:M${postsCount + 4} に ${postsCount} 本のストーリー＋IGキャプションを出力しました。`
  );
}

/* ===== STEP1のみ実行 ===== */
function generateStep1Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:手法）
  const theme  = String(sheet.getRange('A2').getValue() || '').trim();
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!theme)  { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください。'); return; }
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // STEP1出力エリアのみクリア
  sheet.getRange('D5:E14').clearContent();

  // STEP1実行
  executeStep1(sheet, apiKey, theme, method);

  SpreadsheetApp.getUi().alert('STEP1完了：D5:E14 にストーリー設計を出力しました。');
}

/* ===== STEP2のみ実行 ===== */
function generateStep2Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A3:手法）
  const method = String(sheet.getRange('A3').getValue() || '').trim();
  if (!method) { SpreadsheetApp.getUi().alert('A3 に占い手法を入力してください。'); return; }

  // STEP1の出力を取得
  const storyText = String(sheet.getRange('D5').getValue() || '').trim();
  if (!storyText) {
    SpreadsheetApp.getUi().alert('先にSTEP1を実行してください。D5:E14 にストーリー設計が必要です。');
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
  // B5:B14からプロンプト取得、空なら初期化
  let promptStory = String(sheet.getRange('B5').getValue() || '').trim();
  if (!promptStory) {
    promptStory = getStoryDesignPrompt(theme, method);
    sheet.getRange('B5').setValue(promptStory);
  } else {
    // テンプレート変数を置換
    promptStory = promptStory.replace(/\{\{theme\}\}/g, theme).replace(/\{\{method\}\}/g, method);
  }

  const startTime = new Date();
  const storyText = callGPT5(apiKey, promptStory);
  const endTime = new Date();

  // D5:E14に出力
  sheet.getRange('D5').setValue(storyText);

  // ログ出力
  addLog(sheet, 'STEP1: ストーリー設計', promptStory, storyText, startTime, endTime);

  return storyText;
}

/* ===== STEP2実行（共通処理） ===== */
function executeStep2(sheet, apiKey, method, storyText) {
  // C5:C14からプロンプト取得、空なら初期化
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
  const rowsJson = callGPT5(apiKey, promptRows);
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

/* ===== GPT-5呼び出し（温度指定なし） ===== */
function callGPT5(apiKey, prompt) {
  const payload = { model: 'gpt-5', messages: [{ role: 'user', content: prompt }] };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload)
  };
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const json = JSON.parse(res.getContentText());
  return json.choices[0].message.content.trim();
}

/* ===== JSONパース ===== */
function parsePostsObjectsWithCaption(text) {
  let cleaned = text.replace(/^```json|^```|```$/gmi, '').trim();
  const first = cleaned.indexOf('{');
  const last  = cleaned.lastIndexOf('}');
  if (first >= 0 && last > first) cleaned = cleaned.slice(first, last + 1);
  try {
    const obj = JSON.parse(cleaned);
    if (!obj || !Array.isArray(obj.posts)) return null;
    return obj.posts.map(p => ({
      title:      String(p.title || ''),
      l1a:        String(p.l1a || ''),
      l1b:        String(p.l1b || ''),
      l2a:        String(p.l2a || ''),
      l2b:        String(p.l2b || ''),
      l3a:        String(p.l3a || ''),
      l3b:        String(p.l3b || ''),
      ig_caption: String(p.ig_caption || '')
    }));
  } catch {
    return null;
  }
}

/* ===== ログ出力（N列〜P列） ===== */
function addLog(sheet, stepName, request, response, startTime, endTime) {
  const logRow = sheet.getLastRow() + 1;
  const duration = ((endTime - startTime) / 1000).toFixed(2);

  const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestSummary = `[${stepName}]\n実行時間: ${duration}秒\n\nプロンプト:\n${request.substring(0, 500)}${request.length > 500 ? '...' : ''}`;
  const responseSummary = `レスポンス:\n${response.substring(0, 500)}${response.length > 500 ? '...' : ''}`;

  sheet.getRange(logRow, 14, 1, 3).setValues([[timestamp, requestSummary, responseSummary]]);
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
  sheet.getRange('N1').setValue('📊 実行ログ');
  sheet.getRange('O1').setValue('リクエスト');
  sheet.getRange('P1').setValue('レスポンス');

  // 入力エリア（2-3行目）
  sheet.getRange('A2').setValue('テーマを入力');
  sheet.getRange('A3').setValue('占い手法を入力');

  // サブヘッダー（4行目）
  sheet.getRange('B4').setValue('▼ STEP1プロンプト本文');
  sheet.getRange('C4').setValue('▼ STEP2プロンプト本文');
  sheet.getRange('D4').setValue('▼ STEP1出力本文');
  sheet.getRange('F4').setValue('▼ STEP2出力');

  // デフォルトプロンプトを配置（5行目から縦10行結合）
  const defaultPrompt1 = getStoryDesignPrompt('{{theme}}', '{{method}}');
  sheet.getRange('B5').setValue(defaultPrompt1);
  sheet.getRange('B5:B14').merge();

  const defaultPrompt2 = getRowsGenerationPrompt('{{method}}', '{{storyText}}');
  sheet.getRange('C5').setValue(defaultPrompt2);
  sheet.getRange('C5:C14').merge();

  // STEP1出力エリアを結合（D5:E14）
  sheet.getRange('D5:E14').merge();

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

  // プロンプトエリア（B5:B14, C5:C14）
  sheet.getRange('B5:B14').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');
  sheet.getRange('C5:C14').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP1出力（D5:E14）
  sheet.getRange('D5:E14').setBackground('#cfe2f3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP2出力（F5:M以降）
  sheet.getRange('F5:M').setBackground('#f4cccc').setWrap(true);

  // ログエリア（N列以降）
  sheet.getRange('N:P').setBackground('#ead1dc').setWrap(true);

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
  sheet.setRowHeights(5, 10, 60); // 5-14行目（結合セル用）
}

/* ========================================
 * 12星座別コンテンツ生成機能
 * ======================================== */

/* ===== まとめて実行（STEP1+2） ===== */
function generate12ZodiacContent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください（例：恋愛）'); return; }

  // 既存出力クリア（D5:D14, E5以降）
  sheet.getRange('D5:D14').clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    const lastCol = sheet.getLastColumn();
    if (lastCol >= 5) {
      sheet.getRange(5, 5, Math.max(1, lastRow - 4), lastCol - 4).clearContent();
    }
  }

  // STEP1実行
  const subThemes = execute12ZodiacStep1(sheet, apiKey, theme);

  // STEP2実行
  execute12ZodiacStep2(sheet, apiKey, theme, subThemes);

  SpreadsheetApp.getUi().alert('完了：D5:D14にサブテーマ、E5以降に12星座別コンテンツ＋キャプションを出力しました。');
}

/* ===== STEP1のみ実行 ===== */
function generate12ZodiacStep1Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください（例：恋愛）'); return; }

  // STEP1出力エリアのみクリア
  sheet.getRange('D5:D14').clearContent();

  // STEP1実行
  execute12ZodiacStep1(sheet, apiKey, theme);

  SpreadsheetApp.getUi().alert('STEP1完了：D5:D14 にサブテーマ一覧を出力しました。');
}

/* ===== STEP2のみ実行 ===== */
function generate12ZodiacStep2Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ）
  const theme = String(sheet.getRange('A2').getValue() || '').trim();
  if (!theme) { SpreadsheetApp.getUi().alert('A2 にテーマを入力してください（例：恋愛）'); return; }

  // STEP1の出力を取得
  const subThemes = String(sheet.getRange('D5').getValue() || '').trim();
  if (!subThemes) {
    SpreadsheetApp.getUi().alert('先にSTEP1を実行してください。D5:D14 にサブテーマ一覧が必要です。');
    return;
  }

  // STEP2出力エリアのみクリア
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    const lastCol = sheet.getLastColumn();
    if (lastCol >= 5) {
      sheet.getRange(5, 5, Math.max(1, lastRow - 4), lastCol - 4).clearContent();
    }
  }

  // STEP2実行
  execute12ZodiacStep2(sheet, apiKey, theme, subThemes);

  SpreadsheetApp.getUi().alert('STEP2完了：E5以降に12星座別コンテンツ＋キャプションを出力しました。');
}

/* ===== STEP1実行（共通処理） ===== */
function execute12ZodiacStep1(sheet, apiKey, theme) {
  // B5:B14からプロンプト取得、空なら初期化
  let promptThemes = String(sheet.getRange('B5').getValue() || '').trim();
  if (!promptThemes) {
    promptThemes = getZodiacThemesPrompt(theme);
    sheet.getRange('B5').setValue(promptThemes);
  } else {
    // テンプレート変数を置換
    promptThemes = promptThemes.replace(/\{\{theme\}\}/g, theme);
  }

  const startTime = new Date();
  const subThemes = callGPT5(apiKey, promptThemes);
  const endTime = new Date();

  // D5:D14に出力
  sheet.getRange('D5').setValue(subThemes);

  // ログ出力
  addLog(sheet, '12星座STEP1: サブテーマ生成', promptThemes, subThemes, startTime, endTime);

  return subThemes;
}

/* ===== STEP2実行（共通処理） ===== */
function execute12ZodiacStep2(sheet, apiKey, theme, subThemes) {
  // C5:C14からプロンプト取得、空なら初期化
  let promptContents = String(sheet.getRange('C5').getValue() || '').trim();
  if (!promptContents) {
    promptContents = getZodiacContentsPrompt(theme, subThemes);
    sheet.getRange('C5').setValue(promptContents);
  } else {
    // テンプレート変数を置換
    promptContents = promptContents
      .replace(/\{\{theme\}\}/g, theme)
      .replace(/\{\{subThemes\}\}/g, subThemes);
  }

  const startTime = new Date();
  const contentsJson = callGPT5(apiKey, promptContents);
  const endTime = new Date();

  const parsedData = parse12ZodiacContents(contentsJson);
  if (!parsedData) {
    SpreadsheetApp.getUi().alert('12星座コンテンツ生成に失敗しました。');
    return;
  }

  // E5以降に出力
  let currentRow = 5;
  const zodiacOrder = ['牡羊座', '牡牛座', '双子座', '蟹座', '獅子座', '乙女座', '天秤座', '蠍座', '射手座', '山羊座', '水瓶座', '魚座'];

  parsedData.contents.forEach((content, index) => {
    // サブテーマヘッダー
    sheet.getRange(currentRow, 5).setValue(`【${content.subtheme}】`).setFontWeight('bold').setBackground('#ffd966');
    currentRow++;

    // 12星座分のデータ
    zodiacOrder.forEach(zodiac => {
      const text = content.zodiac_texts[zodiac] || '';
      sheet.getRange(currentRow, 5).setValue(zodiac);
      sheet.getRange(currentRow, 6).setValue(text);
      currentRow++;
    });

    // 空行
    currentRow++;
  });

  // キャプション出力
  sheet.getRange(currentRow, 5).setValue('【Instagramキャプション】').setFontWeight('bold').setBackground('#b6d7a8');
  currentRow++;
  sheet.getRange(currentRow, 5, 1, 3).merge().setValue(parsedData.instagram_caption).setWrap(true);

  // ログ出力
  addLog(sheet, '12星座STEP2: コンテンツ生成', promptContents, contentsJson, startTime, endTime);
}

/* ===== 12星座JSONパース ===== */
function parse12ZodiacContents(text) {
  let cleaned = text.replace(/^```json|^```|```$/gmi, '').trim();
  const first = cleaned.indexOf('{');
  const last = cleaned.lastIndexOf('}');
  if (first >= 0 && last > first) cleaned = cleaned.slice(first, last + 1);

  try {
    const obj = JSON.parse(cleaned);
    if (!obj || !Array.isArray(obj.contents)) return null;
    return {
      contents: obj.contents,
      instagram_caption: String(obj.instagram_caption || '')
    };
  } catch {
    return null;
  }
}

/* ===== 12星座シート初期化 ===== */
function initialize12ZodiacSheet() {
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
  sheet.getRange('E1').setValue('💫 12星座コンテンツ');
  sheet.getRange('F1').setValue('内容');
  sheet.getRange('N1').setValue('📊 実行ログ');
  sheet.getRange('O1').setValue('リクエスト');
  sheet.getRange('P1').setValue('レスポンス');

  // 入力エリア（2行目）
  sheet.getRange('A2').setValue('テーマを入力（例：恋愛）');

  // サブヘッダー（4行目）
  sheet.getRange('B4').setValue('▼ STEP1プロンプト本文');
  sheet.getRange('C4').setValue('▼ STEP2プロンプト本文');
  sheet.getRange('D4').setValue('▼ STEP1出力本文');
  sheet.getRange('E4').setValue('▼ STEP2出力');

  // デフォルトプロンプトを配置（5行目から縦10行結合）
  const defaultPrompt1 = getZodiacThemesPrompt('{{theme}}');
  sheet.getRange('B5').setValue(defaultPrompt1);
  sheet.getRange('B5:B14').merge();

  const defaultPrompt2 = getZodiacContentsPrompt('{{theme}}', '{{subThemes}}');
  sheet.getRange('C5').setValue(defaultPrompt2);
  sheet.getRange('C5:C14').merge();

  // STEP1出力エリアを結合（D5:D14）
  sheet.getRange('D5:D14').merge();

  // フォーマット適用
  format12ZodiacSheet(sheet);

  SpreadsheetApp.getUi().alert('12星座シートを初期化しました。A2にテーマを入力してください。');
}

/* ===== 12星座シートフォーマッティング ===== */
function format12ZodiacSheet(sheet) {
  // ヘッダー行（1行目）をボールド＋背景色
  const headerRange = sheet.getRange('A1:P1');
  headerRange.setFontWeight('bold')
             .setBackground('#6aa84f')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  // サブヘッダー行（4行目）をボールド＋背景色
  const subHeaderRange = sheet.getRange('A4:P4');
  subHeaderRange.setFontWeight('bold')
                .setBackground('#93c47d')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');

  // 入力エリア（A2）
  sheet.getRange('A2').setBackground('#fff2cc');

  // プロンプトエリア（B5:B14, C5:C14）
  sheet.getRange('B5:B14').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');
  sheet.getRange('C5:C14').setBackground('#d9ead3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP1出力（D5:D14）
  sheet.getRange('D5:D14').setBackground('#cfe2f3')
                          .setWrap(true)
                          .setVerticalAlignment('top');

  // STEP2出力（E5以降）
  sheet.getRange('E5:F').setBackground('#d9d2e9').setWrap(true);

  // ログエリア（N列以降）
  sheet.getRange('N:P').setBackground('#ead1dc').setWrap(true);

  // 列幅調整
  sheet.setColumnWidth(1, 180);  // A列（入力）
  sheet.setColumnWidth(2, 450);  // B列（STEP1プロンプト）
  sheet.setColumnWidth(3, 450);  // C列（STEP2プロンプト）
  sheet.setColumnWidth(4, 350);  // D列（STEP1出力）
  sheet.setColumnWidth(5, 180);  // E列（星座/サブテーマ）
  sheet.setColumnWidth(6, 250);  // F列（内容）
  sheet.setColumnWidth(14, 150); // N列（タイムスタンプ）
  sheet.setColumnWidth(15, 350); // O列（リクエスト）
  sheet.setColumnWidth(16, 350); // P列（レスポンス）

  // 行の高さ調整
  sheet.setRowHeight(1, 40);  // ヘッダー行
  sheet.setRowHeight(4, 35);  // サブヘッダー行
  sheet.setRowHeights(5, 10, 60); // 5-14行目（結合セル用）
}
