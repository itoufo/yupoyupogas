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

  // 既存出力クリア（D5:E34, F5:X以降）
  sheet.getRange('D5:E34').clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    // F列（6）からX列（24）までクリア（横並び10列＋縦並び6列＋余裕）
    sheet.getRange(5, 6, Math.max(1, lastRow - 4), 19).clearContent();
  }

  // STEP1実行
  const designText = executeRankingStep1(sheet, apiKey, theme, type);

  // STEP2実行
  executeRankingStep2(sheet, apiKey, theme, type, designText);

  SpreadsheetApp.getUi().alert('完了：D5:E34にランキング設計、F5以降に横並び版、S5以降に縦並び版を出力しました。');
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

  // STEP2出力エリアのみクリア（F5:X以降）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    // F列（6）からX列（24）までクリア（横並び10列＋縦並び6列＋余裕）
    sheet.getRange(5, 6, Math.max(1, lastRow - 4), 19).clearContent();
  }

  // STEP2実行
  executeRankingStep2(sheet, apiKey, theme, type, designText);

  SpreadsheetApp.getUi().alert('STEP2完了：F5以降に横並び版、S5以降に縦並び版を出力しました。');
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

  // ========== F5以降に横並びで出力 ==========
  let currentRow = 5;

  // 横並びタイトル
  sheet.getRange(currentRow, 6, 1, 10).merge()
       .setValue('【横並びレイアウト】')
       .setFontWeight('bold')
       .setBackground('#93c47d')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  currentRow++;

  // 3つのブロックに分けて出力
  for (let blockIndex = 0; blockIndex < 3; blockIndex++) {
    const startRank = blockIndex * 10 + 1;

    // ヘッダー行（順位）
    const headerData = [];
    for (let i = 0; i < 10; i++) {
      headerData.push(`${startRank + i}位`);
    }
    sheet.getRange(currentRow, 6, 1, 10).setValues([headerData])
         .setFontWeight('bold')
         .setBackground('#ffd966')
         .setHorizontalAlignment('center');
    currentRow++;

    // 内容行（組み合わせ＋説明）
    const contentData = [];
    for (let i = 0; i < 10; i++) {
      const rankIndex = blockIndex * 10 + i;
      if (rankIndex < parsedData.rankings.length) {
        const item = parsedData.rankings[rankIndex];
        contentData.push(`${item.combination}\n${item.description}`);
      } else {
        contentData.push('');
      }
    }
    sheet.getRange(currentRow, 6, 1, 10).setValues([contentData])
         .setWrap(true)
         .setVerticalAlignment('top')
         .setHorizontalAlignment('center');
    currentRow++;

    // ブロック間の空行
    currentRow++;
  }

  // ========== S5以降に縦並びで出力 ==========
  currentRow = 5;
  const verticalStartCol = 19; // S列

  // 縦並びタイトル
  sheet.getRange(currentRow, verticalStartCol, 1, 6).merge()
       .setValue('【縦並びレイアウト】')
       .setFontWeight('bold')
       .setBackground('#93c47d')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  currentRow++;

  // 3つのブロックに分けて出力（縦に10行ずつ）
  for (let blockIndex = 0; blockIndex < 3; blockIndex++) {
    const startRank = blockIndex * 10 + 1;
    const baseCol = verticalStartCol + (blockIndex * 2); // S列、U列、W列

    // 10行分のデータを縦に出力
    for (let i = 0; i < 10; i++) {
      const rank = startRank + i;
      const rankIndex = blockIndex * 10 + i;

      if (rankIndex < parsedData.rankings.length) {
        const item = parsedData.rankings[rankIndex];

        // 順位列
        sheet.getRange(currentRow + i, baseCol)
             .setValue(`${rank}位`)
             .setFontWeight('bold')
             .setBackground('#ffd966')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');

        // 内容列
        sheet.getRange(currentRow + i, baseCol + 1)
             .setValue(`${item.combination}\n${item.description}`)
             .setWrap(true)
             .setVerticalAlignment('top');
      }
    }
  }
  currentRow += 10;

  // 空行
  currentRow++;

  // キャプション出力（横並びエリアの下）
  const captionRow = 5 + 1 + 2 * 3 + 3; // タイトル + (ヘッダー+内容)*3 + 空行*3
  sheet.getRange(captionRow, 6, 1, 10).merge()
       .setValue('【Instagramキャプション】')
       .setFontWeight('bold')
       .setBackground('#b6d7a8')
       .setHorizontalAlignment('center');

  sheet.getRange(captionRow + 1, 6, 1, 10).merge()
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
  sheet.getRange('F1:O1').merge().setValue('🏆 STEP2出力（横並び）');
  sheet.getRange('S1:X1').merge().setValue('🏆 STEP2出力（縦並び）');

  // ログヘッダー（35行目）
  sheet.getRange('Y35').setValue('📊 実行ログ');
  sheet.getRange('Z35').setValue('リクエスト');
  sheet.getRange('AA35').setValue('レスポンス');

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
  sheet.getRange('F4:O4').merge().setValue('▼ STEP2出力（横並び）');
  sheet.getRange('S4:X4').merge().setValue('▼ STEP2出力（縦並び）');

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
  const headerRange = sheet.getRange('A1:AA1');
  headerRange.setFontWeight('bold')
             .setBackground('#e69138')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  // サブヘッダー行（4行目）をボールド＋背景色
  const subHeaderRange = sheet.getRange('A4:AA4');
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

  // STEP2出力（F5:O以降）- 横並び10列
  sheet.getRange('F:O').setBackground('#fce5cd').setWrap(true);

  // STEP2出力（S5:X以降）- 縦並び6列
  sheet.getRange('S:X').setBackground('#d9d2e9').setWrap(true);

  // ログヘッダー行（35行目）
  sheet.getRange('Y35:AA35').setFontWeight('bold')
                           .setBackground('#c27ba0')
                           .setFontColor('#ffffff')
                           .setHorizontalAlignment('center');

  // ログエリア（36行目以降）
  sheet.getRange('Y36:AA').setBackground('#ead1dc').setWrap(true);

  // 列幅調整
  sheet.setColumnWidth(1, 200);  // A列（入力）
  sheet.setColumnWidth(2, 450);  // B列（STEP1プロンプト）
  sheet.setColumnWidth(3, 450);  // C列（STEP2プロンプト）
  sheet.setColumnWidth(4, 350);  // D列（STEP1出力）
  sheet.setColumnWidth(5, 50);   // E列（結合用）
  sheet.setColumnWidths(6, 10, 150); // F-O列（横並びランキング10列、各150px）
  // P〜R列は空白
  sheet.setColumnWidth(16, 50);  // P列（空白）
  sheet.setColumnWidth(17, 50);  // Q列（空白）
  sheet.setColumnWidth(18, 50);  // R列（空白）
  // S〜X列（縦並びランキング）
  sheet.setColumnWidth(19, 60);  // S列（順位）
  sheet.setColumnWidth(20, 200); // T列（内容）
  sheet.setColumnWidth(21, 60);  // U列（順位）
  sheet.setColumnWidth(22, 200); // V列（内容）
  sheet.setColumnWidth(23, 60);  // W列（順位）
  sheet.setColumnWidth(24, 200); // X列（内容）
  // ログ列
  sheet.setColumnWidth(25, 150); // Y列（タイムスタンプ）
  sheet.setColumnWidth(26, 350); // Z列（リクエスト）
  sheet.setColumnWidth(27, 350); // AA列（レスポンス）

  // 行の高さ調整
  sheet.setRowHeight(1, 40);  // ヘッダー行
  sheet.setRowHeight(4, 35);  // サブヘッダー行
  sheet.setRowHeights(5, 30, 60); // 5-34行目（結合セル用）
  sheet.setRowHeight(35, 35); // ログヘッダー行
}
