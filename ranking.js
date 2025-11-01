/**
 * ランキング30機能
 */

/* ===== まとめて実行（STEP1+2） ===== */
function generateRankingContent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:固定軸、A4:掛け合わせ軸）
  const theme = String(sheet.getRange('A2').getValue() || '').trim().replace(/　/g, '');
  const type1 = String(sheet.getRange('A3').getValue() || '').trim().replace(/　/g, '');
  const type2 = String(sheet.getRange('A4').getValue() || '').trim().replace(/　/g, '');

  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type1) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (!type2) { SpreadsheetApp.getUi().alert('A4 に「血液型」または「誕生月」を選択してください'); return; }

  if (type1 !== '星座' && type1 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }
  if (type2 !== '血液型' && type2 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A4 には「血液型」または「誕生月」のいずれかを入力してください');
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
  const designText = executeRankingStep1(sheet, apiKey, theme, type1, type2);

  // STEP2実行
  executeRankingStep2(sheet, apiKey, theme, type1, type2, designText);

  SpreadsheetApp.getUi().alert('完了：D5:E34にランキング設計、F5以降に横並び版、S5以降に縦並び版を出力しました。');
}

/* ===== STEP1のみ実行 ===== */
function generateRankingStep1Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:固定軸、A4:掛け合わせ軸）
  const theme = String(sheet.getRange('A2').getValue() || '').trim().replace(/　/g, '');
  const type1 = String(sheet.getRange('A3').getValue() || '').trim().replace(/　/g, '');
  const type2 = String(sheet.getRange('A4').getValue() || '').trim().replace(/　/g, '');

  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type1) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (!type2) { SpreadsheetApp.getUi().alert('A4 に「血液型」または「誕生月」を選択してください'); return; }

  if (type1 !== '星座' && type1 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }
  if (type2 !== '血液型' && type2 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A4 には「血液型」または「誕生月」のいずれかを入力してください');
    return;
  }

  // STEP1出力エリアのみクリア
  sheet.getRange('D5:E34').clearContent();

  // STEP1実行
  executeRankingStep1(sheet, apiKey, theme, type1, type2);

  SpreadsheetApp.getUi().alert('STEP1完了：D5:E34 にランキング設計を出力しました。');
}

/* ===== STEP2のみ実行 ===== */
function generateRankingStep2Only() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 入力取得（A2:テーマ、A3:固定軸、A4:掛け合わせ軸）
  const theme = String(sheet.getRange('A2').getValue() || '').trim().replace(/　/g, '');
  const type1 = String(sheet.getRange('A3').getValue() || '').trim().replace(/　/g, '');
  const type2 = String(sheet.getRange('A4').getValue() || '').trim().replace(/　/g, '');

  if (!theme) { SpreadsheetApp.getUi().alert('A2 にランキングテーマを入力してください（例：2025年の恋愛運）'); return; }
  if (!type1) { SpreadsheetApp.getUi().alert('A3 に「星座」または「誕生月」を選択してください'); return; }
  if (!type2) { SpreadsheetApp.getUi().alert('A4 に「血液型」または「誕生月」を選択してください'); return; }

  if (type1 !== '星座' && type1 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A3 には「星座」または「誕生月」のいずれかを入力してください');
    return;
  }
  if (type2 !== '血液型' && type2 !== '誕生月') {
    SpreadsheetApp.getUi().alert('A4 には「血液型」または「誕生月」のいずれかを入力してください');
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
  executeRankingStep2(sheet, apiKey, theme, type1, type2, designText);

  SpreadsheetApp.getUi().alert('STEP2完了：F5以降に横並び版、S5以降に縦並び版を出力しました。');
}

/* ===== STEP1実行（共通処理） ===== */
function executeRankingStep1(sheet, apiKey, theme, type1, type2) {
  // B5:B34からプロンプト取得、空なら初期化
  let promptDesign = String(sheet.getRange('B5').getValue() || '').trim();
  if (!promptDesign) {
    promptDesign = getRankingDesignPrompt(theme, type1, type2);
    sheet.getRange('B5').setValue(promptDesign);
  } else {
    // テンプレート変数を置換
    promptDesign = promptDesign
      .replace(/\{\{theme\}\}/g, theme)
      .replace(/\{\{type1\}\}/g, type1)
      .replace(/\{\{type2\}\}/g, type2);
  }

  const startTime = new Date();
  const designText = callGemini(apiKey, promptDesign);
  const endTime = new Date();

  // D5:E34に出力
  sheet.getRange('D5').setValue(designText);

  // STEP2のプロンプトを更新（type1とtype2が確定したので）
  const promptStep2 = getRankingContentsPrompt(theme, type1, type2, '{{designText}}');
  sheet.getRange('C5').setValue(promptStep2);

  // ログ出力
  addLog(sheet, 'ランキングSTEP1: ランキング設計', promptDesign, designText, startTime, endTime);

  return designText;
}

/* ===== STEP2実行（共通処理） ===== */
function executeRankingStep2(sheet, apiKey, theme, type1, type2, designText) {
  // C5:C34からプロンプト取得、空なら初期化
  let promptRanking = String(sheet.getRange('C5').getValue() || '').trim();
  if (!promptRanking) {
    promptRanking = getRankingContentsPrompt(theme, type1, type2, designText);
    sheet.getRange('C5').setValue(promptRanking);
  } else {
    // テンプレート変数を置換
    promptRanking = promptRanking
      .replace(/\{\{theme\}\}/g, theme)
      .replace(/\{\{type1\}\}/g, type1)
      .replace(/\{\{type2\}\}/g, type2)
      .replace(/\{\{designText\}\}/g, designText);
  }

  const startTime = new Date();
  const rankingJson = callGemini(apiKey, promptRanking);
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
         .setFontSize(22)
         .setBackground('#ffd966')
         .setHorizontalAlignment('center');
    currentRow++;

    // 組み合わせ行
    const combinationData = [];
    for (let i = 0; i < 10; i++) {
      const rankIndex = blockIndex * 10 + i;
      if (rankIndex < parsedData.rankings.length) {
        const item = parsedData.rankings[rankIndex];
        combinationData.push(item.combination);
      } else {
        combinationData.push('');
      }
    }
    sheet.getRange(currentRow, 6, 1, 10).setValues([combinationData])
         .setWrap(true)
         .setVerticalAlignment('middle')
         .setHorizontalAlignment('center')
         .setFontWeight('bold')
         .setFontSize(22);
    currentRow++;

    // 説明行
    const descriptionData = [];
    for (let i = 0; i < 10; i++) {
      const rankIndex = blockIndex * 10 + i;
      if (rankIndex < parsedData.rankings.length) {
        const item = parsedData.rankings[rankIndex];
        descriptionData.push(item.description);
      } else {
        descriptionData.push('');
      }
    }
    sheet.getRange(currentRow, 6, 1, 10).setValues([descriptionData])
         .setWrap(true)
         .setVerticalAlignment('top')
         .setHorizontalAlignment('center')
         .setFontSize(16);
    currentRow++;

    // ブロック間の空行
    currentRow++;
  }

  // ========== S5以降に縦並びで出力 ==========
  currentRow = 5;
  const verticalStartCol = 19; // S列

  // 縦並びタイトル
  sheet.getRange(currentRow, verticalStartCol, 1, 9).merge()
       .setValue('【縦並びレイアウト】')
       .setFontWeight('bold')
       .setBackground('#93c47d')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('center');
  currentRow++;

  // 3つのブロックに分けて出力（縦に10行ずつ）
  for (let blockIndex = 0; blockIndex < 3; blockIndex++) {
    const startRank = blockIndex * 10 + 1;
    const baseCol = verticalStartCol + (blockIndex * 3); // S列、V列、Y列

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
             .setFontSize(22)
             .setBackground('#ffd966')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');

        // 組み合わせ列
        sheet.getRange(currentRow + i, baseCol + 1)
             .setValue(item.combination)
             .setWrap(true)
             .setVerticalAlignment('middle')
             .setHorizontalAlignment('center')
             .setFontWeight('bold')
             .setFontSize(22);

        // 説明列
        sheet.getRange(currentRow + i, baseCol + 2)
             .setValue(item.description)
             .setWrap(true)
             .setVerticalAlignment('top')
             .setFontSize(16);
      }
    }
  }
  currentRow += 10;

  // 空行
  currentRow++;

  // キャプション出力（横並びエリアの下）
  const captionRow = 5 + 1 + 3 * 3 + 3; // タイトル + (ヘッダー+組み合わせ+説明)*3 + 空行*3 = 18
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

  // シート名を更新（数字を継承）
  const currentName = sheet.getName();
  const numberMatch = currentName.match(/\d+/); // 数字を抽出
  const newName = numberMatch ? `ランク${numberMatch[0]}` : 'ランク';
  sheet.setName(newName);

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
  sheet.getRange('S1:AA1').merge().setValue('🏆 STEP2出力（縦並び）');

  // ログヘッダー（35行目）
  sheet.getRange('AB35').setValue('📊 実行ログ');
  sheet.getRange('AC35').setValue('リクエスト');
  sheet.getRange('AD35').setValue('レスポンス');

  // 入力エリア（2-4行目）
  sheet.getRange('A2').setValue('ランキングテーマを入力（例：2025年の恋愛運）');
  sheet.getRange('A3').setValue('星座 or 誕生月を選択');
  sheet.getRange('A4').setValue('血液型 or 誕生月を選択');

  // A3にドロップダウンを設定（空白も選択可能）
  const rule1 = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', '星座', '誕生月'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A3').setDataValidation(rule1);

  // A4にドロップダウンを設定（空白も選択可能）
  const rule2 = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', '血液型', '誕生月'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A4').setDataValidation(rule2);

  // サブヘッダー（4行目）
  sheet.getRange('B4').setValue('▼ STEP1プロンプト本文');
  sheet.getRange('C4').setValue('▼ STEP2プロンプト本文');
  sheet.getRange('D4').setValue('▼ STEP1出力本文');
  sheet.getRange('F4:O4').merge().setValue('▼ STEP2出力（横並び）');
  sheet.getRange('S4:AA4').merge().setValue('▼ STEP2出力（縦並び）');

  // デフォルトプロンプトを配置（5行目から縦30行結合）
  const defaultPrompt1 = getRankingDesignPrompt('{{theme}}', '{{type1}}', '{{type2}}');
  sheet.getRange('B5').setValue(defaultPrompt1);
  sheet.getRange('B5:B34').merge();

  const defaultPrompt2 = getRankingContentsPrompt('{{theme}}', '{{type1}}', '{{type2}}', '{{designText}}');
  sheet.getRange('C5').setValue(defaultPrompt2);
  sheet.getRange('C5:C34').merge();

  // STEP1出力エリアを結合（D5:E34）
  sheet.getRange('D5:E34').merge();

  // フォーマット適用
  formatRankingSheet(sheet);

  SpreadsheetApp.getUi().alert('ランキングシートを初期化しました。\nA2: テーマ\nA3: 星座 or 誕生月\nA4: 血液型 or 誕生月');
}

/* ===== ランキングシートフォーマッティング ===== */
function formatRankingSheet(sheet) {
  // ヘッダー行（1行目）をボールド＋背景色
  const headerRange = sheet.getRange('A1:AD1');
  headerRange.setFontWeight('bold')
             .setBackground('#e69138')
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  // サブヘッダー行（4行目）をボールド＋背景色
  const subHeaderRange = sheet.getRange('A4:AD4');
  subHeaderRange.setFontWeight('bold')
                .setBackground('#f6b26b')
                .setFontColor('#ffffff')
                .setHorizontalAlignment('center');

  // 入力エリア（A2:A4）
  sheet.getRange('A2:A4').setBackground('#fff2cc');

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

  // STEP2出力（S5:AA以降）- 縦並び9列
  sheet.getRange('S:AA').setBackground('#d9d2e9').setWrap(true);

  // ログヘッダー行（35行目）
  sheet.getRange('AB35:AD35').setFontWeight('bold')
                           .setBackground('#c27ba0')
                           .setFontColor('#ffffff')
                           .setHorizontalAlignment('center');

  // ログエリア（36行目以降）
  sheet.getRange('AB36:AD').setBackground('#ead1dc').setWrap(true);

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
  // S〜AA列（縦並びランキング：順位、組み合わせ、説明 × 3グループ）
  sheet.setColumnWidth(19, 60);  // S列（1位〜10位：順位）
  sheet.setColumnWidth(20, 150); // T列（1位〜10位：組み合わせ）
  sheet.setColumnWidth(21, 200); // U列（1位〜10位：説明）
  sheet.setColumnWidth(22, 60);  // V列（11位〜20位：順位）
  sheet.setColumnWidth(23, 150); // W列（11位〜20位：組み合わせ）
  sheet.setColumnWidth(24, 200); // X列（11位〜20位：説明）
  sheet.setColumnWidth(25, 60);  // Y列（21位〜30位：順位）
  sheet.setColumnWidth(26, 150); // Z列（21位〜30位：組み合わせ）
  sheet.setColumnWidth(27, 200); // AA列（21位〜30位：説明）
  // ログ列
  sheet.setColumnWidth(28, 150); // AB列（タイムスタンプ）
  sheet.setColumnWidth(29, 350); // AC列（リクエスト）
  sheet.setColumnWidth(30, 350); // AD列（レスポンス）

  // 行の高さ調整
  sheet.setRowHeight(1, 40);  // ヘッダー行
  sheet.setRowHeight(4, 35);  // サブヘッダー行
  sheet.setRowHeights(5, 30, 60); // 5-34行目（結合セル用）
  sheet.setRowHeight(35, 35); // ログヘッダー行
}
