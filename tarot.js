/**
 * タロット占い機能
 */

/* ===== タロットカード定義 ===== */
function getTarotCards() {
  // 大アルカナ（22枚）のみ使用
  const majorArcana = [
    '愚者', '魔術師', '女教皇', '女帝', '皇帝',
    '教皇', '恋人たち', '戦車', '力', '隠者',
    '運命の輪', '正義', '吊された男', '死神', '節制',
    '悪魔', '塔', '星', '月', '太陽',
    '審判', '世界'
  ];

  return majorArcana;
}

/* ===== 全タロットカード定義（画像生成用） ===== */
function getAllTarotCards() {
  // 大アルカナ（22枚）
  const majorArcana = [
    '愚者', '魔術師', '女教皇', '女帝', '皇帝',
    '教皇', '恋人たち', '戦車', '力', '隠者',
    '運命の輪', '正義', '吊された男', '死神', '節制',
    '悪魔', '塔', '星', '月', '太陽',
    '審判', '世界'
  ];

  // 小アルカナ（56枚）
  const suits = ['ワンド', 'カップ', 'ソード', 'ペンタクルス'];
  const ranks = ['エース', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'ペイジ', 'ナイト', 'クイーン', 'キング'];

  const minorArcana = [];
  suits.forEach(suit => {
    ranks.forEach(rank => {
      minorArcana.push(`${suit}の${rank}`);
    });
  });

  return [...majorArcana, ...minorArcana];
}

/* ===== ランダムにカードを選択 ===== */
function selectRandomCards(count) {
  const allCards = getTarotCards();
  const selected = [];
  const usedIndices = new Set();

  while (selected.length < count) {
    const randomIndex = Math.floor(Math.random() * allCards.length);
    if (!usedIndices.has(randomIndex)) {
      usedIndices.add(randomIndex);
      const position = Math.random() < 0.5 ? '正位置' : '逆位置';
      selected.push({
        name: allCards[randomIndex],
        position: position
      });
    }
  }

  return selected;
}

/* ===== まとめて実行 ===== */
function generateTarot() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  // 既存出力クリア（D5以降）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 4, Math.max(1, lastRow - 4), 20).clearContent();
  }

  // A, B, Cそれぞれに3枚ずつカードを選択（合計9枚）
  const cardsForChoices = {
    A: selectRandomCards(3),
    B: selectRandomCards(3),
    C: selectRandomCards(3)
  };

  // 実行
  executeTarot(sheet, apiKey, cardsForChoices);

  SpreadsheetApp.getUi().alert('完了：A/B/C 3択タロット占いを出力しました。');
}

/* ===== 実行処理（共通） ===== */
function executeTarot(sheet, apiKey, cardsForChoices) {
  // プロンプト生成（関数から直接取得）
  const prompt = getTarotPrompt(cardsForChoices);

  const startTime = new Date();
  const response = callGemini(apiKey, prompt);
  const endTime = new Date();

  // JSONパース
  const parsedData = parseTarotData(response);
  if (!parsedData) {
    throw new Error('JSONのパースに失敗しました。Geminiの応答を確認してください。');
  }

  // シートに出力
  outputTarotToSheet(sheet, parsedData);

  // ログ出力（V列 = 22列目）
  addLogForTarot(sheet, 'タロット3択', prompt, response, startTime, endTime);
}

/* ===== タロット専用ログ出力（V列 = 22列目） ===== */
function addLogForTarot(sheet, stepName, request, response, startTime, endTime) {
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestSummary = `[${stepName}]\n実行時間: ${duration}秒\n\nプロンプト:\n${request.substring(0, 500)}${request.length > 500 ? '...' : ''}`;
  const responseSummary = `レスポンス:\n${response.substring(0, 500)}${response.length > 500 ? '...' : ''}`;

  const logColumn = 22;  // V列
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
function initializeTarotSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // シートをクリア（1行目以外）
  const lastRow = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
  }

  // ヘッダー設定
  sheet.getRange('A1').setValue('タロット3択占い（A/B/C）');
  sheet.getRange('A2').setValue('※大アルカナ22枚から各選択肢3枚ずつランダムに選択');
  sheet.getRange('D1').setValue('出力エリア →');

  // 列幅設定
  sheet.setColumnWidth(1, 200);  // A列
  sheet.setColumnWidth(2, 30);   // B列: 空白
  sheet.setColumnWidth(3, 30);   // C列: 空白
  sheet.setColumnWidth(4, 150);  // D列: ラベル
  sheet.setColumnWidth(5, 600);  // E列: 内容

  SpreadsheetApp.getUi().alert('シートを初期化しました！\n「タロット3択を生成」を実行してください。');
}

/* ===== シートへの出力 ===== */
function outputTarotToSheet(sheet, data) {
  // 出力開始行
  let currentRow = 5;

  // タイトル
  sheet.getRange(currentRow, 4, 1, 2)
    .merge()
    .setValue('タロット3択占い（A/B/C）')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e6d7ff');
  currentRow += 2;

  // テーマ
  sheet.getRange(currentRow, 4).setValue('テーマ').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(currentRow, 5).setValue(data.theme).setWrap(true);
  currentRow++;

  // カード概要
  sheet.getRange(currentRow, 4).setValue('引いたカード').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(currentRow, 5).setValue(data.card_summary).setWrap(true);
  currentRow += 2;

  // A, B, Cそれぞれの詳細
  data.choices.forEach((choice) => {
    // 選択肢のヘッダー
    sheet.getRange(currentRow, 4, 1, 2)
      .merge()
      .setValue(`【選択肢 ${choice.choice}】 ${choice.card_name}（${choice.position}）`)
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#d4c5f9')
      .setHorizontalAlignment('center');
    currentRow++;

    // 詳細説明
    sheet.getRange(currentRow, 4).setValue('詳細説明').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.getRange(currentRow, 5).setValue(choice.description).setWrap(true);
    sheet.setRowHeight(currentRow, 100);
    currentRow++;

    // 本音メッセージ
    sheet.getRange(currentRow, 4).setValue('💬本音').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.getRange(currentRow, 5).setValue(choice.real_voice).setWrap(true);
    currentRow += 2;
  });

  // 総合メッセージ
  sheet.getRange(currentRow, 4).setValue('総合メッセージ').setFontWeight('bold').setBackground('#ffe6e6');
  sheet.getRange(currentRow, 5).setValue(data.overall_message).setWrap(true);
  sheet.setRowHeight(currentRow, 80);
  currentRow++;

  // アドバイス
  sheet.getRange(currentRow, 4).setValue('アドバイス').setFontWeight('bold').setBackground('#e6f7ff');
  sheet.getRange(currentRow, 5).setValue(data.advice).setWrap(true);
  sheet.setRowHeight(currentRow, 80);
  currentRow += 2;

  // Instagramキャプション
  sheet.getRange(currentRow, 4).setValue('Instagramキャプション').setFontWeight('bold').setBackground('#fff9e6');
  sheet.getRange(currentRow, 5).setValue(data.instagram_caption).setWrap(true);
  sheet.setRowHeight(currentRow, 200);
}

/* ===== タロットカード画像一括生成 ===== */
function generateAllTarotImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません。');

  // 確認ダイアログ
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'タロットカード画像生成確認',
    '78枚のタロットカード画像を生成します。\n\n' +
    '予想コスト: 約$3.12（78枚 × $0.04）\n' +
    '予想時間: 約10〜15分\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました。');
    return;
  }

  // Google Driveフォルダを作成または取得
  const folderName = 'タロットカード画像';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // 全カード取得（画像生成用は78枚全て）
  const allCards = getAllTarotCards();
  const majorArcanaCount = 22;

  // シート初期化
  initializeTarotImageSheet();

  // 出力開始行
  let currentRow = 5;

  // エラーカードを記録
  const errorCards = [];
  let successCount = 0;

  // 大アルカナ（1-22）
  ui.alert(`大アルカナ（22枚）の生成を開始します...`);
  for (let i = 0; i < majorArcanaCount; i++) {
    const cardName = allCards[i];
    try {
      sheet.getRange(currentRow, 4).setValue(`${i + 1}/${78} 生成中...`);
      sheet.getRange(currentRow, 5).setValue(cardName);
      SpreadsheetApp.flush();

      const imageUrl = generateSingleTarotImage(apiKey, cardName, true, folder);

      // シートに出力
      sheet.getRange(currentRow, 4).setValue(i + 1);
      sheet.getRange(currentRow, 5).setValue(cardName);
      sheet.getRange(currentRow, 6).setValue('大アルカナ');
      sheet.getRange(currentRow, 7).setValue(imageUrl);

      // 画像をセルに埋め込み（H列）
      try {
        insertImageToCell(sheet, currentRow, 8, imageUrl);
        successCount++;
      } catch (embedError) {
        errorCards.push({ name: cardName, row: currentRow, error: '画像埋め込み失敗' });
      }

      currentRow++;
      SpreadsheetApp.flush();
      Utilities.sleep(1000);  // レート制限対策
    } catch (error) {
      const errorMsg = error.message.length > 50 ? error.message.substring(0, 50) + '...' : error.message;
      sheet.getRange(currentRow, 4).setValue('エラー');
      sheet.getRange(currentRow, 7).setValue(errorMsg);
      errorCards.push({ name: cardName, row: currentRow, error: errorMsg });
      Logger.log(`Error generating ${cardName}: ${error.message}`);
      currentRow++;
    }
  }

  // 小アルカナ（23-78）
  ui.alert(`小アルカナ（56枚）の生成を開始します...`);
  for (let i = majorArcanaCount; i < allCards.length; i++) {
    const cardName = allCards[i];
    try {
      sheet.getRange(currentRow, 4).setValue(`${i + 1}/${78} 生成中...`);
      sheet.getRange(currentRow, 5).setValue(cardName);
      SpreadsheetApp.flush();

      const imageUrl = generateSingleTarotImage(apiKey, cardName, false, folder);

      // シートに出力
      sheet.getRange(currentRow, 4).setValue(i + 1);
      sheet.getRange(currentRow, 5).setValue(cardName);
      sheet.getRange(currentRow, 6).setValue('小アルカナ');
      sheet.getRange(currentRow, 7).setValue(imageUrl);

      // 画像をセルに埋め込み（H列）
      try {
        insertImageToCell(sheet, currentRow, 8, imageUrl);
        successCount++;
      } catch (embedError) {
        errorCards.push({ name: cardName, row: currentRow, error: '画像埋め込み失敗' });
      }

      currentRow++;
      SpreadsheetApp.flush();
      Utilities.sleep(1000);  // レート制限対策
    } catch (error) {
      const errorMsg = error.message.length > 50 ? error.message.substring(0, 50) + '...' : error.message;
      sheet.getRange(currentRow, 4).setValue('エラー');
      sheet.getRange(currentRow, 7).setValue(errorMsg);
      errorCards.push({ name: cardName, row: currentRow, error: errorMsg });
      Logger.log(`Error generating ${cardName}: ${error.message}`);
      currentRow++;
    }
  }

  // 結果レポート
  let resultMessage = `完了：78枚中${successCount}枚の画像生成に成功しました。`;
  if (errorCards.length > 0) {
    resultMessage += `\n\nエラーが発生したカード（${errorCards.length}枚）:\n`;
    errorCards.forEach(card => {
      resultMessage += `- ${card.name}（行${card.row}）: ${card.error}\n`;
    });
    resultMessage += '\n※エラーが発生したカードは、手動で再実行できます。';
  }

  ui.alert(resultMessage);
}

/* ===== 単一カード画像生成 ===== */
function generateSingleTarotImage(apiKey, cardName, isMajor, folder) {
  // プロンプト生成
  const prompt = getTarotImagePrompt(cardName, isMajor);

  // DALL-E 3で画像生成（リトライ機能付き）
  const startTime = new Date();
  const imageUrl = callDallE3WithRetry(apiKey, prompt, 3);
  const endTime = new Date();

  Logger.log(`Generated ${cardName} in ${(endTime - startTime) / 1000}s`);

  // 画像をダウンロード（リトライ付き）
  const imageBlob = downloadImageWithRetry(imageUrl, 3);
  imageBlob.setName(`${cardName}.png`);

  // Google Driveに保存
  const file = folder.createFile(imageBlob);
  const driveUrl = file.getUrl();

  Logger.log(`Saved ${cardName} to Drive: ${driveUrl}`);

  return imageUrl;  // 画像URLを返す
}

/* ===== 画像ダウンロード（リトライ付き） ===== */
function downloadImageWithRetry(imageUrl, maxRetries = 3) {
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`画像ダウンロード（試行${attempt}/${maxRetries}）`);
      const blob = UrlFetchApp.fetch(imageUrl).getBlob();

      if (attempt > 1) {
        Logger.log(`成功: ${attempt}回目の試行でダウンロード成功`);
      }

      return blob;
    } catch (error) {
      lastError = error;
      Logger.log(`ダウンロードエラー（試行${attempt}/${maxRetries}）: ${error.message}`);

      if (attempt < maxRetries) {
        const waitTime = attempt * 1000;  // 1秒、2秒、3秒...
        Logger.log(`${waitTime / 1000}秒待機してリトライします...`);
        Utilities.sleep(waitTime);
      }
    }
  }

  // 全リトライ失敗
  throw new Error(`${maxRetries}回のリトライ後もダウンロード失敗: ${lastError.message}`);
}

/* ===== 画像をセルに埋め込み（リトライ付き） ===== */
function insertImageToCell(sheet, row, col, imageUrl, maxRetries = 3) {
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`画像埋め込み（試行${attempt}/${maxRetries}）`);

      // 画像を数式で埋め込み
      const formula = `=IMAGE("${imageUrl}", 1)`;
      sheet.getRange(row, col).setFormula(formula);
      sheet.setRowHeight(row, 200);  // 行の高さを調整

      if (attempt > 1) {
        Logger.log(`成功: ${attempt}回目の試行で画像埋め込み成功`);
      }

      return;  // 成功
    } catch (error) {
      lastError = error;
      Logger.log(`画像埋め込みエラー（試行${attempt}/${maxRetries}）: ${error.message}`);

      if (attempt < maxRetries) {
        const waitTime = attempt * 500;  // 0.5秒、1秒、1.5秒...
        Logger.log(`${waitTime / 1000}秒待機してリトライします...`);
        Utilities.sleep(waitTime);
      }
    }
  }

  // 全リトライ失敗
  Logger.log(`Error inserting image at row ${row}: ${lastError.message}`);
  sheet.getRange(row, col).setValue('画像読み込みエラー');
  throw new Error(`画像埋め込み失敗: ${lastError.message}`);
}

/* ===== タロット画像シート初期化 ===== */
function initializeTarotImageSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // シートをクリア（1行目以外）
  const lastRow = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clear();
  }

  // ヘッダー設定
  sheet.getRange('A1').setValue('タロットカード画像生成');
  sheet.getRange('D1').setValue('番号').setFontWeight('bold').setBackground('#e6d7ff');
  sheet.getRange('E1').setValue('カード名').setFontWeight('bold').setBackground('#e6d7ff');
  sheet.getRange('F1').setValue('種類').setFontWeight('bold').setBackground('#e6d7ff');
  sheet.getRange('G1').setValue('画像URL').setFontWeight('bold').setBackground('#e6d7ff');
  sheet.getRange('H1').setValue('プレビュー').setFontWeight('bold').setBackground('#e6d7ff');

  // 列幅設定
  sheet.setColumnWidth(1, 30);   // A列: 空白
  sheet.setColumnWidth(2, 30);   // B列: 空白
  sheet.setColumnWidth(3, 30);   // C列: 空白
  sheet.setColumnWidth(4, 60);   // D列: 番号
  sheet.setColumnWidth(5, 200);  // E列: カード名
  sheet.setColumnWidth(6, 100);  // F列: 種類
  sheet.setColumnWidth(7, 400);  // G列: 画像URL
  sheet.setColumnWidth(8, 200);  // H列: プレビュー

  SpreadsheetApp.getUi().alert('シートを初期化しました！\n「タロットカード画像を一括生成」を実行してください。');
}
