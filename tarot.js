/**
 * タロット占い機能
 */

/* ===== タロットカード定義 ===== */
function getTarotCards() {
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

  // ランダムに10枚のカードを選択
  const selectedCards = selectRandomCards(10);

  // 実行
  executeTarot(sheet, apiKey, selectedCards);

  SpreadsheetApp.getUi().alert('完了：D5以降にタロット占いを出力しました。');
}

/* ===== 実行処理（共通） ===== */
function executeTarot(sheet, apiKey, selectedCards) {
  // プロンプト生成（関数から直接取得）
  const prompt = getTarotPrompt(selectedCards);

  const startTime = new Date();
  const response = callGemini(apiKey, prompt);
  const endTime = new Date();

  // JSONパース
  const parsedData = parseTarotData(response);
  if (!parsedData) {
    throw new Error('JSONのパースに失敗しました。Geminiの応答を確認してください。');
  }

  // シートに出力
  outputTarotToSheet(sheet, parsedData, selectedCards);

  // ログ出力（V列 = 22列目）
  addLogForTarot(sheet, 'タロット占い', prompt, response, startTime, endTime);
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
  sheet.getRange('A1').setValue('タロット占い');
  sheet.getRange('A2').setValue('※カードはランダムに選択されます');
  sheet.getRange('D1').setValue('出力エリア →');

  // 列幅設定
  sheet.setColumnWidth(1, 200);  // A列
  sheet.setColumnWidth(2, 30);   // B列: 空白
  sheet.setColumnWidth(3, 30);   // C列: 空白
  sheet.setColumnWidth(4, 100);  // D列: 番号
  sheet.setColumnWidth(5, 200);  // E列: カード名
  sheet.setColumnWidth(6, 100);  // F列: 向き
  sheet.setColumnWidth(7, 400);  // G列: メッセージ

  SpreadsheetApp.getUi().alert('シートを初期化しました！\n「タロット占いを生成」を実行してください。');
}

/* ===== シートへの出力 ===== */
function outputTarotToSheet(sheet, data, selectedCards) {
  // 出力開始行
  let currentRow = 5;

  // タイトル
  sheet.getRange(currentRow, 4, 1, 4)
    .merge()
    .setValue('タロット占い - 10枚のカードからのメッセージ')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e6d7ff');
  currentRow += 2;

  // ヘッダー行
  sheet.getRange(currentRow, 4).setValue('番号').setFontWeight('bold').setBackground('#d4c5f9').setHorizontalAlignment('center');
  sheet.getRange(currentRow, 5).setValue('カード名').setFontWeight('bold').setBackground('#d4c5f9').setHorizontalAlignment('center');
  sheet.getRange(currentRow, 6).setValue('向き').setFontWeight('bold').setBackground('#d4c5f9').setHorizontalAlignment('center');
  sheet.getRange(currentRow, 7).setValue('メッセージ').setFontWeight('bold').setBackground('#d4c5f9').setHorizontalAlignment('center');
  currentRow++;

  // 各カードの情報を出力
  data.cards.forEach((card, index) => {
    sheet.getRange(currentRow, 4).setValue(card.number).setHorizontalAlignment('center');
    sheet.getRange(currentRow, 5).setValue(card.name).setHorizontalAlignment('left');
    sheet.getRange(currentRow, 6).setValue(card.position).setHorizontalAlignment('center');
    sheet.getRange(currentRow, 7).setValue(card.message).setWrap(true).setVerticalAlignment('top');

    // 行の高さを調整
    sheet.setRowHeight(currentRow, 60);

    currentRow++;
  });

  // 枠線を追加
  const dataRange = sheet.getRange(7, 4, data.cards.length + 1, 4);
  dataRange.setBorder(true, true, true, true, true, true);
}
