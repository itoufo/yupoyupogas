/**
 * メインエントリーポイント
 * メニュー定義のみ
 */

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
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🏆 ランキング30')
      .addItem('STEP1+2：まとめて実行（ランキング設計→30位生成）', 'generateRankingContent')
      .addSeparator()
      .addItem('STEP1のみ：ランキング設計（GPT-5）', 'generateRankingStep1Only')
      .addItem('STEP2のみ：ランキング30位生成（GPT-5）', 'generateRankingStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initializeRankingSheet'))
    .addToUi();
}
