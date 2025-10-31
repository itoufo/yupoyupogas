/**
 * メインエントリーポイント
 * メニュー定義のみ
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔮 Gemini占い投稿生成')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📖 7分割ストーリー')
      .addItem('STEP1+2：まとめて実行（ストーリー設計→7分割）', 'generateFortuneProStoryAndRows')
      .addSeparator()
      .addItem('STEP1のみ：ストーリー設計（Gemini）', 'generateStep1Only')
      .addItem('STEP2のみ：7分割＋IGキャプ生成（Gemini）', 'generateStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initializeSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⭐ 12星座別コンテンツ')
      .addItem('STEP1+2：まとめて実行（サブテーマ生成→12星座）', 'generate12ZodiacContent')
      .addSeparator()
      .addItem('STEP1のみ：サブテーマ生成（Gemini）', 'generate12ZodiacStep1Only')
      .addItem('STEP2のみ：12星座＋IGキャプ生成（Gemini）', 'generate12ZodiacStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initialize12ZodiacSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🏆 ランキング30')
      .addItem('STEP1+2：まとめて実行（ランキング設計→30位生成）', 'generateRankingContent')
      .addSeparator()
      .addItem('STEP1のみ：ランキング設計（Gemini）', 'generateRankingStep1Only')
      .addItem('STEP2のみ：ランキング30位生成（Gemini）', 'generateRankingStep2Only')
      .addSeparator()
      .addItem('シート初期化（ヘッダー＋プロンプト配置）', 'initializeRankingSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🌟 今日の星座占い')
      .addItem('今日の星座占いを生成（Gemini）', 'generateTodayHoroscope')
      .addSeparator()
      .addItem('シート初期化（ヘッダー配置）', 'initializeHoroscopeSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🃏 タロット占い')
      .addItem('タロット占いを生成（Gemini）', 'generateTarot')
      .addSeparator()
      .addItem('シート初期化（ヘッダー配置）', 'initializeTarotSheet'))
    .addToUi();
}
