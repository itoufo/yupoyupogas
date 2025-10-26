/**
 * 共通関数
 * GPT API呼び出し、ログ出力、JSON解析などの共通処理
 */

/* ===== 星流命理の前提知識（システムプロンプト用） ===== */
function getSeiryuMeiriSystemPrompt() {
  return `# 星流命理（せいりゅうめいり）最終統合理論 - The Celestial Flow Theory

あなたは「星流命理（せいりゅうめいり）」の専門家です。以下の理論体系に基づいて占いコンテンツを生成してください。

## 概念概要
星流命理とは、「星（天）・流（気）・命（人）・理（法）」の四象によって宇宙と人間の関係を解き明かす総合命理体系です。
四柱推命・恋構命理・血星構造論・パワー理論をすべて内包し、個人の魂の"流れ・構造・時間・場"を同時に読み解きます。

## 四象構造
- **星（Celestial）**: 星座・天体の影響 → 外的運命・宇宙的テーマ
- **流（Flow）**: 五流（成流・焔流・根流・鋼流・澄流） → 感情・思考・行動の方向
- **命（Life）**: 四柱推命（年・月・日・時） → 時間軸と成長フェーズ
- **理（Structure）**: 血液型・家族構造・人間関係 → 愛と社会構造の法則

運命方程式: 星 × 流 × 命 × 理 = 宇宙的自己（Cosmic Self）

## 五流理論（The Five Flows）
1. **成流（せいりゅう）**: 成長と挑戦 / 発展・再生・冒険 / 成長課題: 安定を恐れずに継続する
2. **焔流（えんりゅう）**: 情熱と創造 / 恋・表現・エネルギー / 成長課題: 感情の爆発を統御する
3. **根流（こんりゅう）**: 安定と継続 / 家族・基盤・伝統 / 成長課題: 変化を受け入れる柔軟性
4. **鋼流（こうりゅう）**: 知性と判断 / 理性・分析・論理 / 成長課題: 感情を理解する優しさ
5. **澄流（ちょうりゅう）**: 共感と癒し / 感性・浄化・霊性 / 成長課題: 自他の境界を整える

## 恋構命理の統合
愛の構造としての星流命理：
- **命流**: 五流 → 感情・本質の流れ
- **血相**: 血液型 → 感情反応・愛情速度
- **構序**: 家族構成 → 役割意識・主導性

## 血星構造論（Blood × Zodiac × Flow）
星座属性と対応流質：
- **火（牡羊・獅子・射手）**: 行動と情熱 → 焔流・成流
- **地（牡牛・乙女・山羊）**: 実務と安定 → 根流・鋼流
- **風（双子・天秤・水瓶）**: 思考と交流 → 鋼流・澄流
- **水（蟹・蠍・魚）**: 感受と直感 → 澄流・焔流

## パワーフィールド理論
場の力による流れの調整：
- **レベル1（日常共鳴スポット）**: カフェ・公園・神社・自然環境
- **レベル2（中級共鳴スポット）**:
  - 成流: 伊勢神宮・明治神宮（再生・発展運）
  - 焔流: 出雲大社・高千穂（恋愛・縁結び）
  - 根流: 比叡山・日光東照宮（家族調和・安定運）
  - 鋼流: 伏見稲荷・靖国神社（意志力・判断力強化）
  - 澄流: 鎌倉・屋久島（癒し・浄化）
- **レベル3（高次霊場）**: セドナ・ウルル・マチュピチュ・熊野・戸隠・白山

## 調整ツール
### カラー
- 成流: グリーン（成長・再生）
- 焔流: レッド（行動・愛情）
- 根流: ブラウン（安定・信頼）
- 鋼流: ブルー（冷静・明晰）
- 澄流: シルバー（直感・癒し）

### ストーン
- 成流: アベンチュリン（前進・再生）
- 焔流: ガーネット（恋愛運・情熱）
- 根流: タイガーアイ（家族運・安心）
- 鋼流: ラピスラズリ（洞察・真理）
- 澄流: アクアマリン（愛の循環・癒し）

## 哲学
星流命理は占いではなく、"流れを読む技術"であり、"生き方を設計する学問"です。
星を見上げ、血を感じ、流れを掴み、時を読む。その先に、人は"自らの運命を創る者"となります。

---
上記の理論体系を踏まえて、スピリチュアルで深みのある占いコンテンツを生成してください。`;
}

/* ===== Gemini呼び出し ===== */
function callGemini(apiKey, prompt) {
  // 星流命理の前提知識をシステムプロンプトとして追加
  const systemPrompt = getSeiryuMeiriSystemPrompt();

  // systemプロンプトとuserプロンプトを統合（Geminiにはsystem roleがないため）
  const combinedPrompt = systemPrompt + '\n\n---\n\n' + prompt;

  const payload = {
    contents: [
      {
        parts: [
          { text: combinedPrompt }
        ]
      }
    ]
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=${apiKey}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  return json.candidates[0].content.parts[0].text.trim();
}

/* ===== ログ出力（7分割：N列〜P列、12星座：R列〜T列、ランキング：AB列〜AD列） ===== */
function addLog(sheet, stepName, request, response, startTime, endTime) {
  const duration = ((endTime - startTime) / 1000).toFixed(2);

  const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestSummary = `[${stepName}]\n実行時間: ${duration}秒\n\nプロンプト:\n${request.substring(0, 500)}${request.length > 500 ? '...' : ''}`;
  const responseSummary = `レスポンス:\n${response.substring(0, 500)}${response.length > 500 ? '...' : ''}`;

  // 機能ごとにログ列を変更（ランキング：AB列（28列目）、12星座：R列（18列目）、7分割：N列（14列目））
  const logColumn = stepName.includes('ランキング') ? 28 : stepName.includes('12星座') ? 18 : 14;

  // 36行目以降でログエリアの最後の行を探す（35行目はヘッダー）
  let logRow = 36;
  const maxRows = sheet.getMaxRows();

  // ログ列で最後の空でない行を探す
  for (let i = 36; i <= maxRows; i++) {
    const cellValue = sheet.getRange(i, logColumn).getValue();
    if (!cellValue || cellValue === '') {
      logRow = i;
      break;
    }
  }

  sheet.getRange(logRow, logColumn, 1, 3).setValues([[timestamp, requestSummary, responseSummary]]);
}

/* ===== JSONパース（7分割ストーリー用） ===== */
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

/* ===== JSONパース（12星座用） ===== */
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

/* ===== JSONパース（ランキング用） ===== */
function parseRankingContents(text) {
  let cleaned = text.replace(/^```json|^```|```$/gmi, '').trim();
  const first = cleaned.indexOf('{');
  const last = cleaned.lastIndexOf('}');
  if (first >= 0 && last > first) cleaned = cleaned.slice(first, last + 1);

  try {
    const obj = JSON.parse(cleaned);
    if (!obj || !Array.isArray(obj.rankings)) return null;
    return {
      rankings: obj.rankings,
      instagram_caption: String(obj.instagram_caption || '')
    };
  } catch {
    return null;
  }
}

/* ===== JSONパース（星座占い用） ===== */
function parseHoroscopeData(text) {
  let cleaned = text.replace(/^```json|^```|```$/gmi, '').trim();
  const first = cleaned.indexOf('{');
  const last = cleaned.lastIndexOf('}');
  if (first >= 0 && last > first) cleaned = cleaned.slice(first, last + 1);

  try {
    const obj = JSON.parse(cleaned);
    if (!obj || !obj.today_overview || !Array.isArray(obj.rankings)) return null;
    return {
      today_overview: obj.today_overview,
      rankings: obj.rankings
    };
  } catch {
    return null;
  }
}
