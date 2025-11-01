/**
 * デフォルトプロンプトテンプレート
 * これらのプロンプトはシートのB2, B3セルに初期配置され、ユーザーが自由に編集できます
 */

/**
 * STEP1: ストーリー設計用プロンプト
 * @param {string} theme - テーマ（A2セルの値）
 * @param {string} method - 占い手法（A3セルの値）
 * @returns {string} プロンプトテキスト
 */
function getStoryDesignPrompt(theme, method) {
  return `あなたは日本語の「占い師のプロ」として、信頼・安心・教育的なSNSストーリーを設計します。
シリーズテーマ：「${theme}」
占い手法は**「${method}」のみ**を使用します。他の手法名は出しません。

目的：
- 読者が「私も大丈夫かも」「占いを受けてみたい」と感じ、プロフィール欄から予約に進む。
- 親しみやすさと専門性で信頼を築く（煽り・断定過多なし）。

序盤の共感（重要）：
- 1〜2本目では、本文に「本当のあなた」「安心」「報われない」「頑張ってきた」などの共感ワードを必ず含める（見出しには入れない）。

教育設計（読者体験）：
1) 誤解の解像度を上げる導入（軽い未完）
2) 仕組み・原理の解説（好奇心ギャップ）
3) 活用・変化の橋渡し（小さな成功の想起）
4) 不安の解毒と安心設計（倫理・プロの姿勢）

出力（B2にそのまま貼る文章）：
ストーリー全体のテーマ：
- ${theme}（手法：${method}）

全体の狙い：
（読者に与える安心・気づき・行動の変化を一文で）

シリーズ構成（4〜5本で完結）：
1. 導入（誤解・気づき）／タイトル案／心理的狙い
2. 原理（理解）／タイトル案／心理的狙い
3. 体験（応用）／タイトル案／心理的狙い
4. 安心（信頼）／タイトル案／心理的狙い

最後に：
占い師としての信念と姿勢を一文で`;
}

/**
 * STEP2: 7分割投稿生成用プロンプト
 * @param {string} method - 占い手法（A3セルの値）
 * @param {string} storyText - STEP1で生成されたストーリー設計テキスト
 * @returns {string} プロンプトテキスト
 */
function getRowsGenerationPrompt(method, storyText) {
  return `あなたは日本語の「占い師のプロ」です。手法は**「${method}」のみ**使用してください。
以下の設計に基づき、4〜5本のストーリー投稿＋共通のInstagramキャプション1本を生成します。

【ストーリー設計】
${storyText}

■ 各ストーリー投稿の制約：
- **誠実・上品・穏やか。教えるのではなく「気づかせる」口調**
- 体験・観察・原理を自然に織り交ぜ、依存を生まない助言。プライバシー配慮。
- **見出し（title）は超重要！共感フレーズで読者を引き込む**：
  - 「本当のあなた」「安心」「報われない」「頑張ってきた」「大丈夫」「見えてる」「わかってる」「辛かった」「無理しないで」等を積極的に使用
  - 読者が「これ私のこと！」と思わず足を止める見出しにする
- **投稿1と2**では本文にも共感表現を必ず含む
- **文字数を厳守**：
  - title：12文字以内（共感フレーズ必須、目を引く表現）
  - l1a：10文字以内／l1b：10文字以内
  - l2a：10文字以内／l2b：10文字以内（※1文を途中で分割可）
  - l3a：15文字以内／l3b：15文字以内
- 絵文字・ハッシュタグ・URL・@メンション禁止

■ Instagramキャプション（全投稿共通で1本のみ）：
**文字数：400〜600文字** で、視聴者との深いつながりを作る内容にしてください。

【構成】
1. **冒頭の共感（2〜3行）**
   - 読者の悩みや状況に寄り添う言葉
   - 「最近、こんなふうに感じていませんか？」など、読者の心の声を代弁

2. **シリーズの紹介（3〜4行）**
   - このストーリーシリーズで伝えたい核心メッセージ
   - ${method}の視点から見た気づきや学び
   - 「このシリーズでは...」「${method}で見ると...」

3. **具体的な価値提供（4〜5行）**
   - このシリーズを見ることで得られる変化
   - 読者が「自分も変われるかも」と思える希望
   - 具体例や小さな実践アドバイス

4. **コメント誘導（2〜3行）**
   - 読者に質問を投げかける
   - 例：「あなたはこのシリーズのどの言葉が一番響きましたか？」
   - 例：「同じように感じた方、コメント欄で教えてください💬」
   - 例：「今の気持ちを一言で表すなら？✨」

5. **予約導線（2行）**
   - 「あなたの悩みに寄り添う個別相談も受け付けています」
   - 「無料相談はプロフィール欄をご覧ください。今だけ特別に無料！」

6. **ハッシュタグ（1〜2行）**
   - #占い #スピリチュアル #恋愛相談 #恋愛サポート #星流命理 #${method} #自己理解 #心の癒し など8〜12個

【トーン・書式】
- 優しく、温かく、読者の味方である姿勢
- 押し付けがましくなく、自然な流れで行動を促す
- 絵文字を適度に使用（💫✨🌙💝など、スピリチュアルな雰囲気）
- **改行は必須**：2〜3文ごとに改行を入れ、読みやすく
- 各セクション間は空行（\\n\\n）で区切る
- 長い文章の塊にならないよう、適度に段落分け

【改行の例】
最近、自分らしさがわからなくなっていませんか？
周りに合わせすぎて、本当の気持ちを見失っていませんか？💫

このシリーズでは、${method}の視点から...
（空行）
あなたが本来持っている輝きを...
（空行）
コメント欄で教えてください💬
（空行）
無料相談はプロフィール欄を...

出力（JSONのみ／説明・コードフェンス禁止）：
{
  "posts": [
    {
      "title": "",
      "l1a": "", "l1b": "",
      "l2a": "", "l2b": "",
      "l3a": "", "l3b": ""
    }
  ],
  "instagram_caption": "上記構成に従った400〜600文字のキャプション。冒頭の共感→シリーズ紹介→価値提供→コメント誘導→予約導線→ハッシュタグの順。改行を効果的に使い、絵文字も適度に含める。"
}
※posts配列は4〜5本で完結させてください。`;
}

/**
 * 12星座STEP1: サブテーマ生成用プロンプト
 * @param {string} theme - メインテーマ（A2セルの値、例：恋愛）
 * @returns {string} プロンプトテキスト
 */
function getZodiacThemesPrompt(theme) {
  return `あなたは日本語の「占い師のプロ」として、Instagram投稿用の12星座別コンテンツを設計します。

メインテーマ：「${theme}」

目的：
- 読者が自分の星座や気になる相手の星座をチェックしたくなる
- 保存・シェアしたくなる実用的で共感できる内容
- プロフィール欄から無料占いや予約につながる導線

タスク：
メインテーマ「${theme}」に関連する、魅力的なサブテーマを3〜5個生成してください。

サブテーマの条件：
- 12星座それぞれの特徴が際立つテーマ
- 読者が「当たってる！」「気になる！」と感じるもの
- 各星座15文字以内の短文で表現できるもの
- 具体的で想像しやすいもの

例（恋愛テーマの場合）：
1. 恋に落ちる相手
2. ドキッとする瞬間
3. 愛情表現の仕方
4. 長続きする関係のコツ

出力形式（そのままテキストで、番号付きリスト）：
1. [サブテーマ1]
2. [サブテーマ2]
3. [サブテーマ3]
（3〜5個）

※説明やコメントは不要。サブテーマのリストのみを出力してください。`;
}

/**
 * 12星座STEP2: 12星座別短文＋キャプション生成用プロンプト
 * @param {string} theme - メインテーマ（A2セルの値）
 * @param {string} subThemes - STEP1で生成されたサブテーマリスト
 * @returns {string} プロンプトテキスト
 */
function getZodiacContentsPrompt(theme, subThemes) {
  return `あなたは日本語の「占い師のプロ」です。12星座別のInstagram投稿コンテンツを生成します。

メインテーマ：「${theme}」

【サブテーマ一覧】
${subThemes}

タスク：
各サブテーマごとに、12星座別の短文を生成してください。

制約：
- 各星座の短文は**15文字以内**（簡潔で印象的に）
- 12星座の順序は必ず以下の通り：牡羊座、牡牛座、双子座、蟹座、獅子座、乙女座、天秤座、蠍座、射手座、山羊座、水瓶座、魚座
- 占い師らしい誠実で上品なトーン
- 各星座の性質を活かした内容
- 絵文字・ハッシュタグは使用しない（本文のみ）

出力形式（JSON）：
{
  "contents": [
    {
      "subtheme": "[サブテーマ1]",
      "zodiac_texts": {
        "牡羊座": "",
        "牡牛座": "",
        "双子座": "",
        "蟹座": "",
        "獅子座": "",
        "乙女座": "",
        "天秤座": "",
        "蠍座": "",
        "射手座": "",
        "山羊座": "",
        "水瓶座": "",
        "魚座": ""
      }
    }
  ],
  "instagram_caption": ""
}

Instagramキャプションの条件：
- 冒頭で読者に問いかけ（例：「あなたはどんな相手に恋に落ちやすいですか？」）
- 自分の星座や気になる相手の特徴もチェックするよう誘導
- フォローといいねで恋愛運上昇・引き寄せが始まる旨を記載
- 無料占い実施中（プロフィールのリンク・固定投稿から参加）を明記
- 星座の性質を知ることで恋愛・人間関係がスムーズになることを伝える
- 保存を促す一文
- 最後に感謝のメッセージ
- 区切り線（┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈）
- 自己紹介（最強のご縁を引き寄せる占い師、フォローで恋愛運急上昇など）
- ストーリーズで毎日タロット占い・開運メッセージ配信中
- ハッシュタグ：#占い #恋愛占い師 #恋愛運上昇 #性格診断 #引き寄せの法則 など
- 全体で親しみやすく、でも押し付けがましくないトーン

※コードフェンスや説明は不要。JSONのみを出力してください。`;
}

/**
 * ランキングSTEP1: ランキング設計用プロンプト
 * @param {string} theme - ランキングのテーマ（A2セルの値、例：2025年の恋愛運）
 * @param {string} type1 - 固定軸：星座 or 誕生月（A3セルの値）
 * @param {string} type2 - 掛け合わせ軸：血液型 or 誕生月（A4セルの値）
 * @returns {string} プロンプトテキスト
 */
function getRankingDesignPrompt(theme, type1, type2) {
  // 組み合わせの説明を生成
  let combinationDesc = '';
  let totalCount = 0;

  if (type1 === '星座' && type2 === '血液型') {
    combinationDesc = '12星座 × 血液型（A、B、O、AB）';
    totalCount = 48;
  } else if (type1 === '誕生月' && type2 === '血液型') {
    combinationDesc = '12ヶ月 × 血液型（A、B、O、AB）';
    totalCount = 48;
  } else if (type1 === '星座' && type2 === '誕生月') {
    combinationDesc = '12星座 × 12ヶ月（誕生月）';
    totalCount = 144;
  } else if (type1 === '誕生月' && type2 === '誕生月') {
    combinationDesc = '誕生月 × 誕生月';
    totalCount = 144;
  }

  // 星座×誕生月の整合性に関する注意事項
  let consistencyNote = '';
  if (type1 === '星座' && type2 === '誕生月') {
    consistencyNote = `

【絶対厳守：星座と誕生月の整合性】
星座と誕生月が一致しない組み合わせは絶対に使用しないでください。
各星座には対応する誕生月が決まっています：

- 牡羊座（3/21-4/19）→ 3月、4月のみ
- 牡牛座（4/20-5/20）→ 4月、5月のみ
- 双子座（5/21-6/21）→ 5月、6月のみ
- 蟹座（6/22-7/22）→ 6月、7月のみ
- 獅子座（7/23-8/22）→ 7月、8月のみ
- 乙女座（8/23-9/22）→ 8月、9月のみ
- 天秤座（9/23-10/23）→ 9月、10月のみ
- 蠍座（10/24-11/22）→ 10月、11月のみ
- 射手座（11/23-12/21）→ 11月、12月のみ
- 山羊座（12/22-1/19）→ 12月、1月のみ
- 水瓶座（1/20-2/18）→ 1月、2月のみ
- 魚座（2/19-3/20）→ 2月、3月のみ

例：牡羊×3月、牡羊×4月は○、牡羊×5月は×（絶対NG）
例：獅子×7月、獅子×8月は○、獅子×6月は×（絶対NG）

ランキングでは、整合性のある組み合わせのみを使用してください。`;
  }

  return `あなたは日本語の「占い師」として、Instagram投稿用のランキングコンテンツを設計します。

ランキングテーマ：「${theme}」
分類タイプ：${type1} × ${type2}
合計組み合わせ数：${totalCount}通り${consistencyNote}

目的：
- 読者が自分の組み合わせや気になる人の順位をチェックしたくなる
- 保存・シェアしたくなる魅力的なランキング
- プロフィール欄から無料占いや予約につながる導線

タスク：
このランキングの全体コンセプトと順位付けの基準を設計してください。

設計内容：
1. ランキングの狙い・コンセプト（読者にどんな気づきや楽しさを与えるか）
2. 順位付けの基準（${type1}と${type2}の特性をどう組み合わせて評価するか）
3. 上位になりやすい傾向（どんな組み合わせが高ランクになるか）
4. 下位でもポジティブに（全ての順位に価値がある理由）
5. 占い師としてのメッセージ
6. 簡単な表現でわかりやすくして！
7. 小学生でもわかる単語だけを使うこと

出力形式（そのままテキストで）：
【ランキングコンセプト】
（1〜2行で）

【順位付けの基準】
・基準1: ...
・基準2: ...
・基準3: ...

【上位の傾向】
（1〜2行で）

【順位に関わらず大切なこと】
（1〜2行で）

【占い師からのメッセージ】
（1〜2行で）

※説明やコメントは不要。上記フォーマットのテキストのみを出力してください。`;
}

/**
 * ランキングSTEP2: ランキング30生成用プロンプト
 * @param {string} theme - ランキングのテーマ（A2セルの値）
 * @param {string} type1 - 固定軸：星座 or 誕生月（A3セルの値）
 * @param {string} type2 - 掛け合わせ軸：血液型 or 誕生月（A4セルの値）
 * @param {string} designText - STEP1で生成されたランキング設計テキスト
 * @returns {string} プロンプトテキスト
 */
function getRankingContentsPrompt(theme, type1, type2, designText) {
  // 組み合わせの範囲を生成
  const combinations1 = type1 === '星座'
    ? '牡羊、牡牛、双子、蟹、獅子、乙女、天秤、蠍、射手、山羊、水瓶、魚'
    : '1月、2月、3月、4月、5月、6月、7月、8月、9月、10月、11月、12月';

  const combinations2 = type2 === '血液型'
    ? 'A、B、O、AB'
    : '1月、2月、3月、4月、5月、6月、7月、8月、9月、10月、11月、12月';

  // サンプル組み合わせを生成
  let sampleCombination1, sampleCombination2;
  if (type1 === '星座' && type2 === '血液型') {
    sampleCombination1 = '牡羊×A';
    sampleCombination2 = '双子×O';
  } else if (type1 === '誕生月' && type2 === '血液型') {
    sampleCombination1 = '1月×A';
    sampleCombination2 = '3月×O';
  } else if (type1 === '星座' && type2 === '誕生月') {
    sampleCombination1 = '牡羊×3月';
    sampleCombination2 = '牡牛×5月';
  } else if (type1 === '誕生月' && type2 === '誕生月') {
    sampleCombination1 = '1月×12月';
    sampleCombination2 = '6月×7月';
  }

  // 星座×誕生月の整合性に関する注意事項
  let consistencyNote = '';
  if (type1 === '星座' && type2 === '誕生月') {
    consistencyNote = `

【絶対厳守：星座と誕生月の整合性】
★★★ 星座と誕生月が一致しない組み合わせは絶対に生成禁止 ★★★

各星座に対応する誕生月のみ使用可能：
- 牡羊（3/21-4/19）→ 3月、4月のみ ／ 例：牡羊×3月○、牡羊×4月○、牡羊×5月×
- 牡牛（4/20-5/20）→ 4月、5月のみ ／ 例：牡牛×4月○、牡牛×5月○、牡牛×6月×
- 双子（5/21-6/21）→ 5月、6月のみ
- 蟹（6/22-7/22）→ 6月、7月のみ
- 獅子（7/23-8/22）→ 7月、8月のみ
- 乙女（8/23-9/22）→ 8月、9月のみ
- 天秤（9/23-10/23）→ 9月、10月のみ
- 蠍（10/24-11/22）→ 10月、11月のみ
- 射手（11/23-12/21）→ 11月、12月のみ
- 山羊（12/22-1/19）→ 12月、1月のみ
- 水瓶（1/20-2/18）→ 1月、2月のみ
- 魚（2/19-3/20）→ 2月、3月のみ

上記以外の組み合わせは絶対に使用しないでください。
例えば「牡羊×6月」「獅子×1月」「魚×12月」などはNG（現実にあり得ない）。
30位全てが、この整合性を満たす組み合わせのみで構成されていることを確認してください。`;
  }

  return `あなたは日本語の「占い師」です。ランキングコンテンツを生成します。

ランキングテーマ：「${theme}」
分類タイプ：${type1} × ${type2}

【ランキング設計】
${designText}${consistencyNote}

タスク：
上記の設計に基づき、1位から30位までのランキングを生成してください。

組み合わせの範囲：
- ${type1}：${combinations1}
- ${type2}：${combinations2}

重要制約：
- 各順位の説明は**30文字以内**（簡潔で印象的に）
- ポジティブな表現（下位でも前向きに）
- 絵文字・ハッシュタグは使用しない（本文のみ）
- 同じ${type1}や${type2}が偏らないようバランスよく配置
- 簡単な表現でわかりやすくして！
- 小学生でもわかる単語だけを使うこと
- 漢字は使わないで、ひらがなで
- とにかく、簡単な表現で！！！

出力形式（JSON）：
{
  "rankings": [
    {
      "rank": 1,
      "combination": "${sampleCombination1}",
      "description": "30文字以内の説明文"
    },
    {
      "rank": 2,
      "combination": "${sampleCombination2}",
      "description": "30文字以内の説明文"
    }
  ],
  "instagram_caption": ""
}

Instagramキャプションの条件：
- 冒頭でランキングテーマを紹介（例：「【${theme}ランキングTOP30】」）
- 自分の順位や気になる人の順位もチェックするよう誘導
- フォローといいねで運気上昇・引き寄せが始まる旨を記載
- 無料占い実施中（プロフィールのリンク・固定投稿から参加）を明記
- ランキングは参考程度に、自分らしさを大切にするメッセージ
- 保存を促す一文
- 最後に感謝のメッセージ
- 区切り線（┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈）
- 自己紹介（最強のご縁を引き寄せる占い師、フォローで恋愛運急上昇など）
- ストーリーズで毎日タロット占い・開運メッセージ配信中
- ハッシュタグ：#占い #${theme} #ランキング #性格診断 #引き寄せの法則 など
- 全体で親しみやすく、でも押し付けがましくないトーン

※コードフェンスや説明は不要。JSONのみを出力してください。`;
}

/**
 * 今日の星座占い用のプロンプト
 * @param {string} date - 日付 (例: "2025年10月21日")
 * @return {string} 今日の星座占い用プロンプト
 */
function getTodayHoroscopePrompt(date) {
  return `あなたはプロの占い師です。${date}の星座占いを作成してください。

以下の条件に従ってJSON形式で出力してください：

1. 今日はどんな日か
   - スピリチュアルな観点から今日の全体的な運勢の概要
   - 今日のパワースポット（都道府県レベルまたは具体的な場所）

2. 12星座のランキング（1位〜12位）
   各星座について以下を含める：
   - 順位
   - 星座名
   - ラッキーカラー
   - ラッキーフード
   - ラッキーアクション
   - 恋愛運コメント（短文30文字程度）
   - 恋愛運コメント（長文60文字程度）
   - 総合運コメント（短文30文字程度）
   - 総合運コメント（長文60文字程度）

JSON形式：
{
  "today_overview": {
    "description": "今日はどんな日かのスピリチュアルな解説",
    "power_spot": "今日のパワースポット"
  },
  "rankings": [
    {
      "rank": 1,
      "zodiac": "おひつじ座",
      "lucky_color": "赤",
      "lucky_food": "トマト",
      "lucky_action": "朝日を浴びる",
      "love_short": "運命の出会いが訪れる予感",
      "love_long": "積極的に行動することで、素敵な出会いや関係の進展が期待できる一日です",
      "overall_short": "全てが好転する最高の一日",
      "overall_long": "仕事も人間関係も順調で、チャレンジしたことが良い結果につながる幸運日です"
    }
  ]
}

※コードフェンスや説明は不要。JSONのみを出力してください。`;
}

/**
 * タロット占い（A/B/C 3択）用のプロンプト
 * @param {Object} cardsForChoices - 各選択肢のカード情報 {A: [...], B: [...], C: [...]}
 * @return {string} タロット3択占い用プロンプト
 */
function getTarotPrompt(cardsForChoices) {
  // A, B, Cそれぞれのカード説明を生成（3枚ずつ）
  const choiceA = cardsForChoices.A.map((card, index) =>
    `${index + 1}枚目: ${card.name}（${card.position}）`
  ).join('\n');

  const choiceB = cardsForChoices.B.map((card, index) =>
    `${index + 1}枚目: ${card.name}（${card.position}）`
  ).join('\n');

  const choiceC = cardsForChoices.C.map((card, index) =>
    `${index + 1}枚目: ${card.name}（${card.position}）`
  ).join('\n');

  return `あなたはプロのタロット占い師です。A、B、Cの3択タロット占いを行います。
それぞれの選択肢について、引かれたカードからのメッセージを読み解いてください。

【選択肢A】
${choiceA}

【選択肢B】
${choiceB}

【選択肢C】
${choiceC}

タスク：
各選択肢（A/B/C）について、以下を生成してください：
1. そのカードが示す状況や意味の詳細説明（200〜300文字）
2. 💬本音：形式での一言メッセージ（30〜50文字）

さらに、Instagram投稿用のキャプション（800〜1,200文字）を生成してください。

制約：
- 占い師らしい誠実で上品なトーン
- 正位置と逆位置で異なる解釈を提供
- 読者の心に寄り添う温かいメッセージ
- 深い洞察と具体的なイメージを提供

Instagramキャプションの構成（800〜1,200文字）：

【必須構成】
1. テーマと引いたカードの紹介
   - テーマ：「〇〇」
   - 引いたカード：A 〇〇(正/逆) / B 〇〇(正/逆) / C 〇〇(正/逆)
   - 区切り線（──────────────）

2. A、B、Cそれぞれの詳細説明（各200〜300文字）
   - 区切り線で区切る
   - Aのカード名と位置
   - カードの意味と状況の解説
   - 箇条書きで具体的なポイント（・を使用）
   - 💬本音：「...」形式でまとめ

3. 総合メッセージ（100〜150文字）
   - 3つの選択肢全体から見えること
   - 区切り線で区切る

4. アドバイス（100〜150文字）
   - あなたへのアドバイス
   - ✔マークで箇条書き

5. ハッシュタグ
   - #タロット占い #3択占い #あの人の本音 #恋愛成就 など8〜12個
   - カード名も含める（例：#世界 #悪魔 #愚者逆）

【トーン】
- 深く寄り添う共感的な語り口
- 具体的で想像しやすい表現
- 読者が「わかる！」と思える描写
- 「あなた」と直接語りかける親密さ
- 改行を効果的に使い、読みやすく

【禁止事項】
- 占い予約への誘導文（「プロフィールから〜」など）は不要
- フォロー・いいね誘導は不要
- 区切り線で自己紹介は不要

出力形式（JSON）：
{
  "theme": "あの人が抱える、あなたへの本音",
  "card_summary": "A 世界(正) / B 悪魔(正) / C 愚者(逆)",
  "choices": [
    {
      "choice": "A",
      "card_name": "世界",
      "position": "正位置",
      "description": "このカードは「完成」「到達」「特別な結びつき」。あの人は、あなたとの関係を"中途半端なもの"とは見ていません。（200〜300文字の詳細説明）",
      "real_voice": "あなたは完成形。もうここで終わりたいくらい、理想なんだよ"
    },
    {
      "choice": "B",
      "card_name": "悪魔",
      "position": "正位置",
      "description": "...",
      "real_voice": "..."
    },
    {
      "choice": "C",
      "card_name": "愚者",
      "position": "逆位置",
      "description": "...",
      "real_voice": "..."
    }
  ],
  "overall_message": "あの人の中であなたは、「この人がいい」って決めてる相手（世界）...",
  "advice": "この並びは「拒まれたくない」「失いたくない」がすごく強いから...",
  "instagram_caption": "800〜1,200文字の完全なキャプション。上記構成に従い、テーマ紹介→A/B/C詳細→総合メッセージ→アドバイス→ハッシュタグの順。区切り線（──────────────）を効果的に使用。"
}

※コードフェンスや説明は不要。JSONのみを出力してください。`;
}

/**
 * タロットカード画像生成用のプロンプト
 * @param {string} cardName - カード名（例：「愚者」「ワンドのエース」）
 * @param {boolean} isMajor - 大アルカナかどうか
 * @return {string} タロットカード画像生成用プロンプト
 */
function getTarotImagePrompt(cardName, isMajor) {
  // 厳密なスタイルガイド（再現性を最大化）
  const strictStyleGuide = `
CRITICAL LAYOUT REQUIREMENTS (MUST FOLLOW EXACTLY):
- Portrait orientation, 2:3 aspect ratio (400px × 600px reference)
- Card COMPLETELY FILLS the entire image frame, edge-to-edge
- NO visible background or space outside the card borders
- Card bleeds to all four edges of the image

CARD STRUCTURE (PRECISE POSITIONING):
- Decorative border: INSIDE the card area, 15px from all edges, golden color (#FFD700)
- Top section (60px height): Roman numeral centered, serif font, 36pt, black on cream background
- Main illustration: Centered in middle area (450px height), rich symbolic imagery
- Bottom section (90px height): Card title centered, elegant serif font (Cinzel or similar), 28pt, black on cream background
- Background color: Aged parchment cream (#F5E6D3)

STRICT STYLE SPECIFICATIONS:
- Art style: Classical Rider-Waite-Smith tarot deck aesthetic
- Line work: Bold, clear outlines (3px thickness), black ink style
- Color palette: Rich jewel tones (sapphire blue #0F52BA, emerald green #50C878, ruby red #E0115F, amethyst purple #9966CC, golden yellow #FFD700)
- Shading: Cross-hatching and watercolor wash techniques
- Texture: Visible paper grain, slight aged appearance
- Symbolism: Traditional tarot iconography, NO modern elements

TYPOGRAPHY (EXACT SPECIFICATIONS):
- Font family: Cinzel or Trajan Pro (classical serif)
- Title position: 30px from bottom edge, perfectly centered horizontally
- Roman numeral position: 20px from top edge, perfectly centered horizontally
- Text color: Deep black (#000000) with subtle gold outline
- Letter spacing: 2px for title, 3px for numerals

FORBIDDEN ELEMENTS:
- Modern photography or realistic rendering
- Minimalist or abstract designs
- Rounded corners or modern card shapes
- Sans-serif fonts or modern typography
- Bright neon colors or gradients
- Any space/background visible around the card
`;

  // 大アルカナと小アルカナで異なる説明
  if (isMajor) {
    // 各カードの具体的なシンボリズムを追加
    const cardSymbolism = getMajorArcanaSymbolism(cardName);

    return `Create a traditional tarot card illustration for Major Arcana: "${cardName}" (${getCardEnglishName(cardName)}).

${strictStyleGuide}

SPECIFIC SYMBOLISM FOR THIS CARD:
${cardSymbolism}

COMPOSITION REQUIREMENTS:
1. Roman numeral at top (matching card number)
2. Central figure or scene embodying the card's archetypal meaning
3. Traditional symbols and iconography (${cardName}の伝統的なシンボル)
4. Rich background with relevant spiritual elements
5. Card title "${cardName}" at bottom in Japanese, with English subtitle "${getCardEnglishName(cardName)}"
6. Inner decorative border with mystical patterns (Celtic knots, sacred geometry)

COLOR GUIDANCE FOR ${cardName}:
Use colors that traditionally represent this card's energy and meaning.
Maintain consistency with Rider-Waite-Smith color symbolism.

QUALITY CHECKLIST:
✓ Card fills entire frame (no background visible)
✓ Border is INSIDE the card, not outside
✓ Typography is perfectly centered and legible
✓ Symbolism matches traditional tarot meanings
✓ Colors are rich and vibrant, not washed out
✓ Professional, printable quality illustration`;
  } else {
    // 小アルカナの詳細なガイドライン
    const suitInfo = getMinorArcanaSuitInfo(cardName);

    return `Create a traditional tarot card illustration for Minor Arcana: "${cardName}".

${strictStyleGuide}

SUIT-SPECIFIC REQUIREMENTS:
${suitInfo}

COMPOSITION REQUIREMENTS:
1. Suit symbols arranged according to card rank
2. ${cardName}に適した構図とシンボル配置
3. Element symbolism prominently featured
4. Court cards: Regal figure in traditional pose
5. Number cards: Geometric arrangement of suit symbols
6. Card title "${cardName}" at bottom in Japanese
7. Inner decorative border matching suit's element

RANK-SPECIFIC LAYOUT:
- Ace: Single large central symbol with divine light
- 2-10: Symmetrical arrangement of suit symbols
- Page: Youthful figure, standing pose, learning/discovery theme
- Knight: Dynamic figure, horse or movement, action theme
- Queen: Seated figure, throne, nurturing/mastery theme
- King: Authoritative figure, commanding presence, leadership theme

QUALITY CHECKLIST:
✓ Card fills entire frame (no background visible)
✓ Border is INSIDE the card, not outside
✓ Typography is perfectly centered and legible
✓ Suit symbolism is clear and prominent
✓ Colors match elemental correspondence
✓ Professional, printable quality illustration`;
  }
}

/**
 * 大アルカナカードの具体的なシンボリズムを返す
 */
function getMajorArcanaSymbolism(cardName) {
  const symbolism = {
    '愚者': 'Young traveler at cliff edge, small dog, white rose, mountain peaks, sun, beggar\'s bundle on staff. Colors: bright sky blue, white, yellow.',
    '魔術師': 'Figure at table with tools (wand, cup, sword, pentacle), infinity symbol above head, red robe, white undergarment, garden of roses and lilies. Colors: red, white, yellow.',
    '女教皇': 'Seated between two pillars (B and J), crescent moon crown, cross on chest, Torah scroll, pomegranate tapestry. Colors: blue, white, silver.',
    '女帝': 'Crowned figure on throne, Venus symbol, wheat field, waterfall, cushioned seat, scepter, heart-shaped shield. Colors: green, gold, red.',
    '皇帝': 'Armored figure on stone throne with ram heads, red robe, ankh scepter, orb, mountain background. Colors: red, orange, gold.',
    '教皇': 'Religious figure between two pillars, triple crown, crossed keys, two acolytes, papal cross. Colors: red, white, gold.',
    '恋人たち': 'Man and woman, angel above (Raphael), tree of knowledge with serpent, tree of life, mountain, sun. Colors: yellow, orange, flesh tones.',
    '戦車': 'Armored figure in chariot, two sphinxes (black and white), city background, crescent moons on shoulders, eight-pointed star crown. Colors: blue, gold, black, white.',
    '力': 'Woman gently closing lion\'s mouth, infinity symbol above head, white robe, flower garland, mountain background. Colors: white, yellow, green.',
    '隠者': 'Hooded figure on mountain peak, raised lantern with six-pointed star, long staff, grey robe, snowy ground. Colors: grey, yellow (lantern), white.',
    '運命の輪': 'Large wheel with symbolic creatures (angel, eagle, lion, bull), sphinx on top, serpent descending, Hebrew letters, alchemical symbols. Colors: blue, gold, red.',
    '正義': 'Crowned figure seated, scales in left hand, sword in right hand, purple robe, two pillars. Colors: purple, red, gold.',
    '吊された男': 'Figure hanging upside down from T-shaped tree, halo around head, hands behind back, serene expression. Colors: blue, red, yellow.',
    '死神': 'Skeletal figure on white horse, black banner with white rose, fallen king, pleading figures, sunset, distant towers. Colors: black, white, yellow.',
    '節制': 'Angelic figure with wings, pouring water between two cups, one foot on land one in water, sun crown, iris flowers, mountain path. Colors: blue, white, gold.',
    '悪魔': 'Horned figure on pedestal, inverted pentagram, chained man and woman, torches. Colors: black, red, grey, orange.',
    '塔': 'Tower struck by lightning, crown falling, two figures falling, flames, grey stone, stormy sky. Colors: grey, black, yellow, orange.',
    '星': 'Naked woman kneeling, pouring water into pool and land, large eight-pointed star, seven smaller stars, bird in tree. Colors: blue, yellow, green.',
    '月': 'Moon with face, two towers, two dogs/wolf howling, crayfish in pool, winding path. Colors: blue, yellow, grey.',
    '太陽': 'Large sun with face, naked child on white horse, sunflowers, brick wall. Colors: yellow, orange, white.',
    '審判': 'Angel Gabriel with trumpet, rising figures from coffins (man, woman, child), mountain range, cross banner. Colors: blue, white, flesh tones.',
    '世界': 'Dancing figure with wands, oval wreath, four corner creatures (angel, eagle, lion, bull), purple drapes. Colors: purple, blue, green, gold.'
  };

  return symbolism[cardName] || 'Traditional symbolic imagery representing the card\'s spiritual meaning.';
}

/**
 * 小アルカナのスート情報を返す
 */
function getMinorArcanaSuitInfo(cardName) {
  if (cardName.includes('ワンド')) {
    return `SUIT: Wands (Fire Element)
- Primary colors: Red (#E0115F), orange (#FF8C00), yellow (#FFD700)
- Symbol: Wooden staff/club with leaves sprouting
- Background: Desert, mountains, passionate energy
- Decorative motif: Flames, salamanders, diagonal dynamic lines
- Energy: Active, creative, passionate, entrepreneurial`;
  } else if (cardName.includes('カップ')) {
    return `SUIT: Cups (Water Element)
- Primary colors: Blue (#0F52BA), silver (#C0C0C0), aqua (#00CED1)
- Symbol: Golden chalice with water, sometimes overflowing
- Background: Rivers, seas, emotional landscapes, gardens
- Decorative motif: Waves, fish, lotuses, curved flowing lines
- Energy: Emotional, intuitive, loving, receptive`;
  } else if (cardName.includes('ソード')) {
    return `SUIT: Swords (Air Element)
- Primary colors: Yellow (#FFD700), white (#FFFFFF), pale blue (#ADD8E6)
- Symbol: Steel sword with sharp blade, upright or crossed
- Background: Cloudy skies, windswept landscapes, stormy weather
- Decorative motif: Clouds, birds, butterflies, sharp angular lines
- Energy: Mental, analytical, challenging, truth-seeking`;
  } else if (cardName.includes('ペンタクルス')) {
    return `SUIT: Pentacles (Earth Element)
- Primary colors: Green (#50C878), brown (#8B4513), gold (#FFD700)
- Symbol: Golden coin with five-pointed star (pentagram)
- Background: Gardens, fields, forests, material settings
- Decorative motif: Vines, flowers, fruits, stable geometric patterns
- Energy: Material, practical, stable, abundant`;
  }

  return 'Traditional minor arcana symbolism with appropriate suit elements.';
}

/**
 * タロットカード名の英訳を取得
 * @param {string} cardName - 日本語のカード名
 * @return {string} 英語のカード名
 */
function getCardEnglishName(cardName) {
  const translations = {
    '愚者': 'The Fool',
    '魔術師': 'The Magician',
    '女教皇': 'The High Priestess',
    '女帝': 'The Empress',
    '皇帝': 'The Emperor',
    '教皇': 'The Hierophant',
    '恋人たち': 'The Lovers',
    '戦車': 'The Chariot',
    '力': 'Strength',
    '隠者': 'The Hermit',
    '運命の輪': 'Wheel of Fortune',
    '正義': 'Justice',
    '吊された男': 'The Hanged Man',
    '死神': 'Death',
    '節制': 'Temperance',
    '悪魔': 'The Devil',
    '塔': 'The Tower',
    '星': 'The Star',
    '月': 'The Moon',
    '太陽': 'The Sun',
    '審判': 'Judgement',
    '世界': 'The World'
  };

  // 小アルカナの場合
  if (cardName.includes('ワンド')) return cardName.replace('ワンド', 'Wands');
  if (cardName.includes('カップ')) return cardName.replace('カップ', 'Cups');
  if (cardName.includes('ソード')) return cardName.replace('ソード', 'Swords');
  if (cardName.includes('ペンタクルス')) return cardName.replace('ペンタクルス', 'Pentacles');

  return translations[cardName] || cardName;
}

/**
 * ランキング名量産用のプロンプト
 * @return {string} ランキング名生成用プロンプト
 */
function getRankingTitlesPrompt() {
  return `あなたはプロの占い師です。Instagram投稿用の恋愛関係のランキング名を50個生成してください。

目的：
- スクロールを止めて「おっ」と目を引く魅力的なランキング名
- 保存・シェアしたくなるキャッチーなタイトル
- 自分や気になる人の順位をチェックしたくなる内容

ランキング名の条件：
- 恋愛に関する内容（片思い、両思い、結婚、復縁、相性など）
- 15〜30文字程度で簡潔に
- 「〇〇ランキング」「〇〇TOP30」「〇〇ベスト30」などの形式
- 数字を効果的に使用（例：「2025年」「30位」「12星座」など）
- 具体的で想像しやすい内容
- ポジティブで前向きな表現

ランキング名の例：
- 「2025年に結婚できる星座×血液型ランキングTOP30」
- 「片思いが成就しやすい誕生月ランキング」
- 「モテ期到来！恋愛運上昇中の星座TOP30」
- 「運命の人と出会える確率が高い血液型ランキング」
- 「復縁成功率が高い星座×誕生月TOP30」

各ランキング名には、簡単な説明（30文字以内）も付けてください。

出力形式（JSON）：
{
  "titles": [
    {
      "number": 1,
      "title": "2025年に結婚できる星座×血液型ランキングTOP30",
      "description": "今年結婚の可能性が高い組み合わせを大公開"
    },
    {
      "number": 2,
      "title": "片思いが成就しやすい誕生月ランキング",
      "description": "あの人を振り向かせる力を持つ誕生月"
    }
  ]
}

※コードフェンスや説明は不要。JSONのみを出力してください。
※必ず50個のランキング名を生成してください。`;
}

/**
 * ランキング名量産用のプロンプト（ネガティブ→ポジティブ寄り添い型）
 * @return {string} ランキング名生成用プロンプト
 */
function getRankingTitlesNegativeToPositivePrompt() {
  return `あなたはプロの占い師です。Instagram投稿用の恋愛関係のランキング名を50個生成してください。

特別なコンセプト：
ネガティブに見える特徴を、ポジティブで寄り添う形に変換したランキング名を作成します。
「あなたは一人じゃない」「その特徴は才能だよ」と読者を肯定し、勇気づける内容にしてください。

ネガティブ→ポジティブ変換の例：
- 「涙もろい人」→「感情豊かで心が優しい人ランキング」
- 「傷つきやすい人」→「繊細で共感力が高い人ランキング」
- 「元彼を忘れられない人」→「一途で愛情深い人ランキング」
- 「恋愛に臆病な人」→「慎重で大切な人を守れる人ランキング」
- 「依存しやすい人」→「人を信じる力が強い人ランキング」
- 「嫉妬深い人」→「愛情表現が豊かな人ランキング」
- 「すぐ別れたくなる人」→「自分の気持ちに正直な人ランキング」
- 「恋愛が長続きしない人」→「新しい出会いに恵まれやすい人ランキング」

目的：
- スクロールを止めて「これ、私のこと！」と共感する
- ネガティブな自己認識をポジティブに変換してあげる
- 「そんな自分も悪くない」と思える内容
- 保存・シェアしたくなる温かいメッセージ

ランキング名の条件：
- ネガティブに見える特徴を扱う（恋愛の悩み、弱み、コンプレックスなど）
- それをポジティブで優しい表現に変換する
- 15〜35文字程度で簡潔に
- 「〇〇ランキング」「〇〇な人TOP30」などの形式
- 読者が「救われた」「受け入れられた」と感じる表現
- 説教くさくならず、寄り添う優しいトーン

扱うテーマ例：
- 感情系：泣き虫、感情的、すぐ怒る、ネガティブ思考
- 恋愛行動系：重い、執着する、束縛する、追いかける
- 過去系：元彼を忘れられない、トラウマがある、恋愛恐怖症
- 性格系：優柔不断、自信がない、人見知り、気を使いすぎる
- 関係性：都合のいい女、二番手、友達止まり、片思いばかり

各ランキング名には、温かい説明（40文字以内）も付けてください。

出力形式（JSON）：
{
  "titles": [
    {
      "number": 1,
      "title": "涙もろくて感情豊かな人ランキングTOP30",
      "description": "その優しさと感受性は、あなたの魅力です"
    },
    {
      "number": 2,
      "title": "傷つきやすいけど共感力が高い人ランキング",
      "description": "繊細な心は、人を深く理解する才能です"
    }
  ]
}

※コードフェンスや説明は不要。JSONのみを出力してください。
※必ず50個のランキング名を生成してください。
※すべてのランキング名がネガティブ→ポジティブ変換の形式になっていることを確認してください。`;
}
