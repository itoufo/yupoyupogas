# Gemini占い投稿生成ツール - 開発ドキュメント

## プロジェクト概要

Google Apps Script (GAS) を使用した、Instagram投稿用の占いコンテンツ自動生成ツールです。
OpenAI Gemini-5 APIを活用し、複数の占いフォーマットに対応しています。

## ファイル構成

```
gas/
├── main.js           # メニュー定義のみ
├── common.js         # 共通関数（Gemini呼び出し、ログ、JSONパース）
├── prompts.js        # 全機能のプロンプトテンプレート
├── story.js          # 7分割ストーリー機能
├── zodiac.js         # 12星座別コンテンツ機能
├── ranking.js        # ランキング30機能
├── horoscope.js      # 今日の星座占い機能
└── appsscript.json   # GAS設定ファイル
```

## 共通実装ルール

### 1. APIキーの管理

すべての機能で統一された方法でAPIキーを取得：

```javascript
const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');
```

**設定方法**:
GASエディタで `プロジェクトの設定` → `スクリプト プロパティ` → `GEMINI_API_KEY` を追加

### 2. プロンプト管理

プロンプトはシートのB列、C列などに保存され、ユーザーが編集可能：

```javascript
// B5セルからプロンプト取得、空なら関数から初期化
let prompt = String(sheet.getRange('B5').getValue() || '').trim();
if (!prompt) {
  prompt = getPromptFunction('{{variable}}');  // テンプレート変数を含む
  sheet.getRange('B5').setValue(prompt);
}

// テンプレート変数を実際の値に置換
prompt = prompt.replace(/\{\{variable\}\}/g, actualValue);
```

**テンプレート変数の例**:
- `{{theme}}` - テーマ（ランキング、12星座など）
- `{{type}}` - 星座 or 誕生月
- `{{date}}` - 日付

### 3. 関数命名規則

各機能は以下の関数構成で統一：

```javascript
// まとめて実行（メニューから呼ばれる）
function generate[FeatureName]() {
  // 入力検証
  // 出力エリアクリア
  // execute関数を呼び出し
}

// STEP1のみ実行（2段階の場合）
function generate[FeatureName]Step1Only() { ... }

// STEP2のみ実行（2段階の場合）
function generate[FeatureName]Step2Only() { ... }

// 実行処理（共通化された処理）
function execute[FeatureName]Step1(sheet, apiKey, ...) { ... }
function execute[FeatureName]Step2(sheet, apiKey, ...) { ... }

// シート初期化
function initialize[FeatureName]Sheet() { ... }
```

### 4. ログ出力

共通関数 `addLog` を使用（common.js）：

```javascript
addLog(sheet, stepName, request, response, startTime, endTime);
```

ログは機能ごとに異なる列に出力：
- 7分割ストーリー: N列（14列目）
- 12星座別コンテンツ: R列（18列目）
- ランキング30: AB列（28列目）
- 今日の星座占い: R列（18列目）

### 5. JSONパース

各機能専用のパース関数を `common.js` に実装：

```javascript
// 7分割ストーリー用
function parsePostsObjectsWithCaption(text) { ... }

// 12星座別コンテンツ用
function parse12ZodiacContents(text) { ... }

// ランキング30用
function parseRankingContents(text) { ... }

// 今日の星座占い用
function parseHoroscopeData(text) { ... }
```

すべてのパース関数は以下の処理を行う：
1. コードフェンス（```json）の除去
2. 最初の `{` から最後の `}` までを抽出
3. JSON.parse()
4. 必要なフィールドの検証
5. エラー時は `null` を返す

## 各機能の詳細

### 1. 7分割ストーリー機能 (story.js)

Instagram用の7分割ストーリーコンテンツを生成。

**入力**:
- A2: テーマ（例：「2025年の恋愛運」）

**処理フロー**:
- STEP1: ストーリー設計（全体コンセプト、各スライドの役割）
- STEP2: 7分割コンテンツ生成 + Instagramキャプション

**出力**:
- D5:E34: STEP1の設計内容
- F5以降: 7分割コンテンツ（横並び）

### 2. 12星座別コンテンツ機能 (zodiac.js)

12星座それぞれのコンテンツを生成。

**入力**:
- A2: メインテーマ（例：「2025年の恋愛運」）

**処理フロー**:
- STEP1: 12個のサブテーマ生成
- STEP2: 各星座のコンテンツ生成 + Instagramキャプション

**出力**:
- D5:E34: STEP1のサブテーマ一覧
- F5以降: 12星座コンテンツ（横並び12列）

### 3. ランキング30機能 (ranking.js)

星座 or 誕生月 × 血液型（48通り）から30位までのランキング生成。

**入力**:
- A2: ランキングテーマ（例：「2025年の恋愛運」）
- A3: 「星座」または「誕生月」（ドロップダウン選択）

**処理フロー**:
- STEP1: ランキング設計（コンセプト、順位付け基準）
- STEP2: 1位〜30位のランキング生成 + Instagramキャプション

**出力**:
- D5:E34: STEP1の設計内容
- F5以降: 横並び版（1-10位、11-20位、21-30位を3行に分けて表示）
- S5以降: 縦並び版（10行×3グループで表示）

**レイアウト詳細**:

横並び版（F列〜O列）:
```
1位   2位   3位   ... 10位
組み  組み  組み  ... 組み
説明  説明  説明  ... 説明

11位  12位  13位  ... 20位
組み  組み  組み  ... 組み
説明  説明  説明  ... 説明

21位  22位  23位  ... 30位
組み  組み  組み  ... 組み
説明  説明  説明  ... 説明
```

縦並び版（S列〜AA列）:
```
1位  組み  説明  |  11位 組み 説明  |  21位 組み 説明
2位  組み  説明  |  12位 組み 説明  |  22位 組み 説明
...
10位 組み  説明  |  20位 組み 説明  |  30位 組み 説明
```

### 4. 今日の星座占い機能 (horoscope.js)

12星座の今日の運勢ランキングを生成。

**入力**:
- A2: 日付（空の場合は自動的に今日の日付）

**処理フロー**:
- 1段階で完結（日付 → 12星座ランキング + 今日の概要）

**出力**:
- D5以降: 今日の概要、パワースポット、12星座ランキング

**出力項目**:
- 今日はどんな日？（スピリチュアルな観点）
- 今日のパワースポット
- 12星座ランキング（1位〜12位）
  - ラッキーカラー
  - ラッキーフード
  - ラッキーアクション
  - 恋愛運コメント（短文30文字／長文60文字）
  - 総合運コメント（短文30文字／長文60文字）

## プロンプト設計 (prompts.js)

すべてのプロンプトは `prompts.js` に関数として定義：

```javascript
// 7分割ストーリー用
function getStoryDesignPrompt(theme) { ... }
function getRowsGenerationPrompt(theme, designText) { ... }

// 12星座別コンテンツ用
function getZodiacThemesPrompt(theme) { ... }
function getZodiacContentsPrompt(theme, themesText) { ... }

// ランキング30用
function getRankingDesignPrompt(theme, type) { ... }
function getRankingContentsPrompt(theme, type, designText) { ... }

// 今日の星座占い用
function getTodayHoroscopePrompt(date) { ... }
```

すべてのプロンプトは以下の共通フォーマット：
- 役割定義（「あなたはプロの占い師です」）
- タスク説明
- JSON出力形式の明示
- 制約条件（文字数、トーンなど）
- 最後に `※コードフェンスや説明は不要。JSONのみを出力してください。`

## メニュー構成 (main.js)

```
🔮 Gemini占い投稿生成
├─ 📖 7分割ストーリー
│  ├─ STEP1+2：まとめて実行（ストーリー設計→7分割）
│  ├─ STEP1のみ：ストーリー設計（Gemini-5）
│  ├─ STEP2のみ：7分割＋IGキャプ生成（Gemini-5）
│  └─ シート初期化（ヘッダー＋プロンプト配置）
├─ ⭐ 12星座別コンテンツ
│  ├─ STEP1+2：まとめて実行（サブテーマ生成→12星座）
│  ├─ STEP1のみ：サブテーマ生成（Gemini-5）
│  ├─ STEP2のみ：12星座＋IGキャプ生成（Gemini-5）
│  └─ シート初期化（ヘッダー＋プロンプト配置）
├─ 🏆 ランキング30
│  ├─ STEP1+2：まとめて実行（ランキング設計→30位生成）
│  ├─ STEP1のみ：ランキング設計（Gemini-5）
│  ├─ STEP2のみ：ランキング30位生成（Gemini-5）
│  └─ シート初期化（ヘッダー＋プロンプト配置）
└─ 🌟 今日の星座占い
   ├─ 今日の星座占いを生成（Gemini-5）
   └─ シート初期化（ヘッダー配置）
```

## デプロイメント

### 初期セットアップ

```bash
# claspのインストール
npm install -g @google/clasp

# ログイン
clasp login

# プロジェクトのクローン（既存プロジェクトの場合）
clasp clone [スクリプトID]
```

### 開発フロー

```bash
# ローカルで編集後、GASにプッシュ
clasp push

# GASの最新版をローカルに取得
clasp pull
```

### 環境変数設定

GASエディタで以下を設定：
1. `プロジェクトの設定` → `スクリプト プロパティ`
2. `プロパティを追加` → キー: `GEMINI_API_KEY`, 値: `AIzaSy...`

## トラブルシューティング

### clasp push でエラーが出る

```bash
# ログインし直す
clasp login

# プッシュ
clasp push
```

### APIキーエラー

エラーメッセージ: `GEMINI_API_KEY が設定されていません。`

→ GASエディタでスクリプトプロパティに `GEMINI_API_KEY` を設定

### JSONパースエラー

Geminiの応答がJSON形式でない場合に発生。
ログを確認してGeminiの応答を確認し、プロンプトを調整。

## 今後の拡張

新しい機能を追加する場合の手順：

1. **prompts.jsにプロンプト関数を追加**
   - テンプレート変数を使用（`{{variable}}`形式）
   - JSON出力形式を明示

2. **common.jsにJSONパース関数を追加**
   - 既存のパース関数を参考に実装
   - エラー処理を忘れずに

3. **新しい機能用のjsファイルを作成**
   - `generate[FeatureName]()`: まとめて実行
   - `execute[FeatureName]()`: 実際の処理
   - `initialize[FeatureName]Sheet()`: シート初期化
   - `output[FeatureName]ToSheet()`: 出力処理

4. **main.jsにメニューを追加**
   - 他の機能と統一したメニュー構造

5. **clasp pushでデプロイ**
   ```bash
   clasp push
   ```

## 参考リソース

- [Google Apps Script 公式ドキュメント](https://developers.google.com/apps-script)
- [Google Gemini API リファレンス](https://ai.google.dev/docs)
- [clasp CLI](https://github.com/google/clasp)
