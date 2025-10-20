/**
 * 共通関数
 * GPT API呼び出し、ログ出力、JSON解析などの共通処理
 */

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

/* ===== ログ出力（7分割：N列〜P列、12星座：R列〜T列、ランキング：Y列〜AA列） ===== */
function addLog(sheet, stepName, request, response, startTime, endTime) {
  const duration = ((endTime - startTime) / 1000).toFixed(2);

  const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestSummary = `[${stepName}]\n実行時間: ${duration}秒\n\nプロンプト:\n${request.substring(0, 500)}${request.length > 500 ? '...' : ''}`;
  const responseSummary = `レスポンス:\n${response.substring(0, 500)}${response.length > 500 ? '...' : ''}`;

  // 機能ごとにログ列を変更（ランキング：Y列（25列目）、12星座：R列（18列目）、7分割：N列（14列目））
  const logColumn = stepName.includes('ランキング') ? 25 : stepName.includes('12星座') ? 18 : 14;

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
