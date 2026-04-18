/**
 * @fileoverview AI顧客レビュー自動分析
 * スプレッドシートのA列に入力された顧客レビューをGemini APIで分析し、
 * 感情分類（ポジティブ／ネガティブ／中立）と要約をB列・C列に自動書き込みする。
 *
 * 【設定方法】
 * 1. Google AI Studio（https://aistudio.google.com）でAPIキーを取得する
 * 2. スクリプトエディタを開く（拡張機能 → Apps Script）
 * 3. initializeSettings() を実行してスクリプトプロパティを初期化する
 * 4. スクリプトプロパティの GEMINI_API_KEY に取得したAPIキーを設定する
 * 5. analyzeAllReviews() を実行する（またはボタンにトリガーを設定する）
 */

// ─── 定数 ───────────────────────────────────────────────

/** スクリプトプロパティのキー名 */
const ANALYZER_PROPS = {
  GEMINI_API_KEY: 'GEMINI_API_KEY',
};

/** Gemini API設定 */
const GEMINI_CONFIG = {
  MODEL: 'gemini-1.5-flash',
  API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
  MAX_RETRIES: 3,
  RETRY_INTERVAL_MS: 1000,
};

/** スプレッドシートの列インデックス（1始まり） */
const COLUMNS = {
  REVIEW: 1,   // A列: 顧客レビュー（入力）
  SENTIMENT: 2, // B列: 感情分類（出力）
  SUMMARY: 3,   // C列: 要約（出力）
};

/** 感情分類の選択肢 */
const SENTIMENT_LABELS = {
  POSITIVE: 'ポジティブ',
  NEGATIVE: 'ネガティブ',
  NEUTRAL: '中立',
};

// ─── メイン処理 ──────────────────────────────────────────

/**
 * A列の全レビューを分析してB列・C列に結果を書き込む。
 * スプレッドシートのボタンまたは手動実行から呼び出す。
 * @return {void}
 */
function analyzeAllReviews() {
  const apiKey = _getApiKey();
  if (!apiKey) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    console.log('[ai-review-analyzer] 分析対象のデータがありません（2行目以降にレビューを入力してください）。');
    return;
  }

  const reviews = sheet.getRange(2, COLUMNS.REVIEW, lastRow - 1, 1).getValues().flat();
  let successCount = 0;
  let errorCount = 0;

  reviews.forEach((review, index) => {
    const rowNumber = index + 2;

    // 空セルはスキップ
    if (!review || String(review).trim() === '') return;

    // 既に分析済みの行はスキップ（B列に値がある場合）
    const existingSentiment = sheet.getRange(rowNumber, COLUMNS.SENTIMENT).getValue();
    if (existingSentiment !== '') {
      console.log(`[ai-review-analyzer] 行${rowNumber}: 分析済みのためスキップ`);
      return;
    }

    const result = _analyzeReview(apiKey, String(review));

    if (result) {
      sheet.getRange(rowNumber, COLUMNS.SENTIMENT).setValue(result.sentiment);
      sheet.getRange(rowNumber, COLUMNS.SUMMARY).setValue(result.summary);
      successCount++;
      console.log(`[ai-review-analyzer] 行${rowNumber}: 分析完了 → ${result.sentiment}`);
    } else {
      sheet.getRange(rowNumber, COLUMNS.SENTIMENT).setValue('分析エラー');
      errorCount++;
      console.error(`[ai-review-analyzer] 行${rowNumber}: 分析失敗`);
    }

    // API レート制限対策（1秒待機）
    if (index < reviews.length - 1) {
      Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS);
    }
  });

  console.log(`[ai-review-analyzer] 分析完了: 成功=${successCount}件, エラー=${errorCount}件`);
}

/**
 * アクティブな行のレビューのみを分析する（単一行処理用）。
 * @return {void}
 */
function analyzeCurrentRow() {
  const apiKey = _getApiKey();
  if (!apiKey) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('2行目以降のレビューセルを選択してください。');
    return;
  }

  const review = sheet.getRange(row, COLUMNS.REVIEW).getValue();
  if (!review) {
    SpreadsheetApp.getUi().alert('A列にレビューが入力されていません。');
    return;
  }

  const result = _analyzeReview(apiKey, String(review));
  if (result) {
    sheet.getRange(row, COLUMNS.SENTIMENT).setValue(result.sentiment);
    sheet.getRange(row, COLUMNS.SUMMARY).setValue(result.summary);
  }
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * スクリプトプロパティからGemini APIキーを取得する。
 * @return {string|null} APIキー。未設定の場合はnullを返す。
 */
function _getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(ANALYZER_PROPS.GEMINI_API_KEY);
  if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY') {
    console.error('[ai-review-analyzer] GEMINI_API_KEY が未設定です。initializeSettings() を実行後、APIキーを設定してください。');
    return null;
  }
  return apiKey;
}

/**
 * Gemini APIを呼び出してレビューを分析する。
 * @param {string} apiKey Gemini APIキー
 * @param {string} reviewText 分析対象のレビューテキスト
 * @return {{sentiment: string, summary: string}|null} 分析結果。エラー時はnullを返す。
 */
function _analyzeReview(apiKey, reviewText) {
  const prompt = _buildPrompt(reviewText);
  const endpoint = `${GEMINI_CONFIG.API_BASE_URL}/${GEMINI_CONFIG.MODEL}:generateContent?key=${apiKey}`;

  const requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1, // 分類タスクのため低めに設定
      maxOutputTokens: 256,
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
  };

  for (let attempt = 1; attempt <= GEMINI_CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(endpoint, options);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        console.error(`[ai-review-analyzer] API呼び出し失敗 (試行${attempt}/${GEMINI_CONFIG.MAX_RETRIES}): ステータス=${statusCode}`);
        if (attempt < GEMINI_CONFIG.MAX_RETRIES) Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS * attempt);
        continue;
      }

      const json = JSON.parse(response.getContentText());
      return _parseGeminiResponse(json);

    } catch (e) {
      console.error(`[ai-review-analyzer] APIリクエストでエラーが発生しました (試行${attempt}/${GEMINI_CONFIG.MAX_RETRIES}): ${e.message}`);
      if (attempt < GEMINI_CONFIG.MAX_RETRIES) Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS * attempt);
    }
  }

  return null;
}

/**
 * Gemini APIに送信するプロンプトを生成する。
 * @param {string} reviewText 分析対象のレビューテキスト
 * @return {string} プロンプト文字列
 */
function _buildPrompt(reviewText) {
  return `以下の顧客レビューを分析してください。

【レビュー】
${reviewText}

【出力形式】
以下のJSON形式で回答してください。他の文字列は含めないこと。

{
  "sentiment": "ポジティブ" または "ネガティブ" または "中立",
  "summary": "レビューの要約を30文字以内で記述"
}`;
}

/**
 * Gemini APIのレスポンスをパースして分析結果を返す。
 * @param {Object} json Gemini APIのレスポンスJSON
 * @return {{sentiment: string, summary: string}|null} 分析結果。パース失敗時はnullを返す。
 */
function _parseGeminiResponse(json) {
  try {
    const rawText = json.candidates[0].content.parts[0].text.trim();

    // JSONブロックの抽出（```json ... ``` 形式にも対応）
    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.error(`[ai-review-analyzer] JSONの抽出に失敗しました。レスポンス: ${rawText}`);
      return null;
    }

    const parsed = JSON.parse(jsonMatch[0]);
    const validSentiments = Object.values(SENTIMENT_LABELS);

    if (!validSentiments.includes(parsed.sentiment)) {
      console.error(`[ai-review-analyzer] 無効な感情分類値: "${parsed.sentiment}"`);
      return null;
    }

    return {
      sentiment: parsed.sentiment,
      summary: parsed.summary || '',
    };
  } catch (e) {
    console.error(`[ai-review-analyzer] レスポンスのパースに失敗しました: ${e.message}`);
    return null;
  }
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。
 * 初回セットアップ時に一度だけ実行する。
 * @return {void}
 */
function initializeSettings() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [ANALYZER_PROPS.GEMINI_API_KEY]: 'YOUR_GEMINI_API_KEY',
  });

  // ヘッダー行のセットアップ
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, COLUMNS.REVIEW).setValue('顧客レビュー（A列に入力）');
  sheet.getRange(1, COLUMNS.SENTIMENT).setValue('感情分類');
  sheet.getRange(1, COLUMNS.SUMMARY).setValue('要約');

  console.log('[ai-review-analyzer] 初期設定が完了しました。');
  console.log('次のステップ:');
  console.log('  1. Google AI Studio（https://aistudio.google.com）でAPIキーを取得する');
  console.log('  2. スクリプトプロパティの GEMINI_API_KEY にAPIキーを設定する');
  console.log('  3. A列にレビューを入力して analyzeAllReviews() を実行する');
}
