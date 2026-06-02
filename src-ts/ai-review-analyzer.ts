/**
 * @fileoverview AI顧客レビュー自動分析（TypeScript版）
 * スプレッドシートのA列に入力された顧客レビューをGemini APIで分析し、
 * 感情分類（ポジティブ／ネガティブ／中立）と要約をB列・C列に自動書き込みする。
 *
 * GS版（ai-review-analyzer.gs）との違い:
 *   - Gemini API レスポンス構造をインターフェースで型定義
 *   - 戻り値・引数すべてに型注釈を付与
 *   - null の扱いを型レベルで明示（string | null）
 */

// ─── 定数 ───────────────────────────────────────────────

const ANALYZER_PROPS = {
  GEMINI_API_KEY: 'GEMINI_API_KEY',
} as const;

const GEMINI_CONFIG = {
  MODEL: 'gemini-2.0-flash',
  API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
  MAX_RETRIES: 3,
  RETRY_INTERVAL_MS: 1000,
} as const;

const COLUMNS = {
  REVIEW: 1,
  SENTIMENT: 2,
  SUMMARY: 3,
} as const;

const SENTIMENT_LABELS = {
  POSITIVE: 'ポジティブ',
  NEGATIVE: 'ネガティブ',
  NEUTRAL: '中立',
} as const;

// ─── インターフェース ─────────────────────────────────────

/** analyzeReview の戻り値 */
export interface AnalysisResult {
  sentiment: string;
  summary: string;
}

/** Gemini API レスポンスの型 */
export interface GeminiResponse {
  candidates: Array<{
    content: {
      parts: Array<{ text: string }>;
    };
  }>;
}

/** Gemini API リクエストボディの型 */
interface GeminiRequestBody {
  contents: Array<{ parts: Array<{ text: string }> }>;
  generationConfig: {
    temperature: number;
    maxOutputTokens: number;
  };
}

// ─── メイン処理 ──────────────────────────────────────────

/**
 * A列の全レビューを分析してB列・C列に結果を書き込む。
 */
export function analyzeAllReviews(): void {
  const apiKey: string | null = _getApiKey();
  if (!apiKey) return;

  const sheet: GoogleAppsScript.Spreadsheet.Sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow: number = sheet.getLastRow();

  if (lastRow < 2) {
    console.log('[ai-review-analyzer] 分析対象のデータがありません（2行目以降にレビューを入力してください）。');
    return;
  }

  const reviews: unknown[] = sheet
    .getRange(2, COLUMNS.REVIEW, lastRow - 1, 1)
    .getValues()
    .flat();

  let successCount = 0;
  let errorCount = 0;

  reviews.forEach((review: unknown, index: number) => {
    const rowNumber: number = index + 2;
    const reviewText = String(review ?? '').trim();

    if (!reviewText) return;

    const existingSentiment: unknown = sheet
      .getRange(rowNumber, COLUMNS.SENTIMENT)
      .getValue();
    if (existingSentiment !== '') {
      console.log(`[ai-review-analyzer] 行${rowNumber}: 分析済みのためスキップ`);
      return;
    }

    const result: AnalysisResult | null = _analyzeReview(apiKey, reviewText);

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

    if (index < reviews.length - 1) {
      Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS);
    }
  });

  console.log(`[ai-review-analyzer] 分析完了: 成功=${successCount}件, エラー=${errorCount}件`);
}

/**
 * アクティブな行のレビューのみを分析する（単一行処理用）。
 */
export function analyzeCurrentRow(): void {
  const apiKey: string | null = _getApiKey();
  if (!apiKey) return;

  const sheet: GoogleAppsScript.Spreadsheet.Sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row: number = sheet.getActiveCell().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('2行目以降のレビューセルを選択してください。');
    return;
  }

  const review: unknown = sheet.getRange(row, COLUMNS.REVIEW).getValue();
  if (!review) {
    SpreadsheetApp.getUi().alert('A列にレビューが入力されていません。');
    return;
  }

  const result: AnalysisResult | null = _analyzeReview(apiKey, String(review));
  if (result) {
    sheet.getRange(row, COLUMNS.SENTIMENT).setValue(result.sentiment);
    sheet.getRange(row, COLUMNS.SUMMARY).setValue(result.summary);
  }
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * スクリプトプロパティからGemini APIキーを取得する。
 * @returns APIキー。未設定の場合は null
 */
function _getApiKey(): string | null {
  const apiKey: string | null = PropertiesService.getScriptProperties().getProperty(
    ANALYZER_PROPS.GEMINI_API_KEY
  );
  if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY') {
    console.error(
      '[ai-review-analyzer] GEMINI_API_KEY が未設定です。initializeSettings() を実行後、APIキーを設定してください。'
    );
    return null;
  }
  return apiKey;
}

/**
 * Gemini APIを呼び出してレビューを分析する。
 * @param apiKey Gemini APIキー
 * @param reviewText 分析対象のレビューテキスト
 * @returns 分析結果。エラー時は null
 */
function _analyzeReview(
  apiKey: string,
  reviewText: string
): AnalysisResult | null {
  const prompt: string = _buildPrompt(reviewText);
  const endpoint = `${GEMINI_CONFIG.API_BASE_URL}/${GEMINI_CONFIG.MODEL}:generateContent?key=${apiKey}`;

  const requestBody: GeminiRequestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 256,
    },
  };

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
  };

  for (let attempt = 1; attempt <= GEMINI_CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response: GoogleAppsScript.URL_Fetch.HTTPResponse =
        UrlFetchApp.fetch(endpoint, options);
      const statusCode: number = response.getResponseCode();

      if (statusCode !== 200) {
        console.error(
          `[ai-review-analyzer] API呼び出し失敗 (試行${attempt}/${GEMINI_CONFIG.MAX_RETRIES}): ステータス=${statusCode}`
        );
        if (attempt < GEMINI_CONFIG.MAX_RETRIES) {
          Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS * attempt);
        }
        continue;
      }

      const json: GeminiResponse = JSON.parse(response.getContentText()) as GeminiResponse;
      return _parseGeminiResponse(json);
    } catch (e) {
      console.error(
        `[ai-review-analyzer] APIリクエストでエラーが発生しました (試行${attempt}/${GEMINI_CONFIG.MAX_RETRIES}): ${(e as Error).message}`
      );
      if (attempt < GEMINI_CONFIG.MAX_RETRIES) {
        Utilities.sleep(GEMINI_CONFIG.RETRY_INTERVAL_MS * attempt);
      }
    }
  }

  return null;
}

/**
 * Gemini APIに送信するプロンプトを生成する。
 * @param reviewText 分析対象のレビューテキスト
 * @returns プロンプト文字列
 */
export function _buildPrompt(reviewText: string): string {
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
 * @param json Gemini APIのレスポンスJSON
 * @returns 分析結果。パース失敗時は null
 */
export function _parseGeminiResponse(json: GeminiResponse): AnalysisResult | null {
  try {
    const rawText: string = json.candidates[0].content.parts[0].text.trim();

    const jsonMatch: RegExpMatchArray | null = rawText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.error(`[ai-review-analyzer] JSONの抽出に失敗しました。レスポンス: ${rawText}`);
      return null;
    }

    const parsed = JSON.parse(jsonMatch[0]) as { sentiment: string; summary?: string };
    const validSentiments: string[] = Object.values(SENTIMENT_LABELS);

    if (!validSentiments.includes(parsed.sentiment)) {
      console.error(`[ai-review-analyzer] 無効な感情分類値: "${parsed.sentiment}"`);
      return null;
    }

    return {
      sentiment: parsed.sentiment,
      summary: parsed.summary ?? '',
    };
  } catch (e) {
    console.error(`[ai-review-analyzer] レスポンスのパースに失敗しました: ${(e as Error).message}`);
    return null;
  }
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。初回セットアップ時に一度だけ実行する。
 */
export function initializeSettings(): void {
  const props: GoogleAppsScript.Properties.Properties =
    PropertiesService.getScriptProperties();
  props.setProperties({
    [ANALYZER_PROPS.GEMINI_API_KEY]: 'YOUR_GEMINI_API_KEY',
  });

  const sheet: GoogleAppsScript.Spreadsheet.Sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, COLUMNS.REVIEW).setValue('顧客レビュー（A列に入力）');
  sheet.getRange(1, COLUMNS.SENTIMENT).setValue('感情分類');
  sheet.getRange(1, COLUMNS.SUMMARY).setValue('要約');

  console.log('[ai-review-analyzer] 初期設定が完了しました。');
}
