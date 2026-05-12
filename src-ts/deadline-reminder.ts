/**
 * @fileoverview 期限管理リマインダー（TypeScript版）
 * スプレッドシートのB列に入力された期限日をチェックし、
 * 指定日前になるとSlackへ自動通知するスクリプト。
 *
 * GS版（deadline-reminder.gs）との違い:
 *   - 全関数・変数に型注釈を付与
 *   - GAS固有API（SpreadsheetApp等）に @types/google-apps-script の型を適用
 *   - インターフェースで戻り値の構造を明示
 */

// ─── 定数 ───────────────────────────────────────────────

/** スクリプトプロパティのキー名 */
const DEADLINE_PROPS = {
  SLACK_WEBHOOK_URL: 'DEADLINE_SLACK_WEBHOOK_URL',
  NOTIFY_DAYS: 'DEADLINE_NOTIFY_DAYS',
} as const;

/** デフォルトの通知日数（期限の何日前に通知するか） */
const DEFAULT_NOTIFY_DAYS: number[] = [10, 3, 0];

// ─── インターフェース ─────────────────────────────────────

/** Slack 通知ペイロード */
interface SlackPayload {
  text: string;
}

/** 期限チェック対象の行データ */
interface DeadlineRow {
  rowNumber: number;
  taskName: string;
  deadline: Date;
  daysUntilDeadline: number;
}

// ─── メイン処理 ──────────────────────────────────────────

/**
 * 期限チェックのメイン関数。
 * 時間主導型トリガー（毎日9:00）から呼び出す。
 */
function checkDeadlines(): void {
  const props: GoogleAppsScript.Properties.Properties =
    PropertiesService.getScriptProperties();
  const webhookUrl: string | null = props.getProperty(
    DEADLINE_PROPS.SLACK_WEBHOOK_URL
  );

  if (!webhookUrl) {
    console.error(
      '[deadline-reminder] SLACK_WEBHOOK_URL が未設定です。initializeSettings() を実行してください。'
    );
    return;
  }

  const notifyDays: number[] = _getNotifyDays(props);
  const sheet: GoogleAppsScript.Spreadsheet.Sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow: number = sheet.getLastRow();

  if (lastRow < 2) {
    console.log('[deadline-reminder] チェック対象のデータがありません。');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rows: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(2, 1, lastRow - 1, 2);
  const values: unknown[][] = rows.getValues();
  let notifiedCount = 0;

  values.forEach((row: unknown[], index: number) => {
    const taskName = String(row[0] ?? '');
    const deadlineRaw = row[1];

    if (!taskName || !deadlineRaw || !(deadlineRaw instanceof Date)) return;

    const deadline = new Date(deadlineRaw);
    deadline.setHours(0, 0, 0, 0);
    const diffMs: number = deadline.getTime() - today.getTime();
    const daysUntil: number = Math.round(diffMs / (1000 * 60 * 60 * 24));

    if (!notifyDays.includes(daysUntil)) return;

    const deadlineRow: DeadlineRow = {
      rowNumber: index + 2,
      taskName,
      deadline,
      daysUntilDeadline: daysUntil,
    };

    _sendSlackNotification(webhookUrl, deadlineRow);
    notifiedCount++;
  });

  console.log(`[deadline-reminder] 通知完了: ${notifiedCount}件`);
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * スクリプトプロパティから通知日数リストを取得する。
 * @param props スクリプトプロパティ
 * @returns 通知する日数の配列
 */
function _getNotifyDays(
  props: GoogleAppsScript.Properties.Properties
): number[] {
  const raw: string | null = props.getProperty(DEADLINE_PROPS.NOTIFY_DAYS);
  if (!raw) return DEFAULT_NOTIFY_DAYS;

  const parsed: number[] = raw
    .split(',')
    .map((s: string) => parseInt(s.trim(), 10))
    .filter((n: number) => !isNaN(n));

  return parsed.length > 0 ? parsed : DEFAULT_NOTIFY_DAYS;
}

/**
 * Slack Webhook で期限通知を送信する。
 * @param webhookUrl Slack Incoming Webhook URL
 * @param row 通知対象の行データ
 */
function _sendSlackNotification(
  webhookUrl: string,
  row: DeadlineRow
): void {
  const daysLabel: string =
    row.daysUntilDeadline === 0
      ? '*本日が期限です！*'
      : `あと *${row.daysUntilDeadline}日* で期限です`;

  const payload: SlackPayload = {
    text: `[期限リマインダー] ${row.taskName} — ${daysLabel}（期限: ${row.deadline.toLocaleDateString('ja-JP')}）`,
  };

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response: GoogleAppsScript.URL_Fetch.HTTPResponse =
    UrlFetchApp.fetch(webhookUrl, options);

  if (response.getResponseCode() !== 200) {
    console.error(
      `[deadline-reminder] Slack 通知失敗: ステータス=${response.getResponseCode()}`
    );
  } else {
    console.log(`[deadline-reminder] Slack 通知送信: "${row.taskName}"`);
  }
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。
 * 初回セットアップ時に一度だけ実行する。
 */
function initializeSettings(): void {
  const props: GoogleAppsScript.Properties.Properties =
    PropertiesService.getScriptProperties();

  props.setProperties({
    [DEADLINE_PROPS.SLACK_WEBHOOK_URL]: 'YOUR_SLACK_WEBHOOK_URL',
    [DEADLINE_PROPS.NOTIFY_DAYS]: DEFAULT_NOTIFY_DAYS.join(','),
  });

  console.log('[deadline-reminder] 初期設定が完了しました。');
  console.log('次のステップ:');
  console.log('  1. スクリプトプロパティの DEADLINE_SLACK_WEBHOOK_URL に Webhook URL を設定する');
  console.log('  2. 時間主導型トリガー（毎日9:00）を設定する');
}
