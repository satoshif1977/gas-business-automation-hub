/**
 * @fileoverview 期限管理リマインダー
 * スプレッドシートのB列に入力された期限日をチェックし、
 * 指定日前になるとSlackへ自動通知するスクリプト。
 *
 * 【設定方法】
 * 1. スクリプトエディタを開く（拡張機能 → Apps Script）
 * 2. initializeSettings() を実行してスクリプトプロパティを初期化する
 * 3. スクリプトプロパティの値を実際のSlack Webhook URLに書き換える
 * 4. トリガーを設定する（時間主導型 → 毎日 → 午前9時）
 */

// ─── 定数 ───────────────────────────────────────────────

/** スクリプトプロパティのキー名 */
const DEADLINE_PROPS = {
  SLACK_WEBHOOK_URL: 'DEADLINE_SLACK_WEBHOOK_URL',
  NOTIFY_DAYS: 'DEADLINE_NOTIFY_DAYS', // カンマ区切り例: "10,3,0"
};

/** デフォルトの通知日数（期限の何日前に通知するか） */
const DEFAULT_NOTIFY_DAYS = [10, 3, 0];

// ─── メイン処理 ──────────────────────────────────────────

/**
 * 期限チェックのメイン関数。
 * 時間主導型トリガー（毎日9:00）から呼び出す。
 * @return {void}
 */
function checkDeadlines() {
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty(DEADLINE_PROPS.SLACK_WEBHOOK_URL);

  if (!webhookUrl) {
    console.error('[deadline-reminder] SLACK_WEBHOOK_URL が未設定です。initializeSettings() を実行してください。');
    return;
  }

  const notifyDays = _getNotifyDays(props);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('[deadline-reminder] データ行がありません。');
    return;
  }

  // A列: 会社名, B列: 期限日
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  data.forEach(([companyName, deadlineDate]) => {
    if (!companyName || !deadlineDate) return;

    const deadline = new Date(deadlineDate);
    deadline.setHours(0, 0, 0, 0);

    const diffDays = Math.round((deadline - today) / (1000 * 60 * 60 * 24));

    if (notifyDays.includes(diffDays)) {
      const message = _buildSlackMessage(companyName, deadline, diffDays);
      _sendSlackNotification(webhookUrl, message);
    }
  });
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * スクリプトプロパティから通知日数リストを取得する。
 * プロパティが未設定の場合はデフォルト値を返す。
 * @param {GoogleAppsScript.Properties.Properties} props スクリプトプロパティ
 * @return {number[]} 通知日数の配列（例: [10, 3, 0]）
 */
function _getNotifyDays(props) {
  const raw = props.getProperty(DEADLINE_PROPS.NOTIFY_DAYS);
  if (!raw) return DEFAULT_NOTIFY_DAYS;
  return raw.split(',').map((d) => parseInt(d.trim(), 10)).filter((d) => !isNaN(d));
}

/**
 * Slackに送信するメッセージ文字列を生成する。
 * @param {string} companyName 会社名
 * @param {Date} deadline 期限日
 * @param {number} diffDays 今日からの残り日数
 * @return {string} Slack通知メッセージ
 */
function _buildSlackMessage(companyName, deadline, diffDays) {
  const deadlineStr = Utilities.formatDate(deadline, 'Asia/Tokyo', 'yyyy/MM/dd');

  if (diffDays === 0) {
    return `🚨 *【本日期限】* ${companyName} の期限は *本日（${deadlineStr}）* です！`;
  }
  return `⏰ *【期限アラート】* ${companyName} の期限まで *あと${diffDays}日*（${deadlineStr}）です。`;
}

/**
 * Slack Incoming Webhookへメッセージを送信する。
 * @param {string} webhookUrl Slack Webhook URL
 * @param {string} message 送信するメッセージ
 * @return {void}
 */
function _sendSlackNotification(webhookUrl, message) {
  const payload = JSON.stringify({ text: message });
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    if (response.getResponseCode() !== 200) {
      console.error(`[deadline-reminder] Slack通知に失敗しました。レスポンスコード: ${response.getResponseCode()}`);
    } else {
      console.log(`[deadline-reminder] Slack通知成功: ${message}`);
    }
  } catch (e) {
    console.error(`[deadline-reminder] Slack通知でエラーが発生しました: ${e.message}`);
  }
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。
 * 初回セットアップ時に一度だけ実行する。
 * 実行後、スクリプトプロパティの値を実際の値に書き換えること。
 * @return {void}
 */
function initializeSettings() {
  const props = PropertiesService.getScriptProperties();

  props.setProperties({
    [DEADLINE_PROPS.SLACK_WEBHOOK_URL]: 'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
    [DEADLINE_PROPS.NOTIFY_DAYS]: '10,3,0',
  });

  console.log('[deadline-reminder] スクリプトプロパティを初期化しました。');
  console.log('以下の値を実際の設定に書き換えてください:');
  console.log(`  ${DEADLINE_PROPS.SLACK_WEBHOOK_URL}: Slack Incoming Webhook URL`);
  console.log(`  ${DEADLINE_PROPS.NOTIFY_DAYS}: 通知タイミング（カンマ区切りの日数）例: "10,3,0"`);
}
