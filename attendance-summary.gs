/**
 * @fileoverview 勤怠集計・月次サマリー自動通知
 * スプレッドシートに記録された勤怠データを月次集計し、
 * Slack へサマリーレポートを自動送信するスクリプト。
 *
 * 【スプレッドシート形式】
 * A列: 日付 (YYYY/MM/DD)
 * B列: 氏名
 * C列: 出勤時刻 (HH:MM)
 * D列: 退勤時刻 (HH:MM)
 * E列: 備考（遅刻・早退・有休 等）
 *
 * 【設定方法】
 * 1. スクリプトエディタを開く（拡張機能 → Apps Script）
 * 2. initializeAttendanceSettings() を実行してプロパティを初期化する
 * 3. スクリプトプロパティの値を実際の Slack Webhook URL に書き換える
 * 4. トリガーを設定する（時間主導型 → 月ベース → 毎月1日 午前9時）
 */

// ─── 定数 ───────────────────────────────────────────────────

/** スクリプトプロパティのキー名 */
const ATTENDANCE_PROPS = {
  SLACK_WEBHOOK_URL: 'ATTENDANCE_SLACK_WEBHOOK_URL',
  SHEET_NAME: 'ATTENDANCE_SHEET_NAME',
  NOTIFY_CHANNEL: 'ATTENDANCE_NOTIFY_CHANNEL',
};

/** 勤怠区分の判定しきい値 */
const ATTENDANCE_THRESHOLDS = {
  LATE_MINUTE: 10,         // 定時から何分以上遅れたら遅刻とみなすか
  STANDARD_HOURS: 8,       // 所定労働時間（時間）
  OVERTIME_THRESHOLD: 8,   // 残業とみなす基準（時間）
};

// ─── メインエントリ ──────────────────────────────────────────

/**
 * 先月の勤怠データを集計して Slack に送信する。
 * 毎月1日のトリガーから呼び出す想定。
 */
function sendMonthlyAttendanceSummary() {
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty(ATTENDANCE_PROPS.SLACK_WEBHOOK_URL);
  const sheetName = props.getProperty(ATTENDANCE_PROPS.SHEET_NAME) || '勤怠';

  if (!webhookUrl) {
    Logger.log('[ERROR] ATTENDANCE_SLACK_WEBHOOK_URL が設定されていません。');
    return;
  }

  const lastMonth = getLastMonth_();
  const records = fetchAttendanceRecords_(sheetName, lastMonth.year, lastMonth.month);

  if (records.length === 0) {
    Logger.log(`[WARN] ${lastMonth.year}年${lastMonth.month}月の勤怠データが見つかりませんでした。`);
    return;
  }

  const summary = aggregateAttendance_(records);
  const message = buildSlackMessage_(summary, lastMonth.year, lastMonth.month);

  postToSlack_(webhookUrl, message);
  Logger.log(`[INFO] ${lastMonth.year}年${lastMonth.month}月の勤怠サマリーを送信しました。`);
}

// ─── データ取得 ─────────────────────────────────────────────

/**
 * 指定シートから対象年月の勤怠レコードを取得する。
 * @param {string} sheetName - シート名
 * @param {number} year - 対象年
 * @param {number} month - 対象月（1〜12）
 * @returns {Array<Object>} 勤怠レコードの配列
 */
function fetchAttendanceRecords_(sheetName, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`[ERROR] シート "${sheetName}" が見つかりません。`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const records = [];

  data.forEach((row) => {
    const dateCell = row[0];
    if (!dateCell) return;

    const date = new Date(dateCell);
    if (date.getFullYear() !== year || date.getMonth() + 1 !== month) return;

    const name = String(row[1]).trim();
    const clockIn = String(row[2]).trim();
    const clockOut = String(row[3]).trim();
    const note = String(row[4]).trim();

    if (!name) return;

    const workedHours = calcWorkedHours_(clockIn, clockOut);
    records.push({ date, name, clockIn, clockOut, note, workedHours });
  });

  return records;
}

// ─── 集計処理 ───────────────────────────────────────────────

/**
 * 勤怠レコードをメンバー別に集計する。
 * @param {Array<Object>} records - 勤怠レコードの配列
 * @returns {Object} メンバー別集計結果
 */
function aggregateAttendance_(records) {
  const summary = {};

  records.forEach((rec) => {
    if (!summary[rec.name]) {
      summary[rec.name] = {
        workDays: 0,
        totalHours: 0,
        overtimeDays: 0,
        leaveDays: 0,
        lateDays: 0,
      };
    }

    const s = summary[rec.name];

    // 有休・休暇の判定
    if (rec.note.includes('有休') || rec.note.includes('休暇')) {
      s.leaveDays += 1;
      return;
    }

    s.workDays += 1;
    s.totalHours += rec.workedHours;

    if (rec.workedHours > ATTENDANCE_THRESHOLDS.OVERTIME_THRESHOLD) {
      s.overtimeDays += 1;
    }
    if (rec.note.includes('遅刻')) {
      s.lateDays += 1;
    }
  });

  return summary;
}

// ─── Slack メッセージ生成 ────────────────────────────────────

/**
 * Slack 投稿用のメッセージオブジェクトを生成する。
 * @param {Object} summary - メンバー別集計結果
 * @param {number} year - 対象年
 * @param {number} month - 対象月
 * @returns {Object} Slack Incoming Webhook 用ペイロード
 */
function buildSlackMessage_(summary, year, month) {
  const members = Object.keys(summary);
  const lines = members.map((name) => {
    const s = summary[name];
    const avgHours = s.workDays > 0
      ? (s.totalHours / s.workDays).toFixed(1)
      : '0.0';
    return [
      `*${name}*`,
      `  出勤: ${s.workDays}日 / 有休: ${s.leaveDays}日`,
      `  総労働時間: ${s.totalHours.toFixed(1)}h（平均 ${avgHours}h/日）`,
      s.overtimeDays > 0 ? `  残業: ${s.overtimeDays}日` : null,
      s.lateDays > 0 ? `  遅刻: ${s.lateDays}回` : null,
    ].filter(Boolean).join('\n');
  });

  return {
    text: `📊 *${year}年${month}月 勤怠サマリー*`,
    blocks: [
      {
        type: 'header',
        text: {
          type: 'plain_text',
          text: `📊 ${year}年${month}月 勤怠サマリー`,
        },
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: lines.length > 0
            ? lines.join('\n\n')
            : '集計対象のレコードがありませんでした。',
        },
      },
      {
        type: 'context',
        elements: [
          {
            type: 'mrkdwn',
            text: `集計日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}`,
          },
        ],
      },
    ],
  };
}

// ─── ユーティリティ ─────────────────────────────────────────

/**
 * 出退勤時刻から労働時間（時間単位）を計算する。
 * @param {string} clockIn - 出勤時刻 "HH:MM"
 * @param {string} clockOut - 退勤時刻 "HH:MM"
 * @returns {number} 労働時間（小数）
 */
function calcWorkedHours_(clockIn, clockOut) {
  if (!clockIn || !clockOut) return 0;
  const toMinutes = (timeStr) => {
    const parts = timeStr.split(':');
    if (parts.length !== 2) return 0;
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  };
  const diffMinutes = toMinutes(clockOut) - toMinutes(clockIn);
  return diffMinutes > 0 ? diffMinutes / 60 : 0;
}

/**
 * 先月の年・月を返す。
 * @returns {{ year: number, month: number }}
 */
function getLastMonth_() {
  const now = new Date();
  const year = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const month = now.getMonth() === 0 ? 12 : now.getMonth();
  return { year, month };
}

/**
 * Slack の Incoming Webhook にメッセージを送信する。
 * @param {string} webhookUrl - Webhook URL
 * @param {Object} payload - 送信するメッセージオブジェクト
 */
function postToSlack_(webhookUrl, payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(webhookUrl, options);
  if (response.getResponseCode() !== 200) {
    Logger.log(`[ERROR] Slack 送信失敗: ${response.getContentText()}`);
  }
}

// ─── 初期設定 ───────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。
 * 初回セットアップ時に一度だけ実行する。
 */
function initializeAttendanceSettings() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [ATTENDANCE_PROPS.SLACK_WEBHOOK_URL]: 'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
    [ATTENDANCE_PROPS.SHEET_NAME]: '勤怠',
    [ATTENDANCE_PROPS.NOTIFY_CHANNEL]: '#勤怠',
  });
  Logger.log('[INFO] 勤怠集計スクリプトの初期設定が完了しました。');
  Logger.log('[INFO] ATTENDANCE_SLACK_WEBHOOK_URL を実際の Webhook URL に書き換えてください。');
}
