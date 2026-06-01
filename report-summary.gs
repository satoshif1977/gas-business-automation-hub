/**
 * @fileoverview 週次レポートサマリー送信
 * 週次管理スプレッドシートの各シートを集計し、
 * 担当者へサマリーレポートをメールで自動送信するスクリプト。
 *
 * 【スプレッドシート列構成（2行目以降がデータ行）】
 *   A列: 日付    B列: タスク名    C列: 担当者    D列: ステータス
 *   ステータス値: "完了" / "進行中" / "未着手"
 *
 * 【設定方法】
 * 1. スクリプトエディタを開く（拡張機能 → Apps Script）
 * 2. initializeSettings() を実行してスクリプトプロパティを初期化する
 * 3. スクリプトプロパティの値を実際のメールアドレスに書き換える
 * 4. トリガーを設定する（時間主導型 → 週ベース → 毎週月曜 午前8時）
 */

// ─── 定数 ───────────────────────────────────────────────

const REPORT_PROPS = {
  TO_EMAIL: 'REPORT_TO_EMAIL',
  SHEET_NAME: 'REPORT_SHEET_NAME',
};

const STATUS = {
  DONE: '完了',
  IN_PROGRESS: '進行中',
  NOT_STARTED: '未着手',
};

// ─── メイン処理 ──────────────────────────────────────────

/**
 * 週次サマリーレポートのメイン関数。
 * 時間主導型トリガー（毎週月曜8:00）から呼び出す。
 * @return {void}
 */
function sendWeeklyReport() {
  const props = PropertiesService.getScriptProperties();
  const toEmail = props.getProperty(REPORT_PROPS.TO_EMAIL);

  if (!toEmail) {
    console.error('[report-summary] REPORT_TO_EMAIL が未設定です。initializeSettings() を実行してください。');
    return;
  }

  const sheetName = props.getProperty(REPORT_PROPS.SHEET_NAME) || null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();

  if (!sheet) {
    console.error(`[report-summary] シート "${sheetName}" が見つかりません。`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('[report-summary] データ行がありません。メール送信をスキップします。');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const summary = _buildSummary(data);
  const { subject, body } = _buildEmailContent(sheet.getName(), summary);

  MailApp.sendEmail({ to: toEmail, subject, body });
  console.log(`[report-summary] レポートを ${toEmail} に送信しました。`);
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * スプレッドシートデータからサマリー統計を生成する。
 * @param {any[][]} data 2次元配列（[日付, タスク名, 担当者, ステータス][]）
 * @return {{ total: number, done: number, inProgress: number, notStarted: number, byAssignee: Object }}
 */
function _buildSummary(data) {
  const summary = {
    total: 0,
    done: 0,
    inProgress: 0,
    notStarted: 0,
    byAssignee: {},
  };

  data.forEach(([, taskName, assignee, status]) => {
    if (!taskName) return;
    summary.total++;

    if (status === STATUS.DONE) summary.done++;
    else if (status === STATUS.IN_PROGRESS) summary.inProgress++;
    else summary.notStarted++;

    const key = assignee || '未割当';
    if (!summary.byAssignee[key]) {
      summary.byAssignee[key] = { total: 0, done: 0 };
    }
    summary.byAssignee[key].total++;
    if (status === STATUS.DONE) summary.byAssignee[key].done++;
  });

  return summary;
}

/**
 * メールの件名と本文を生成する。
 * @param {string} sheetName シート名
 * @param {Object} summary _buildSummary の戻り値
 * @return {{ subject: string, body: string }}
 */
function _buildEmailContent(sheetName, summary) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const completionRate = summary.total > 0
    ? Math.round((summary.done / summary.total) * 100)
    : 0;

  const subject = `[週次レポート] ${sheetName} — ${today} (完了率 ${completionRate}%)`;

  const lines = [
    `■ 週次タスクサマリー（${today}）`,
    `シート: ${sheetName}`,
    '',
    '【全体集計】',
    `  合計   : ${summary.total} 件`,
    `  完了   : ${summary.done} 件`,
    `  進行中 : ${summary.inProgress} 件`,
    `  未着手 : ${summary.notStarted} 件`,
    `  完了率 : ${completionRate}%`,
    '',
    '【担当者別完了状況】',
  ];

  Object.entries(summary.byAssignee).forEach(([name, stat]) => {
    const rate = stat.total > 0 ? Math.round((stat.done / stat.total) * 100) : 0;
    lines.push(`  ${name}: ${stat.done}/${stat.total} 件完了 (${rate}%)`);
  });

  lines.push('', '-- 自動送信: gas-business-automation-hub / report-summary.gs');

  return { subject, body: lines.join('\n') };
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * スクリプトプロパティの初期値を設定する。
 * 初回セットアップ時に一度だけ実行する。
 * @return {void}
 */
function initializeReportSettings() {
  const props = PropertiesService.getScriptProperties();

  props.setProperties({
    [REPORT_PROPS.TO_EMAIL]: 'your-email@example.com',
    [REPORT_PROPS.SHEET_NAME]: '週次タスク',
  });

  console.log('[report-summary] スクリプトプロパティを初期化しました。');
  console.log('以下の値を実際の設定に書き換えてください:');
  console.log(`  ${REPORT_PROPS.TO_EMAIL}: 送信先メールアドレス`);
  console.log(`  ${REPORT_PROPS.SHEET_NAME}: 対象シート名`);
}
