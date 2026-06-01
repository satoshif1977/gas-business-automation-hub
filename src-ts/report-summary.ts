/**
 * @fileoverview 週次レポートサマリー送信（TypeScript 並置実装）
 * GAS版 report-summary.gs の純粋関数部分を TypeScript で実装。
 * Jest によるユニットテスト対象。
 */

// ─── 型定義 ─────────────────────────────────────────────

export interface TaskRow {
  date: string;
  taskName: string;
  assignee: string;
  status: string;
}

export interface AssigneeStat {
  total: number;
  done: number;
}

export interface ReportSummary {
  total: number;
  done: number;
  inProgress: number;
  notStarted: number;
  byAssignee: Record<string, AssigneeStat>;
}

export interface EmailContent {
  subject: string;
  body: string;
}

export const STATUS = {
  DONE: '完了',
  IN_PROGRESS: '進行中',
  NOT_STARTED: '未着手',
} as const;

// ─── ビジネスロジック ──────────────────────────────────

/**
 * タスク行の配列からサマリー統計を生成する。
 */
export function buildSummary(rows: TaskRow[]): ReportSummary {
  const summary: ReportSummary = {
    total: 0,
    done: 0,
    inProgress: 0,
    notStarted: 0,
    byAssignee: {},
  };

  for (const row of rows) {
    if (!row.taskName) continue;
    summary.total++;

    if (row.status === STATUS.DONE) summary.done++;
    else if (row.status === STATUS.IN_PROGRESS) summary.inProgress++;
    else summary.notStarted++;

    const key = row.assignee || '未割当';
    if (!summary.byAssignee[key]) {
      summary.byAssignee[key] = { total: 0, done: 0 };
    }
    summary.byAssignee[key].total++;
    if (row.status === STATUS.DONE) summary.byAssignee[key].done++;
  }

  return summary;
}

/**
 * 完了率を計算する（0〜100 の整数）。
 */
export function calcCompletionRate(done: number, total: number): number {
  if (total === 0) return 0;
  return Math.round((done / total) * 100);
}

/**
 * メールの件名と本文を生成する。
 * @param sheetName シート名
 * @param summary buildSummary の戻り値
 * @param today 基準日（テストで差し替え可能）
 */
export function buildEmailContent(
  sheetName: string,
  summary: ReportSummary,
  today: Date = new Date(),
): EmailContent {
  const dateStr = today.toLocaleDateString('ja-JP', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  });
  const rate = calcCompletionRate(summary.done, summary.total);

  const subject = `[週次レポート] ${sheetName} — ${dateStr} (完了率 ${rate}%)`;

  const lines: string[] = [
    `■ 週次タスクサマリー（${dateStr}）`,
    `シート: ${sheetName}`,
    '',
    '【全体集計】',
    `  合計   : ${summary.total} 件`,
    `  完了   : ${summary.done} 件`,
    `  進行中 : ${summary.inProgress} 件`,
    `  未着手 : ${summary.notStarted} 件`,
    `  完了率 : ${rate}%`,
    '',
    '【担当者別完了状況】',
  ];

  for (const [name, stat] of Object.entries(summary.byAssignee)) {
    const r = calcCompletionRate(stat.done, stat.total);
    lines.push(`  ${name}: ${stat.done}/${stat.total} 件完了 (${r}%)`);
  }

  lines.push('', '-- 自動送信: gas-business-automation-hub / report-summary.gs');

  return { subject, body: lines.join('\n') };
}
