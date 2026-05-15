/**
 * @fileoverview 純粋ユーティリティ関数
 * GAS 固有 API に依存しない純粋関数をまとめたモジュール。
 * Jest によるユニットテストの対象。
 */

// ─── 期限管理ユーティリティ ────────────────────────────────

/**
 * 通知日数リストを文字列からパースする。
 * @param raw スクリプトプロパティの生文字列（例: "10,3,0"）
 * @param defaults パースに失敗した場合のデフォルト値
 */
export function parseNotifyDays(raw: string | null, defaults: number[]): number[] {
  if (!raw) return defaults;
  const parsed = raw
    .split(',')
    .map((s) => parseInt(s.trim(), 10))
    .filter((n) => !isNaN(n));
  return parsed.length > 0 ? parsed : defaults;
}

/**
 * 基準日から期限日までの残日数を計算する（時刻を除いた日付ベース）。
 * @param today 基準日
 * @param deadline 期限日
 */
export function calcDaysUntil(today: Date, deadline: Date): number {
  const t = new Date(today);
  const d = new Date(deadline);
  t.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((d.getTime() - t.getTime()) / (1000 * 60 * 60 * 24));
}

/**
 * Slack 向け期限通知メッセージを生成する。
 * @param taskName タスク名
 * @param daysUntil 残日数（0: 当日、負: 超過）
 * @param deadline 期限日
 */
export function formatDeadlineMessage(taskName: string, daysUntil: number, deadline: Date): string {
  const daysLabel =
    daysUntil === 0
      ? '*本日が期限です！*'
      : `あと *${daysUntil}日* で期限です`;
  const dateStr = deadline.toLocaleDateString('ja-JP');
  return `[期限リマインダー] ${taskName} — ${daysLabel}（期限: ${dateStr}）`;
}

// ─── 勤怠集計ユーティリティ ────────────────────────────────

/** 年月オブジェクト */
export interface YearMonth {
  year: number;
  month: number; // 1〜12
}

/**
 * 指定した日付の前月を返す。
 * @param now 基準日
 */
export function getLastMonth(now: Date): YearMonth {
  const d = new Date(now);
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return { year: d.getFullYear(), month: d.getMonth() + 1 };
}

/**
 * HH:MM 形式の文字列を分単位の数値に変換する。
 * @param timeStr "09:30" などの文字列
 * @returns 分数（パース失敗時は NaN）
 */
export function parseTimeToMinutes(timeStr: string): number {
  const [h, m] = timeStr.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return NaN;
  return h * 60 + m;
}

/**
 * 勤務時間を計算する（退勤 - 出勤、分単位）。
 * @param startStr 出勤時刻 "HH:MM"
 * @param endStr   退勤時刻 "HH:MM"
 * @returns 勤務分数（パース失敗時は 0）
 */
export function calcWorkMinutes(startStr: string, endStr: string): number {
  const start = parseTimeToMinutes(startStr);
  const end = parseTimeToMinutes(endStr);
  if (isNaN(start) || isNaN(end) || end <= start) return 0;
  return end - start;
}
