import {
  parseNotifyDays,
  calcDaysUntil,
  formatDeadlineMessage,
  getLastMonth,
  parseTimeToMinutes,
  calcWorkMinutes,
} from '../utils';

// ── parseNotifyDays ─────────────────────────────────────────

describe('parseNotifyDays', () => {
  const defaults = [10, 3, 0];

  test('null の場合はデフォルト値を返す', () => {
    expect(parseNotifyDays(null, defaults)).toEqual(defaults);
  });

  test('空文字の場合はデフォルト値を返す', () => {
    expect(parseNotifyDays('', defaults)).toEqual(defaults);
  });

  test('カンマ区切り文字列を正しくパースする', () => {
    expect(parseNotifyDays('10,3,0', defaults)).toEqual([10, 3, 0]);
  });

  test('スペースを含むカンマ区切りも正しくパースする', () => {
    expect(parseNotifyDays('10, 3, 0', defaults)).toEqual([10, 3, 0]);
  });

  test('数値以外を含む場合は除外する', () => {
    expect(parseNotifyDays('10,abc,3', defaults)).toEqual([10, 3]);
  });

  test('全て無効な場合はデフォルト値を返す', () => {
    expect(parseNotifyDays('abc,xyz', defaults)).toEqual(defaults);
  });

  test('単一の数値も正しくパースする', () => {
    expect(parseNotifyDays('7', defaults)).toEqual([7]);
  });
});

// ── calcDaysUntil ──────────────────────────────────────────

describe('calcDaysUntil', () => {
  test('3日後を正しく計算する', () => {
    const today = new Date('2026-05-15');
    const deadline = new Date('2026-05-18');
    expect(calcDaysUntil(today, deadline)).toBe(3);
  });

  test('当日（0日）を正しく判定する', () => {
    const today = new Date('2026-05-15');
    const deadline = new Date('2026-05-15');
    expect(calcDaysUntil(today, deadline)).toBe(0);
  });

  test('期限超過（負の値）を正しく計算する', () => {
    const today = new Date('2026-05-15');
    const deadline = new Date('2026-05-10');
    expect(calcDaysUntil(today, deadline)).toBe(-5);
  });

  test('時刻に関わらず日付ベースで計算する', () => {
    const today = new Date('2026-05-15T23:59:59');
    const deadline = new Date('2026-05-16T00:00:00');
    expect(calcDaysUntil(today, deadline)).toBe(1);
  });

  test('30日後を正しく計算する', () => {
    const today = new Date('2026-05-01');
    const deadline = new Date('2026-05-31');
    expect(calcDaysUntil(today, deadline)).toBe(30);
  });
});

// ── formatDeadlineMessage ──────────────────────────────────

describe('formatDeadlineMessage', () => {
  const deadline = new Date('2026/05/20');

  test('当日（daysUntil=0）のメッセージを生成する', () => {
    const msg = formatDeadlineMessage('報告書提出', 0, deadline);
    expect(msg).toContain('本日が期限です');
    expect(msg).toContain('報告書提出');
  });

  test('残日数あり（daysUntil>0）のメッセージを生成する', () => {
    const msg = formatDeadlineMessage('月次レポート', 5, deadline);
    expect(msg).toContain('あと *5日* で期限です');
    expect(msg).toContain('月次レポート');
  });

  test('メッセージに期限日を含む', () => {
    const msg = formatDeadlineMessage('タスク', 3, deadline);
    expect(msg).toContain('期限:');
  });

  test('メッセージに [期限リマインダー] プレフィックスを含む', () => {
    const msg = formatDeadlineMessage('タスク', 1, deadline);
    expect(msg).toContain('[期限リマインダー]');
  });
});

// ── getLastMonth ───────────────────────────────────────────

describe('getLastMonth', () => {
  test('通常月（5月）の前月は4月', () => {
    const result = getLastMonth(new Date('2026-05-15'));
    expect(result).toEqual({ year: 2026, month: 4 });
  });

  test('1月の前月は前年12月', () => {
    const result = getLastMonth(new Date('2026-01-01'));
    expect(result).toEqual({ year: 2025, month: 12 });
  });

  test('12月の前月は11月', () => {
    const result = getLastMonth(new Date('2026-12-31'));
    expect(result).toEqual({ year: 2026, month: 11 });
  });

  test('月の途中の日付でも正しく前月を返す', () => {
    const result = getLastMonth(new Date('2026-03-20'));
    expect(result).toEqual({ year: 2026, month: 2 });
  });
});

// ── parseTimeToMinutes ─────────────────────────────────────

describe('parseTimeToMinutes', () => {
  test('"09:30" を 570分 に変換する', () => {
    expect(parseTimeToMinutes('09:30')).toBe(570);
  });

  test('"00:00" は 0分', () => {
    expect(parseTimeToMinutes('00:00')).toBe(0);
  });

  test('"17:00" は 1020分', () => {
    expect(parseTimeToMinutes('17:00')).toBe(1020);
  });

  test('不正な形式は NaN を返す', () => {
    expect(parseTimeToMinutes('abc')).toBeNaN();
  });
});

// ── calcWorkMinutes ────────────────────────────────────────

describe('calcWorkMinutes', () => {
  test('"09:00" から "18:00" は 540分（9時間）', () => {
    expect(calcWorkMinutes('09:00', '18:00')).toBe(540);
  });

  test('"09:30" から "17:30" は 480分（8時間）', () => {
    expect(calcWorkMinutes('09:30', '17:30')).toBe(480);
  });

  test('退勤が出勤より前の場合は 0 を返す', () => {
    expect(calcWorkMinutes('18:00', '09:00')).toBe(0);
  });

  test('同じ時刻は 0 を返す', () => {
    expect(calcWorkMinutes('09:00', '09:00')).toBe(0);
  });

  test('不正な入力は 0 を返す', () => {
    expect(calcWorkMinutes('invalid', '09:00')).toBe(0);
  });
});
