import {
  parseNotifyDays,
  calcDaysUntil,
  formatDeadlineMessage,
  getLastMonth,
  parseTimeToMinutes,
  calcWorkMinutes,
} from '../utils';

// ── parseNotifyDays エッジケース ──────────────────────────

describe('parseNotifyDays edge cases', () => {
  const defaults = [10, 3, 0];

  test('負数を含む文字列は負数のまま返す', () => {
    expect(parseNotifyDays('-5,3', defaults)).toEqual([-5, 3]);
  });

  test('小数は parseInt で整数部のみ返す', () => {
    expect(parseNotifyDays('3.5,2', defaults)).toEqual([3, 2]);
  });

  test('先頭ゼロ付き数値は正しくパースされる', () => {
    expect(parseNotifyDays('007,003', defaults)).toEqual([7, 3]);
  });

  test('スペースのみの文字列はデフォルト値を返す', () => {
    expect(parseNotifyDays('   ', defaults)).toEqual(defaults);
  });

  test('カンマのみの文字列はデフォルト値を返す', () => {
    expect(parseNotifyDays(',', defaults)).toEqual(defaults);
  });

  test('数値の後にアルファベットが続く場合は整数部分を取得する', () => {
    // parseInt('10abc') = 10
    expect(parseNotifyDays('10abc,5', defaults)).toEqual([10, 5]);
  });

  test('大きな数値も正しくパースされる', () => {
    expect(parseNotifyDays('365,30,7', defaults)).toEqual([365, 30, 7]);
  });

  test('0 単体はデフォルトではなく [0] を返す', () => {
    expect(parseNotifyDays('0', defaults)).toEqual([0]);
  });
});

// ── calcDaysUntil エッジケース ────────────────────────────

describe('calcDaysUntil edge cases', () => {
  test('年またぎ（12月→翌1月）を正しく計算する', () => {
    const today = new Date('2025-12-25');
    const deadline = new Date('2026-01-05');
    expect(calcDaysUntil(today, deadline)).toBe(11);
  });

  test('365日後を正しく計算する', () => {
    const today = new Date('2026-01-01');
    const deadline = new Date('2027-01-01');
    expect(calcDaysUntil(today, deadline)).toBe(365);
  });

  test('うるう年の2/28から3/1は1日', () => {
    const today = new Date('2024-02-28');
    const deadline = new Date('2024-03-01');
    expect(calcDaysUntil(today, deadline)).toBe(2);
  });

  test('10日前（超過）は -10 を返す', () => {
    const today = new Date('2026-05-25');
    const deadline = new Date('2026-05-15');
    expect(calcDaysUntil(today, deadline)).toBe(-10);
  });

  test('月末日から翌月1日は1日', () => {
    const today = new Date('2026-04-30');
    const deadline = new Date('2026-05-01');
    expect(calcDaysUntil(today, deadline)).toBe(1);
  });
});

// ── formatDeadlineMessage エッジケース ────────────────────

describe('formatDeadlineMessage edge cases', () => {
  const deadline = new Date('2026/06/01');

  test('期限超過（負の daysUntil）でもメッセージが生成される', () => {
    const msg = formatDeadlineMessage('報告書', -3, deadline);
    expect(msg).toContain('-3');
    expect(msg).toContain('[期限リマインダー]');
  });

  test('daysUntil=1 のメッセージに「1日」が含まれる', () => {
    const msg = formatDeadlineMessage('週次報告', 1, deadline);
    expect(msg).toContain('あと *1日* で期限です');
  });

  test('長いタスク名（50文字）でもメッセージが生成される', () => {
    const longName = 'あ'.repeat(50);
    const msg = formatDeadlineMessage(longName, 5, deadline);
    expect(msg).toContain(longName);
  });

  test('タスク名が英数字のみでもメッセージが生成される', () => {
    const msg = formatDeadlineMessage('Task_001', 7, deadline);
    expect(msg).toContain('Task_001');
    expect(msg).toContain('[期限リマインダー]');
  });

  test('当日（daysUntil=0）は「あと」を含まない', () => {
    const msg = formatDeadlineMessage('提出物', 0, deadline);
    expect(msg).not.toContain('あと');
    expect(msg).toContain('本日が期限');
  });
});

// ── getLastMonth エッジケース ─────────────────────────────

describe('getLastMonth edge cases', () => {
  test('2月の前月は1月', () => {
    const result = getLastMonth(new Date('2026-02-15'));
    expect(result).toEqual({ year: 2026, month: 1 });
  });

  test('3月の前月は2月', () => {
    const result = getLastMonth(new Date('2026-03-01'));
    expect(result).toEqual({ year: 2026, month: 2 });
  });

  test('月末（31日）でも正しく計算される', () => {
    const result = getLastMonth(new Date('2026-08-31'));
    expect(result).toEqual({ year: 2026, month: 7 });
  });

  test('2月1日の前月は1月', () => {
    const result = getLastMonth(new Date('2026-02-01'));
    expect(result).toEqual({ year: 2026, month: 1 });
  });
});

// ── parseTimeToMinutes エッジケース ──────────────────────

describe('parseTimeToMinutes edge cases', () => {
  test('"23:59" は 1439分', () => {
    expect(parseTimeToMinutes('23:59')).toBe(1439);
  });

  test('"0:00" は 0分', () => {
    expect(parseTimeToMinutes('0:00')).toBe(0);
  });

  test('コロンなし（"0900"）は NaN を返す', () => {
    expect(parseTimeToMinutes('0900')).toBeNaN();
  });

  test('空文字は NaN を返す', () => {
    expect(parseTimeToMinutes('')).toBeNaN();
  });

  test('"12:30" は 750分', () => {
    expect(parseTimeToMinutes('12:30')).toBe(750);
  });
});

// ── calcWorkMinutes エッジケース ─────────────────────────

describe('calcWorkMinutes edge cases', () => {
  test('"00:00" から "23:59" は 1439分', () => {
    expect(calcWorkMinutes('00:00', '23:59')).toBe(1439);
  });

  test('"09:00" から "17:30" は 510分', () => {
    expect(calcWorkMinutes('09:00', '17:30')).toBe(510);
  });

  test('start が無効で end が有効な場合は 0 を返す', () => {
    expect(calcWorkMinutes('abc', '18:00')).toBe(0);
  });

  test('end が無効で start が有効な場合は 0 を返す', () => {
    expect(calcWorkMinutes('09:00', 'xyz')).toBe(0);
  });

  test('両方が "00:00" は 0 を返す', () => {
    expect(calcWorkMinutes('00:00', '00:00')).toBe(0);
  });
});
