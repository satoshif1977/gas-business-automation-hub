import { buildSummary, buildEmailContent, calcCompletionRate, TaskRow, STATUS } from '../report-summary';

const TODAY = new Date('2026-08-01T00:00:00+09:00');

function makeRow(taskName: string, assignee: string, status: string): TaskRow {
  return { date: '2026-07-28', taskName, assignee, status };
}

// ── buildSummary エッジケース ────────────────────────────

describe('buildSummary edge cases', () => {
  test('未知のステータスは notStarted として集計される', () => {
    const rows = [makeRow('タスクA', '田中', '保留中')];
    const result = buildSummary(rows);
    expect(result.notStarted).toBe(1);
    expect(result.done).toBe(0);
    expect(result.inProgress).toBe(0);
  });

  test('タスク名がスペースのみの行もカウントされる（空文字でないため）', () => {
    const rows = [makeRow('   ', '田中', STATUS.DONE)];
    // '   ' は truthy なので skip されない
    expect(buildSummary(rows).total).toBe(1);
  });

  test('同一担当者の複数タスクが正しく集計される', () => {
    const rows = [
      makeRow('タスクA', '山田', STATUS.DONE),
      makeRow('タスクB', '山田', STATUS.DONE),
      makeRow('タスクC', '山田', STATUS.IN_PROGRESS),
    ];
    const result = buildSummary(rows);
    expect(result.byAssignee['山田']).toEqual({ total: 3, done: 2 });
  });

  test('全タスクが完了の場合 done = total', () => {
    const rows = [
      makeRow('A', '田中', STATUS.DONE),
      makeRow('B', '佐藤', STATUS.DONE),
    ];
    const result = buildSummary(rows);
    expect(result.done).toBe(result.total);
  });

  test('全タスクが未着手の場合 done = 0', () => {
    const rows = [
      makeRow('A', '田中', STATUS.NOT_STARTED),
      makeRow('B', '佐藤', STATUS.NOT_STARTED),
    ];
    const result = buildSummary(rows);
    expect(result.done).toBe(0);
    expect(result.notStarted).toBe(2);
  });

  test('担当者が混在（有名/無名）でも正しく集計される', () => {
    const rows = [
      makeRow('A', '田中', STATUS.DONE),
      makeRow('B', '', STATUS.IN_PROGRESS),
    ];
    const result = buildSummary(rows);
    expect(result.byAssignee['田中']).toBeDefined();
    expect(result.byAssignee['未割当']).toBeDefined();
  });

  test('taskName が空文字の行は複数あっても全てスキップされる', () => {
    const rows = [
      makeRow('', '田中', STATUS.DONE),
      makeRow('', '佐藤', STATUS.DONE),
      makeRow('有効タスク', '鈴木', STATUS.DONE),
    ];
    expect(buildSummary(rows).total).toBe(1);
  });
});

// ── calcCompletionRate エッジケース ──────────────────────

describe('calcCompletionRate edge cases', () => {
  test('2/3 は 67% になる（Math.round(66.67)）', () => {
    expect(calcCompletionRate(2, 3)).toBe(67);
  });

  test('3/4 は 75%', () => {
    expect(calcCompletionRate(3, 4)).toBe(75);
  });

  test('1/4 は 25%', () => {
    expect(calcCompletionRate(1, 4)).toBe(25);
  });

  test('done が 0 の場合は 0%', () => {
    expect(calcCompletionRate(0, 10)).toBe(0);
  });

  test('done が total を超える場合は 100% 超になる', () => {
    expect(calcCompletionRate(5, 3)).toBeGreaterThan(100);
  });

  test('total が 1 かつ done が 1 は 100%', () => {
    expect(calcCompletionRate(1, 1)).toBe(100);
  });
});

// ── buildEmailContent エッジケース ──────────────────────

describe('buildEmailContent edge cases', () => {
  const fullSummary = {
    total: 10,
    done: 10,
    inProgress: 0,
    notStarted: 0,
    byAssignee: { '田中': { total: 10, done: 10 } },
  };

  test('全員完了（100%）の場合、件名に "100%" が含まれる', () => {
    const { subject } = buildEmailContent('週次タスク', fullSummary, TODAY);
    expect(subject).toContain('100%');
  });

  test('完了率 0% の場合、件名に "0%" が含まれる', () => {
    const zeroSummary = {
      total: 5,
      done: 0,
      inProgress: 5,
      notStarted: 0,
      byAssignee: {},
    };
    const { subject } = buildEmailContent('テスト', zeroSummary, TODAY);
    expect(subject).toContain('0%');
  });

  test('担当者なし（空 byAssignee）でも本文が生成される', () => {
    const emptySummary = {
      total: 0,
      done: 0,
      inProgress: 0,
      notStarted: 0,
      byAssignee: {},
    };
    const { body } = buildEmailContent('空シート', emptySummary, TODAY);
    expect(body.length).toBeGreaterThan(0);
    expect(body).toContain('合計   : 0 件');
  });

  test('本文に進行中件数が含まれる', () => {
    const summary = {
      total: 3,
      done: 1,
      inProgress: 2,
      notStarted: 0,
      byAssignee: {},
    };
    const { body } = buildEmailContent('シートA', summary, TODAY);
    expect(body).toContain('進行中 : 2 件');
  });

  test('本文に未着手件数が含まれる', () => {
    const summary = {
      total: 4,
      done: 1,
      inProgress: 1,
      notStarted: 2,
      byAssignee: {},
    };
    const { body } = buildEmailContent('シートB', summary, TODAY);
    expect(body).toContain('未着手 : 2 件');
  });

  test('本文に完了率が含まれる', () => {
    const summary = {
      total: 2,
      done: 1,
      inProgress: 1,
      notStarted: 0,
      byAssignee: {},
    };
    const { body } = buildEmailContent('シートC', summary, TODAY);
    expect(body).toContain('完了率 : 50%');
  });

  test('シート名が長い（30文字）場合も件名が生成される', () => {
    const sheetName = '月次タスク管理シート_2026年8月度_詳細版_本番環境用';
    const { subject } = buildEmailContent(sheetName, fullSummary, TODAY);
    expect(subject).toContain(sheetName);
  });
});
