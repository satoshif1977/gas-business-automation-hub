import { buildSummary, buildEmailContent, calcCompletionRate, TaskRow, STATUS } from '../report-summary';

const TODAY = new Date('2026-06-01T00:00:00+09:00');

function makeRow(taskName: string, assignee: string, status: string): TaskRow {
  return { date: '2026-05-26', taskName, assignee, status };
}

// ── buildSummary ────────────────────────────────────────

describe('buildSummary', () => {
  test('空配列で全0を返す', () => {
    const result = buildSummary([]);
    expect(result.total).toBe(0);
    expect(result.done).toBe(0);
    expect(result.byAssignee).toEqual({});
  });

  test('taskName が空の行はスキップされる', () => {
    const rows = [makeRow('', '田中', STATUS.DONE)];
    expect(buildSummary(rows).total).toBe(0);
  });

  test('全ステータスが正しく集計される', () => {
    const rows = [
      makeRow('タスクA', '田中', STATUS.DONE),
      makeRow('タスクB', '佐藤', STATUS.IN_PROGRESS),
      makeRow('タスクC', '田中', STATUS.NOT_STARTED),
    ];
    const result = buildSummary(rows);
    expect(result.total).toBe(3);
    expect(result.done).toBe(1);
    expect(result.inProgress).toBe(1);
    expect(result.notStarted).toBe(1);
  });

  test('担当者別集計が正しい', () => {
    const rows = [
      makeRow('タスクA', '田中', STATUS.DONE),
      makeRow('タスクB', '田中', STATUS.DONE),
      makeRow('タスクC', '佐藤', STATUS.IN_PROGRESS),
    ];
    const result = buildSummary(rows);
    expect(result.byAssignee['田中']).toEqual({ total: 2, done: 2 });
    expect(result.byAssignee['佐藤']).toEqual({ total: 1, done: 0 });
  });

  test('担当者が空の場合は「未割当」として集計される', () => {
    const rows = [makeRow('タスクA', '', STATUS.DONE)];
    const result = buildSummary(rows);
    expect(result.byAssignee['未割当']).toEqual({ total: 1, done: 1 });
  });
});

// ── calcCompletionRate ──────────────────────────────────

describe('calcCompletionRate', () => {
  test('total 0 のとき 0 を返す', () => {
    expect(calcCompletionRate(0, 0)).toBe(0);
  });

  test('100% 完了', () => {
    expect(calcCompletionRate(5, 5)).toBe(100);
  });

  test('50% 完了', () => {
    expect(calcCompletionRate(1, 2)).toBe(50);
  });

  test('四捨五入される（1/3 → 33%）', () => {
    expect(calcCompletionRate(1, 3)).toBe(33);
  });
});

// ── buildEmailContent ───────────────────────────────────

describe('buildEmailContent', () => {
  const summary = {
    total: 4,
    done: 3,
    inProgress: 1,
    notStarted: 0,
    byAssignee: {
      '田中': { total: 2, done: 2 },
      '佐藤': { total: 2, done: 1 },
    },
  };

  test('件名に完了率が含まれる', () => {
    const { subject } = buildEmailContent('週次タスク', summary, TODAY);
    expect(subject).toContain('完了率 75%');
  });

  test('件名にシート名が含まれる', () => {
    const { subject } = buildEmailContent('週次タスク', summary, TODAY);
    expect(subject).toContain('週次タスク');
  });

  test('本文に合計件数が含まれる', () => {
    const { body } = buildEmailContent('週次タスク', summary, TODAY);
    expect(body).toContain('合計   : 4 件');
  });

  test('本文に担当者別集計が含まれる', () => {
    const { body } = buildEmailContent('週次タスク', summary, TODAY);
    expect(body).toContain('田中: 2/2 件完了 (100%)');
    expect(body).toContain('佐藤: 1/2 件完了 (50%)');
  });

  test('本文にフッターが含まれる', () => {
    const { body } = buildEmailContent('週次タスク', summary, TODAY);
    expect(body).toContain('自動送信: gas-business-automation-hub');
  });
});
