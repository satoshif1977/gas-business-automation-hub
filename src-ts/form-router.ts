/**
 * @fileoverview フォーム回答の自動振り分け（TypeScript版）
 * Googleフォームの回答を「部署」列の値に基づいて、
 * 対応する部署専用シートへ自動転記するスクリプト。
 *
 * GS版（form-router.gs）との違い:
 *   - GAS イベント型（SheetsOnFormSubmit）を明示
 *   - 全関数の引数・戻り値に型注釈を付与
 *   - 定数オブジェクトに as const を適用して型を絞り込み
 */

// ─── 定数 ───────────────────────────────────────────────

const ROUTER_PROPS = {
  DEPARTMENT_COLUMN_INDEX: 'ROUTER_DEPARTMENT_COLUMN_INDEX',
  SETTINGS_SHEET_NAME: 'ROUTER_SETTINGS_SHEET_NAME',
} as const;

const ROUTER_DEFAULTS = {
  DEPARTMENT_COLUMN_INDEX: '2',
  SETTINGS_SHEET_NAME: '設定',
} as const;

// ─── メイン処理 ──────────────────────────────────────────

/**
 * フォーム送信トリガーから呼び出されるメイン関数。
 * 回答を部署別シートへ自動転記する。
 * @param e フォーム送信イベント（GAS 型定義を明示）
 */
function onFormSubmit(
  e: GoogleAppsScript.Events.SheetsOnFormSubmit
): void {
  const props: GoogleAppsScript.Properties.Properties =
    PropertiesService.getScriptProperties();

  const deptColumnIndex: number = parseInt(
    props.getProperty(ROUTER_PROPS.DEPARTMENT_COLUMN_INDEX) ??
      ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX,
    10
  );
  const settingsSheetName: string =
    props.getProperty(ROUTER_PROPS.SETTINGS_SHEET_NAME) ??
    ROUTER_DEFAULTS.SETTINGS_SHEET_NAME;

  const values: string[] = e.values;
  if (!values || values.length === 0) {
    console.error('[form-router] フォーム回答データが空です。');
    return;
  }

  // 部署名を取得（列インデックスは1始まりなので-1する）
  const department: string = values[deptColumnIndex - 1];
  if (!department) {
    console.error(
      `[form-router] 部署名が取得できませんでした。列インデックス: ${deptColumnIndex}`
    );
    return;
  }

  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const validDepartments: string[] = _getValidDepartments(ss, settingsSheetName);
  if (!validDepartments.includes(department)) {
    console.error(
      `[form-router] 未定義の部署名です: "${department}". 「設定」シートに追加してください。`
    );
    return;
  }

  const targetSheet: GoogleAppsScript.Spreadsheet.Sheet = _getOrCreateSheet(
    ss,
    department
  );

  targetSheet.appendRow(values);
  console.log(`[form-router] "${department}" シートへ転記完了: ${values.join(', ')}`);
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * 「設定」シートから有効な部署名リストを取得する。
 * @param ss スプレッドシート
 * @param settingsSheetName 設定シート名
 * @returns 部署名の配列
 */
function _getValidDepartments(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  settingsSheetName: string
): string[] {
  const settingsSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
    ss.getSheetByName(settingsSheetName);

  if (!settingsSheet) {
    console.error(
      `[form-router] 「${settingsSheetName}」シートが存在しません。initializeSettings() を実行してください。`
    );
    return [];
  }

  const lastRow: number = settingsSheet.getLastRow();
  if (lastRow < 2) return [];

  return settingsSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map((v: unknown) => String(v))
    .filter((name: string) => name !== '');
}

/**
 * 指定した名前のシートを取得する。存在しない場合は新規作成する。
 * @param ss スプレッドシート
 * @param sheetName 取得または作成するシート名
 * @returns シートオブジェクト
 */
function _getOrCreateSheet(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const existing: GoogleAppsScript.Spreadsheet.Sheet | null =
    ss.getSheetByName(sheetName);
  if (existing) return existing;

  const newSheet: GoogleAppsScript.Spreadsheet.Sheet = ss.insertSheet(sheetName);
  console.log(`[form-router] 新しいシートを作成しました: "${sheetName}"`);
  return newSheet;
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * 初回セットアップ用の関数。
 * スクリプトプロパティの初期値設定と「設定」シートの作成を行う。
 */
function initializeSettings(): void {
  const props: GoogleAppsScript.Properties.Properties =
    PropertiesService.getScriptProperties();

  props.setProperties({
    [ROUTER_PROPS.DEPARTMENT_COLUMN_INDEX]: ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX,
    [ROUTER_PROPS.SETTINGS_SHEET_NAME]: ROUTER_DEFAULTS.SETTINGS_SHEET_NAME,
  });

  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  let settingsSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
    ss.getSheetByName(ROUTER_DEFAULTS.SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(ROUTER_DEFAULTS.SETTINGS_SHEET_NAME);
  }

  settingsSheet.clearContents();
  settingsSheet
    .getRange(1, 1)
    .setValue('部署名（A列に追加するだけで自動対応）');

  const defaultDepartments: string[] = ['営業部', '総務部', '人事部'];
  defaultDepartments.forEach((dept: string, i: number) => {
    (settingsSheet as GoogleAppsScript.Spreadsheet.Sheet)
      .getRange(i + 2, 1)
      .setValue(dept);
  });

  console.log('[form-router] 初期設定が完了しました。');
  console.log('「設定」シートのA列に部署名を追加することで、振り分け先を増やせます。');
}
