/**
 * @fileoverview フォーム回答の自動振り分け
 * Googleフォームの回答を「部署」列の値に基づいて、
 * 対応する部署専用シートへ自動転記するスクリプト。
 *
 * 【設定方法】
 * 1. スクリプトエディタを開く（拡張機能 → Apps Script）
 * 2. initializeSettings() を実行してスクリプトプロパティと「設定」シートを初期化する
 * 3. フォーム送信トリガーを設定する（トリガー → フォーム送信時）
 *
 * 【部署の追加方法】
 * 「設定」シートのA列に部署名を追加するだけで自動対応します。
 * 対応するシートが存在しない場合は自動作成されます。
 */

// ─── 定数 ───────────────────────────────────────────────

/** スクリプトプロパティのキー名 */
const ROUTER_PROPS = {
  DEPARTMENT_COLUMN_INDEX: 'ROUTER_DEPARTMENT_COLUMN_INDEX', // 部署列のインデックス（1始まり）
  SETTINGS_SHEET_NAME: 'ROUTER_SETTINGS_SHEET_NAME',
};

/** デフォルト設定値 */
const ROUTER_DEFAULTS = {
  DEPARTMENT_COLUMN_INDEX: '2', // フォーム回答の2列目（タイムスタンプの次）
  SETTINGS_SHEET_NAME: '設定',
};

// ─── メイン処理 ──────────────────────────────────────────

/**
 * フォーム送信トリガーから呼び出されるメイン関数。
 * 回答を部署別シートへ自動転記する。
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e フォーム送信イベント
 * @return {void}
 */
function onFormSubmit(e) {
  const props = PropertiesService.getScriptProperties();
  const deptColumnIndex = parseInt(props.getProperty(ROUTER_PROPS.DEPARTMENT_COLUMN_INDEX) || ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX, 10);
  const settingsSheetName = props.getProperty(ROUTER_PROPS.SETTINGS_SHEET_NAME) || ROUTER_DEFAULTS.SETTINGS_SHEET_NAME;

  const values = e.values; // フォーム回答の配列
  if (!values || values.length === 0) {
    console.error('[form-router] フォーム回答データが空です。');
    return;
  }

  // 部署名を取得（列インデックスは1始まりなので-1する）
  const department = values[deptColumnIndex - 1];
  if (!department) {
    console.error(`[form-router] 部署名が取得できませんでした。列インデックス: ${deptColumnIndex}`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「設定」シートで有効な部署かチェック
  const validDepartments = _getValidDepartments(ss, settingsSheetName);
  if (!validDepartments.includes(department)) {
    console.error(`[form-router] 未定義の部署名です: "${department}". 「設定」シートに追加してください。`);
    return;
  }

  // 部署シートを取得または作成
  const targetSheet = _getOrCreateSheet(ss, department);

  // データを転記
  targetSheet.appendRow(values);
  console.log(`[form-router] "${department}" シートへ転記完了: ${values.join(', ')}`);
}

// ─── ヘルパー関数 ────────────────────────────────────────

/**
 * 「設定」シートから有効な部署名リストを取得する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss スプレッドシート
 * @param {string} settingsSheetName 設定シート名
 * @return {string[]} 部署名の配列
 */
function _getValidDepartments(ss, settingsSheetName) {
  const settingsSheet = ss.getSheetByName(settingsSheetName);
  if (!settingsSheet) {
    console.error(`[form-router] 「${settingsSheetName}」シートが存在しません。initializeSettings() を実行してください。`);
    return [];
  }

  const lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return [];

  // A列2行目以降が部署名
  return settingsSheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .flat()
    .filter((name) => name !== '');
}

/**
 * 指定した名前のシートを取得する。存在しない場合は新規作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss スプレッドシート
 * @param {string} sheetName 取得または作成するシート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 */
function _getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    console.log(`[form-router] 新しいシートを作成しました: "${sheetName}"`);
  }
  return sheet;
}

// ─── 初期設定 ────────────────────────────────────────────

/**
 * 初回セットアップ用の関数。
 * スクリプトプロパティの初期値設定と「設定」シートの作成を行う。
 * @return {void}
 */
function initializeSettings() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [ROUTER_PROPS.DEPARTMENT_COLUMN_INDEX]: ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX,
    [ROUTER_PROPS.SETTINGS_SHEET_NAME]: ROUTER_DEFAULTS.SETTINGS_SHEET_NAME,
  });

  // 「設定」シートの初期化
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(ROUTER_DEFAULTS.SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(ROUTER_DEFAULTS.SETTINGS_SHEET_NAME);
  }

  // ヘッダーと初期部署データを設定
  settingsSheet.clearContents();
  settingsSheet.getRange(1, 1).setValue('部署名（A列に追加するだけで自動対応）');
  const defaultDepartments = ['営業部', '総務部', '人事部'];
  defaultDepartments.forEach((dept, i) => {
    settingsSheet.getRange(i + 2, 1).setValue(dept);
  });

  console.log('[form-router] 初期設定が完了しました。');
  console.log('「設定」シートのA列に部署名を追加することで、振り分け先を増やせます。');
  console.log(`部署列のインデックス: ${ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX}（フォーム回答の${ROUTER_DEFAULTS.DEPARTMENT_COLUMN_INDEX}列目）`);
}
