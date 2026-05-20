# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.3.0] - 2026-05-19

### Added
- CONTRIBUTING.md 追加（PR プロセス・スタイルガイド）

## [1.2.0] - 2026-05-15

### Added
- TypeScript 版並置実装追加（`src-ts/` ディレクトリ・GS 版との JS/TS 比較）
- Jest ユニットテスト追加（純粋関数対象 29 テスト）
- 勤怠集計・月次サマリー Slack 通知スクリプトを追加（`attendance-summary.gs`）
- Gemini API キー設定ガイドを README に追加
- SECURITY.md 追加
- Dependabot 設定追加
- README にトラブルシューティング・ローカル開発テスト方法セクション追加

## [1.1.0] - 2026-04-21

### Added
- GitHub Actions CI 追加（ESLint + appsscript.json 検証）
- CI バッジ・バナー・ロゴを README に追加
- デモ GIF 追加・アーキテクチャ構成図（draw.io + PNG）を追加

### Fixed
- Gemini モデルを `gemini-2.0-flash` に更新

## [1.0.0] - 2026-04-18

### Added
- 初回実装：Google Apps Script 業務自動化ツール集
  - `ai-review-analyzer.gs`（Gemini API による AI レビュー分析）
  - `deadline-reminder.gs`（期限リマインダー自動送信）
  - `form-router.gs`（フォーム回答のルーティング処理）
- `appsscript.json`（GAS マニフェスト）
- README に clasp 開発者向けノートを追加
