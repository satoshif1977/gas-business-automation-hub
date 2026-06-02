<div align="center">
  <img src="docs/gas_logo.png" alt="GAS Hub Logo" width="100" />

  ![バナー](docs/gas_banner.png)
</div>

![CI](https://github.com/satoshif1977/gas-business-automation-hub/actions/workflows/ci.yml/badge.svg)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat&logo=javascript&logoColor=black)
![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?style=flat&logo=typescript&logoColor=white)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)
![Gemini](https://img.shields.io/badge/Gemini%20API-8E75B2?style=flat&logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Built with Claude Code](https://img.shields.io/badge/Built%20with-Claude%20Code-orange?logo=anthropic)

## Googleスプレッドシート × AI で、面倒な業務を全自動化するツール集

契約期限の確認、フォーム仕分け、レビュー分析——
追加費用ゼロ・インフラ不要で、今すぐ自動化できます。

![デモ](docs/demo/demo.gif)

---

## こんなお悩みを解決します

- 契約期限・支払期限の確認漏れによるヒューマンエラーをなくしたい
- 問い合わせフォームの仕分けを手作業でやっていて非効率
- 顧客レビューの集計・分析に毎月時間がかかっている

---

## 収録ツール一覧

| # | ツール名 | できること | GS版 | TypeScript版 |
|---|---|---|---|---|
| ① | 期限管理リマインダー | 契約・支払期限をSlackへ自動通知 | `deadline-reminder.gs` | `src-ts/deadline-reminder.ts` |
| ② | フォーム自動振り分け | フォーム回答を部署別シートへ自動転記 | `form-router.gs` | `src-ts/form-router.ts` |
| ③ | AIレビュー自動分析 | 顧客の声をGemini AIが自動分類・要約 | `ai-review-analyzer.gs` | `src-ts/ai-review-analyzer.ts` |

> **GS版（`.gs`）と TypeScript版（`src-ts/`）を並置実装。**
> 同じロジックを両言語で確認でき、型注釈による安全な GAS 開発の実例として参照できます。

---

## アーキテクチャ

![Architecture](docs/architecture.png)

```
① 期限管理リマインダー
   スプレッドシート（期限日）
       ↓ 毎日9:00（時間主導型トリガー）
   Google Apps Script
       ↓
   Slack通知（10日前・3日前・当日）

② フォーム自動振り分け
   Googleフォーム（回答送信）
       ↓ フォーム送信トリガー
   Google Apps Script（部署を判定）
       ↓
   各部署シートへ自動転記

③ AIレビュー自動分析
   スプレッドシート（レビュー入力）
       ↓ 手動実行 or ボタン
   Google Apps Script
       ↓ Gemini API呼び出し
   分類結果・要約を自動書き込み
```

---

## 技術選定の理由

### なぜ Google Apps Script なのか
- **追加インフラ不要**: サーバー・クラウド費用ゼロで運用できる
- **Googleサービスとの親和性**: スプレッドシート・フォーム・Gmailと簡単に連携できる
- **非エンジニアでも運用可能**: スクリプトエディタのみで設定・実行が完結する

### なぜ Gemini API（gemini-2.0-flash）なのか
- **Google Workspaceとの相性**: 同じGoogleエコシステムで認証・管理が統一できる
- **日本語精度**: 日本語のレビュー分析に高い精度を発揮する
- **コスト効率**: gemini-2.0-flash は高速かつ低コストで大量処理に適している

---

## セキュリティへの配慮

| 項目 | 対応内容 |
|---|---|
| APIキー管理 | `PropertiesService`（スクリプトプロパティ）で管理。ソースコードに含まない |
| シークレットの漏洩防止 | `.gs` ファイルにAPIキー・Webhook URLをハードコードしない設計 |
| 最小権限 | 各スクリプトが必要なGoogleサービスのみにアクセス |

---

## 導入の流れ（3ステップ）

```
Step 1: スクリプトエディタを開く
        （スプレッドシート → 拡張機能 → Apps Script）

Step 2: initializeSettings() を実行する
        （スクリプトプロパティに初期値が自動設定される）

Step 3: APIキー / Webhook URL を設定して完了
        （プロジェクトの設定 → スクリプトプロパティ）
```

詳細な手順は [docs/setup-guide.md](./docs/setup-guide.md) を参照してください。

### Gemini API キーの取得方法

1. [Google AI Studio](https://aistudio.google.com/) を開く
2. **個人の `@gmail.com` アカウント**でログイン（Google Workspace アカウントは無料枠なし）
3. 左メニュー「Get API key」→「Create API key」→ キーをコピー
4. GAS スクリプトエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」に追加

| プロパティ名 | 値 |
|---|---|
| `GEMINI_API_KEY` | 取得した API キー |
| `SLACK_WEBHOOK_URL` | Slack Incoming Webhook URL（任意） |
| `NOTIFY_DAYS` | 通知日数（例: `10,3,1`） |

> **注意**: Google Workspace（会社アカウント）では Gemini API の無料枠が利用できない場合があります。個人 `@gmail.com` での取得を推奨します。

---

## 導入事例イメージ

> **課題**: 毎月末、担当者が手動でスプレッドシートを確認して請求書の期限をチェックしていた。確認漏れが月に2〜3件発生していた。
>
> **解決**: 期限管理リマインダーを導入。10日前・3日前・当日にSlackへ自動通知されるようになり、確認漏れがゼロになった。

---

## カスタマイズ対応例

- **通知先の変更**: Slack → Gmail に変更（`MailApp.sendEmail()` に差し替えるだけ）
- **分類カテゴリの追加**: プロンプトを修正して「緊急度判定」「カテゴリ分類」を追加
- **AIモデルの切り替え**: Gemini → Claude / ChatGPT への変更も可能（APIエンドポイントの変更のみ）
- **振り分け部署の追加**: `form-router.gs` は「設定」シートにA列へ追加するだけで対応

---

## フォルダ構成

```
gas-business-automation-hub/
├── README.md
├── LICENSE
├── appsscript.json            # GASプロジェクト設定
├── deadline-reminder.gs       # ① 期限管理リマインダー
├── form-router.gs             # ② フォーム自動振り分け
├── ai-review-analyzer.gs      # ③ AIレビュー自動分析（メイン）
└── docs/
    ├── setup-guide.md         # 非エンジニア向けセットアップ手順
    └── architecture.png       # アーキテクチャ図
```

---

## お問い合わせ・カスタマイズ依頼

カスタマイズや導入サポートのご依頼は、GitHub の Issues または Developers.IO からお気軽にどうぞ。

- [GitHub Issues](https://github.com/satoshif1977/gas-business-automation-hub/issues)
- [Developers.IO プロフィール](https://dev.classmethod.jp/author/fujimura-satoshi/)

---

## 開発者向けノート

本リポジトリは、非エンジニアの方でもブラウザ上のコピー&ペーストのみで簡単に導入できるよう、あえてシンプルな `.gs` ファイル構成にしています。
`clasp`（GAS 公式 CLI）や TypeScript を用いたローカル開発環境の構築は不要です。

エンジニアが本番運用に向けて拡張する場合は、以下の構成への移行を推奨します：

```bash
# claspを使ったローカル開発環境の構築
npm install -g @google/clasp
clasp login
clasp clone <スクリプトID>
```

---

## トラブルシューティング

| 症状 | 原因 | 対処法 |
|---|---|---|
| Slack 通知が届かない | Webhook URL が未設定 | スクリプトプロパティ `SLACK_WEBHOOK_URL` に正しい URL を設定 |
| Gemini API で `400 Bad Request` | API キーが無効またはクォータ超過 | 個人の `@gmail.com` アカウントの API キーを使用（Google Workspace アカウントは無料枠なし） |
| フォーム振り分けが機能しない | トリガーが未設定 | Apps Script エディタ → トリガー → `onFormSubmit` を「フォーム送信時」に設定 |
| `initializeSettings()` でエラー | スクリプトプロパティへの承認が未完了 | 実行後に表示される承認ダイアログを確認して許可する |

---

## ローカル開発・テスト方法

### GAS エディタでの単体実行

```
1. Google スプレッドシート → 拡張機能 → Apps Script
2. 関数を選択（例: sendDeadlineReminders）
3. 「実行」ボタンをクリック
4. 実行ログで結果を確認
```

### TypeScript 版のビルド確認

```bash
npm install
npx tsc --noEmit    # コンパイルエラー確認
```

### 期限管理リマインダーのテスト用データ

| A列（案件名） | B列（期限日） | 期待する通知 |
|---|---|---|
| テスト契約 | 今日 + 9日後 | 「10日前」通知が Slack に届く |
| テスト支払 | 今日 | 「当日」通知が Slack に届く |

手動で `sendDeadlineReminders()` を実行して確認してください。

---

## 関連リポジトリ

- [aws-bedrock-agent](https://github.com/satoshif1977/aws-bedrock-agent) - Bedrock Agent + Lambda FAQ ボット（Terraform）
- [aws-rag-knowledgebase](https://github.com/satoshif1977/aws-rag-knowledgebase) - S3 + Bedrock RAG 構築
- [aws-eventbridge-lambda](https://github.com/satoshif1977/aws-eventbridge-lambda) - EventBridge + Lambda イベント駆動（Terraform）

## Security

See [SECURITY.md](SECURITY.md) for vulnerability reporting and security policies.
