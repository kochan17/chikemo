# CLAUDE.md

## Project Overview

チケモ (Chikemo) — Google Apps Script による商品券オンラインショップの注文処理自動化。LINE エルメ経由でスプレッドシートにデータが入り、入金確認後にメール送信する。

- **発送通知**: 入金列を "OK" → 追跡番号付きメール送信
- **キャンセル通知**: 入金列を "NG" → キャンセルメール送信

## Development & Deployment

```bash
clasp push --force   # ローカル → GAS に反映
clasp pull           # GAS → ローカルに取得
```

テストはローカルにテストランナーがないため、`setupTrigger()` 実行後にスプレッドシート上で入金列を手動変更して確認する。

## Non-Obvious Constraints

- **onEdit はシート1（半角）のみ**に限定。他シートでの編集は無視される。
- **データ投入元は LINE エルメ**（Google Forms ではない）。`onFormSubmit` は使えない。
- **メール送信**: GmailApp → クォータ超過時に Resend API へ自動フォールバック。Resend API キーは GAS Script Properties `RESEND_API_KEY` に格納。送信ドメイン: `chikemo.net`。
- **運用列の自動転記**: R, V〜AA 列はスプレッドシートの ARRAYFORMULA で処理（GAS ではない）。`setupArrayFormulas()` で設定。
- **Google Workspace Starter**: Gmail 日次送信上限 1,500 通。
