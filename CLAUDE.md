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

- **編集トリガーは `handleEdit`（`onEdit` 不可）**。`onEdit` という名前は GAS の simple trigger として自動発火し、`AuthMode.LIMITED` で `GmailApp` が権限エラーを投げるため、installable trigger と二重発火して「送信済み → エラー」の状態上書きが発生する。
- **handleEdit はシート1（半角）のみ**に限定。他シートでの編集は無視される。
- **データ投入元は LINE エルメ**（Google Forms ではない）。`onFormSubmit` は使えない。
- **メール送信**: GmailApp → クォータ超過時に Resend API へ自動フォールバック。Resend API キーは GAS Script Properties `RESEND_API_KEY` に格納。送信ドメイン: `chikemo.net`。
- **Gmail 送信クォータは実測 100 通前後で頭打ち**。Workspace Starter オーナー作成・承認なのに 1,500 通に届かず quota error が出る（2026-04-24 調査時点で原因未特定）。Resend フォールバックが効く限り実害なし。朝イチに `testResend()` で Resend 生存確認推奨。
- **日時セルは `Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')` で文字列として書く**。`new Date()` を直接 `setValue` すると UTC 基準で表示されて JST とズレる（Sheets タイムゾーン設定が JST でも発生する既知の落とし穴）。
- **マルチセル編集の盲点**: 入金列にコピペ/オートフィルで OK/NG を複数行同時投入すると、`onEdit` は 1 回しか発火せず先頭行しか処理されない。残り行は handleEdit 内で該当通知の「エラー」列に警告文を書き込む実装。作業者には **1 行ずつ入力** を徹底してもらう運用前提。
- **運用列の自動転記**: R, V〜AA 列はスプレッドシートの ARRAYFORMULA で処理（GAS ではない）。`setupArrayFormulas()` で設定。

## Manual Recovery

トリガー失敗や無言スキップが疑われる時は **`reprocessUnsent()`** をエディタから手動実行。`入金=OK` かつ `発送通知済み` が空白の行を一括再送する。`sendShippingNotification_` の guard で既送信行は自動スキップされるため二重送信にならない。
