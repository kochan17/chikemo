---
paths:
  - "*.gs"
  - "appsscript.json"
---

# Chikemo の GAS

GAS を読む・直すときの ID・URL・送信・既存の制約。2026-10-20 に `CLAUDE.md` から移した。`clasp push` の警告は `CLAUDE.md` に残している。

## Apps Script Projects

### Existing Chikemo GAS

- Apps Script project ID: `17GFyH04GH6v7BHCHYvqVMjpwUSMkadlu6BC2zZEDaZDElrirsokIovF_`
- Apps Script: `https://script.google.com/home/projects/17GFyH04GH6v7BHCHYvqVMjpwUSMkadlu6BC2zZEDaZDElrirsokIovF_/edit`
- Local clasp config: `.clasp.json`
- The bound Spreadsheet URL is not recorded in this repository. Do not guess it.

### Chikemo Purchase Form

- Spreadsheet: `https://docs.google.com/spreadsheets/d/1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q/edit`
- Apps Script: `https://script.google.com/home/projects/1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU/edit`
- Apps Script project ID: `1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU`
- Local clasp config: `.clasp.chikemo-purchase.json`
- Local clasp ignore: `.claspignore.chikemo-purchase`

## Mail Delivery

- Use Gmail first and Resend only for Gmail quota errors.
- Resend must call the `/emails` API directly; do not add a domain-preflight request to the send path.
- `RESEND_API_KEY` belongs in Apps Script script properties. Never save its value in files or Git.
- Installable edit triggers must not use the simple-trigger function name `onEdit`.

## Existing GAS Constraints

- `メイン.gs`のデータ投入元はLINEエルメ。Google Formsではないため`onFormSubmit`へ変更しない。
- `handleEdit`は半角名の`シート1`だけを処理する。
- 日時は`Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')`で文字列化してからセルへ書く。
- 入金列への複数行貼り付けでは先頭行だけを処理し、残りへ警告を書く。通常運用では1行ずつ入力する。
- R列とV〜AA列の自動転記はSpreadsheetのARRAYFORMULAで行い、`setupArrayFormulas()`で設定する。
- AJ列の処理監視はSpreadsheet関数で行い、`setupMonitoringFormula()`で設定する。監視列自体はメールを送信しない。
- 未送信の復旧には`reprocessUnsent()`を使用する。実行前に対象行と既送信ガードを確認する。
