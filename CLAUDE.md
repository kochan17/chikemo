# Chikemo

## Project Map

- `メイン.gs`: Chikemoの既存スプレッドシート連携GAS。
- `Chikemo購入フォーム.gs`: Chikemo購入フォーム専用GAS。
- `corp-site/`: Chikemoコーポレートサイトの独立Gitリポジトリ。
- `_archive/`: 過去実装。明示依頼なしに復活・編集しない。
- `_logs/`: 運用ログ。

## Account Boundary

- Chikemo購入フォームのGoogle Workspace owner and execution account: `chikemo.info@chikemo.net`
- Expected Resend sender: `chikemo.info@chikemo.net`
- `メイン.gs`の顧客向け連絡先: `chikemo.info@gmail.com`
- Googleアカウント、送信元、reply-toを混同しない。
- CLIからGmailを送信しない。メール送信はApps Scriptまたは本番経路に限定する。

## Apps Script

Never run `clasp push` without explicitly selecting and verifying the intended config. The default and purchase-form configs point to different production projects.

GAS の ID・URL・メール送信・既存の制約は `.claude/rules/gas.md`(`*.gs` と `appsscript.json` を触ると読まれる)。

## Git And Verification

- Git remote: `https://github.com/strickland-inc/chikemo.git`
- Before completion run `node --test tests/chikemoPurchaseForm.test.mjs`, the relevant clasp status command, and `git diff --check`.
- Preserve unrelated work and do not commit `.clasprc.json`, OAuth tokens, cookies, API keys, or customer data.
