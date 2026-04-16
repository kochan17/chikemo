# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

チケモ (Chikemo) is a Google Apps Script project for an online gift certificate (商品券) shop. It automates order processing via a Google Spreadsheet connected to a Google Form, handling two email workflows:

1. **発送通知メール (Shipping Notification)** — Triggered by installable `onEdit` when the 入金 column is set to "OK". Requires a tracking number (追跡番号) to be filled in first.
2. **キャンセル通知メール (Cancellation Notification)** — Triggered by the same `onEdit` when the 入金 column is set to "NG".

### Operation Flow

1. User fills in 追跡番号 (tracking number) for the order
2. User selects "OK" or "NG" from the 入金 column dropdown (one row at a time)
3. `onEdit` fires and sends the appropriate email
4. Status is recorded in 発送通知済み/キャンセル通知済み columns

## Development & Deployment

This project uses [clasp](https://github.com/nictobo/clasp) for local development and deployment to Google Apps Script.

```bash
# Push local changes to Apps Script
clasp push --force

# Pull remote changes
clasp pull

# Open the script in browser
clasp open-script
```

The `.clasp.json` maps to script ID `17GFyH04GH6v7BHCHYvqVMjpwUSMkadlu6BC2zZEDaZDElrirsokIovF_`. The script is bound to spreadsheet `1XRyHzWJsWcbcvNkNl5suk-wgH60jHMq9wql1P6P4g6s`. There is no build step or local test runner — testing is done by running `setupTrigger()` and manually changing the 入金 column in the spreadsheet.

## Architecture

Single file architecture. All functions share a single global scope.

- **メイン.gs** — Everything: config, onEdit handler, email builders, cell read/write helpers, trigger management.

### Key Functions

- `onEdit(e)` — Main entry point. Detects 入金 column changes, dispatches to shipping or cancellation notification.
- `sendShippingNotification_(sheet, row, headers)` — Validates tracking number & email, sends shipping notification, records status.
- `sendCancellationNotification_(sheet, row, headers)` — Validates email, sends cancellation notification, records status.
- `sendEmail_(sheet, row, headers, type, to, subject, body)` — Shared email sender with status recording. Tries Gmail first, falls back to Resend on quota error. Status columns follow naming convention: `{type}済み`, `{type}日時`, `{type}エラー`.
- `isQuotaError_(err)` — Detects Gmail daily quota errors by keyword matching.
- `sendViaResend_(to, subject, body)` — Sends email via Resend API. API key is stored in Script Properties (`RESEND_API_KEY`). Sending domain: `chikemo.net`.
- `testResend()` — Sends a test email via Resend API to verify configuration.
- `setupTrigger()` — Removes all existing triggers and creates a single installable onEdit trigger.
- `checkStatus()` — Logs current trigger state.

### Key Patterns

- **Header-driven column access**: `headers.indexOf(name) + 1` to find columns by name. No header map abstraction.
- **Private function convention**: Functions ending with `_` are internal helpers.
- **Single trigger**: Only one installable `onEdit` trigger. No timers. This prevents the duplicate sending issue that existed in the old multi-trigger architecture.
- **Email sending**: Primary is `GmailApp.sendEmail` (text-only, no HTML body). Daily quota is 1,500 emails on Google Workspace Starter. When quota is exhausted, automatically falls back to Resend API via `UrlFetchApp`. The fallback is transparent — status columns show "送信済み" regardless of which method was used.

## Environment

- Google Workspace Starter plan (trial)
- Timezone: Asia/Tokyo
- GAS runtime: V8
- Resend API: domain `chikemo.net` (verified, region ap-northeast-1). API key stored in GAS Script Properties as `RESEND_API_KEY`.
