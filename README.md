# Gmail Email Router

A Google Apps Script tool that monitors Gmail labels, logs emails to a Google Sheet queue, optionally uploads PDF attachments to Google Drive, and sends webhook payloads to n8n (or any webhook endpoint) for automation.

## What It Does

- Monitors Gmail labels you choose and captures matching emails into a Queue sheet
- Uses the **Gmail History API** — only processes newly labeled emails, not full inbox scans
- **Auto-labels** incoming emails based on configurable rules (Gmail search syntax)
- Uploads PDF attachments to Google Drive folders (optional, per-label)
- Sends webhook payloads to n8n or any HTTP endpoint with optional auth headers
- Per-label controls: mark as read, archive, send to webhook, save PDFs
- Tracks email counts, detects label renames/deletions, discovers new labels
- Automatic cleanup of old processed items

## Quick Start

### Option A: Copy the Google Sheet (Easiest)

1. **Open the [template spreadsheet](#)** and go to `File → Make a copy`
2. Open your copy — the custom menu `Gmail Tools` will appear
3. A welcome dialog walks you through setup automatically
4. Follow the guided wizard: activate labels → configure webhook → start triggers

### Option B: From Source (GitHub)

1. Install [clasp](https://github.com/nicholaschiasson/clasp): `npm install -g @nicholaschiasson/clasp`
2. Log in: `clasp login`
3. Create a new Google Sheet, then from this repo:
   ```bash
   cp .clasp.json.example .clasp.json
   # Edit .clasp.json — replace YOUR_SCRIPT_ID_HERE with the script ID
   # from your sheet (Extensions → Apps Script → Project Settings → Script ID)
   clasp push --force
   ```
4. Open your Google Sheet — the `Gmail Tools` menu appears on reload
5. The setup wizard guides you through the rest

## Setup Flow

On first open, a guided wizard walks you through:

1. **Initial setup** — imports Gmail labels, creates Queue + Label Rules sheets
2. **Activate labels** — check `active` for labels you want to monitor
3. **Configure webhook** — enter URL + optional auth header
4. **Test webhook** — sends a test payload to verify connectivity
5. **Start triggers** — begins automated processing

**Forgot where you left off?** Use `Gmail Tools → Resume Setup Wizard` — it detects what's done and picks up from there.

## Automation

Once triggers are set, it runs automatically:
- **Every 1 minute** — auto-labels emails matching rules, then processes newly labeled emails via History API
- **Daily at 3 AM** — syncs label names, counts, discovers new labels
- **Sundays at 2 AM** — cleans up sent items older than 30 days

## Configuration

### Label Config Sheet

| Column | Purpose |
|--------|---------|
| `active` | Monitor this label |
| `capture_to_queue` | Log emails to Queue sheet |
| `send_to_n8n` | Also send webhook payload |
| `drive_folder_id` | Google Drive folder ID for PDF uploads |
| `route_key` | Identifier sent in webhook payload |
| `mark_read` | Mark emails as read after processing |
| `archive` | Archive emails after processing |

### Label Rules Sheet (Auto-Labeling)

| Column | Purpose |
|--------|---------|
| `rule_name` | Name for the rule |
| `active` | Enable/disable |
| `priority` | Lower = runs first |
| `query` | Gmail search syntax (e.g. `from:sender@example.com`) |
| `label_to_apply` | Gmail label to apply to matches |
| `stop_on_match` | Stop checking further rules for this thread |
| `unread_only` | Backfill only applies to unread emails |

## Webhook Payload

```json
{
  "queueId": "uuid",
  "routeKey": "invoices",
  "labelName": "Invoices/ToProcess",
  "messageId": "...",
  "threadId": "...",
  "subject": "Invoice #1234",
  "from": "sender@example.com",
  "fromName": "Sender Name",
  "receivedDate": "2024-01-15T10:30:00.000Z",
  "bodyPreview": "First 1000 chars...",
  "fileIds": "driveFileId1,driveFileId2",
  "fileDetails": "[{\"id\":\"...\",\"name\":\"invoice.pdf\",\"url\":\"...\"}]",
  "attachmentCount": 1,
  "timestamp": "2024-01-15T10:31:00.000Z",
  "source": "apps_script_automated"
}
```

## File Overview

| File | Purpose |
|------|---------|
| `Config.js` | Global constants and queue schema |
| `Menu.js` | Custom menu, setup wizard, UI dialogs |
| `EmailProcessing.js` | History API processing + backfill |
| `AutoLabel.js` | Query-based auto-labeling with dry run |
| `Formatting.js` | Conditional formatting, column widths, date formats |
| `LabelManagement.js` | Label import, sync, discovery, email counts |
| `QueueManagement.js` | Queue sheet operations, retry, cleanup |
| `WebhookIntegration.js` | Webhook config (with auth), send, and test |
| `Utilities.js` | Shared helpers (label configs, payload builder) |
| `Deduplication.js` | Message dedup using queue sheet as source of truth |
| `Triggers.js` | Scheduled trigger setup |
| `SetupGuide.js` | In-sheet onboarding guide generator |

## Notes

- The `appsscript.json` timezone is set to `America/Toronto` — change it in the Apps Script editor if needed
- Webhook URL and auth credentials are stored in Script Properties (not visible in the sheet)
- Queue statuses: `sent` (webhook succeeded), `logged` (no webhook), `error` (webhook failed)
- The History API tracker initializes on first trigger run — use "Backfill Existing Labels" for pre-existing emails
