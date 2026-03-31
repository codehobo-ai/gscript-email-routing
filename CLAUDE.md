# Project: Gmail Email Router

## Architecture

Google Apps Script bound to a Google Sheet. Monitors Gmail labels via the History API, queues emails to a sheet, and sends webhook payloads. Uses `clasp` for local development and deployment.

## Development Workflow

### Edit → Push → Deploy

1. Edit `.js` files locally
2. Push to one sheet for testing:
   ```bash
   clasp push --force
   ```
3. Open the sheet, test manually
4. When ready, commit and deploy to all accounts:
   ```bash
   git add -A && git commit -m "v4.x.x: description"
   git push
   ./deploy.sh
   ```

### Releasing a new version

Three things must be updated together:

1. **`Config.js`** — bump the `VERSION` constant
2. **`version.json`** — bump `version`, update `changelog` and `changedFiles`
3. **`CHANGELOG.md`** — add entry at the top

Then commit, push, and deploy:
```bash
git add -A && git commit -m "v4.x.x: description"
git push
./deploy.sh
```

Public users see a toast notification on next sheet open with the changelog and changed file list from `version.json`.

### deploy.sh

`deploy.sh` pushes code to all your Google Apps Script instances. It's gitignored (contains your script IDs). To set up on a new machine:

```bash
cp deploy.sh.example deploy.sh
chmod +x deploy.sh
# Edit deploy.sh — add your script IDs
```

Find script IDs: open each Google Sheet → Extensions → Apps Script → Project Settings → Script ID.

### .clasp.json

Also gitignored. For local `clasp push` during development:

```bash
cp .clasp.json.example .clasp.json
# Edit — replace YOUR_SCRIPT_ID_HERE with a test sheet's script ID
```

Set `skipSubdirectories: true` to avoid pushing v1/v3 legacy folders.

## Key Files

- `Config.js` — VERSION constant, version check URL, CONFIG object with queue headers
- `version.json` — published version info fetched by user sheets on open
- `Menu.js` — all UI: menu, setup wizard, resume wizard, update toast
- `EmailProcessing.js` — History API processing + backfill
- `AutoLabel.js` — rule-based auto-labeling
- `Triggers.js` — `runEmailProcessing()` is the 1-min entry point (auto-label then history API)

## Statuses

Queue items use: `sent` (webhook OK), `logged` (no webhook), `error` (webhook failed).

## Testing

No automated tests. Test by:
1. `clasp push --force` to a test sheet
2. Gmail Tools → Run Processing Now
3. Gmail Tools → Webhook → Test Webhook
4. Check Queue sheet for results
