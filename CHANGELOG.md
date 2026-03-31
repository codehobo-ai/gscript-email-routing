# Changelog

## v4.0.0 (2026-03-31)

Merged v1 + v3 into a single unified version.

### From v3
- Gmail History API replaces label polling (only processes newly labeled emails)
- Auto-label rules via Label Rules sheet (Gmail search syntax, priority, stop-on-match)
- Dry run + backfill for auto-label rules
- Sheet formatting: conditional rules, column widths, date formats, row banding
- Webhook auth headers (X-API-Key, Bearer token, custom)
- Webhook returns `{ success, executionId, error }` — statuses are now `sent/logged/error`
- Per-label `mark_read` and `archive` controls
- Backfill function for pre-existing labeled emails
- History tracker reset UI
- Explicit OAuth scopes in appsscript.json

### From v1 improvements
- Guided setup wizard (step-by-step prompts after initial setup)
- Resume Setup Wizard menu item (detects current state, picks up where you left off)
- Test Webhook button (sends test payload, reports result)
- Configure webhook shows current URL before prompting
- Cached dedup set in processNewEmails (avoids re-reading queue sheet per message)
- Single Gmail API call in scheduledSync (shared across syncLabelNames + discoverNewLabels)
- Queue headers defined once in CONFIG.queueHeaders
- viewTriggers shows ui.alert instead of only logging
- Removed legacy dedup menu item

### New in v4
- Version check on sheet open (fetches version.json from GitHub, shows toast if update available)
- deploy.sh for pushing to multiple Google Apps Script instances
- README, CLAUDE.md, CHANGELOG.md
- .gitignore, .clasp.json.example, deploy.sh.example
