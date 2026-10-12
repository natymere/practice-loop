# Practice Loop

A private spaced-repetition tracker built with Google Sheets and Apps Script. Save practice questions, record attempts, and review them at intervals of 1, 3, 7, 14, 30, 60, and 90 days. Includes a question bank, spreadsheet imports, notes, and optional email or Telegram reminders.

The browser interface and server live in `src`. Local tests live in `tests`, commit checks in `scripts`, and the blank import template in `templates`. Only the five files allowed by `.claspignore` upload to Apps Script.

Use Node 20 or later; `.nvmrc` selects Node 24. Install dependencies and enable the local commit guard:

```sh
npm ci --ignore-scripts
git config core.hooksPath .githooks
npm test
```

The focused tests use simulated Google services. `npm run test:all` runs the broader suite when needed.

To connect a new checkout, copy `.clasp.example.json` to `.clasp.json` and set the ID of your existing Apps Script project. Keep `rootDir` set to `src`. Enable the Apps Script API in [Google Apps Script settings](https://script.google.com/home/usersettings), then run:

```sh
npm run source:login   # Sign in with the project owner account
npm run source:status  # Preview the five upload files
npm run source:push    # Run focused tests and upload source
```

Run uploads in an interactive terminal so you can review any manifest replacement prompt. Keep `executeAs: USER_DEPLOYING` and `access: MYSELF`. Do not use force or watch mode. Uploading source can affect editor runs and triggers, but does not update a versioned web app or run the data migration.

For an existing installation that needs the question-bank update:

1. Back up the spreadsheet and close open app tabs.
2. Upload the source, then run `upgrade` in the Apps Script editor as the owner. It preserves questions and review history, migrates paused questions to the bank, and creates the Import tab. It is safe to repeat.
3. Edit the existing web app under Deploy → Manage deployments. Choose New version, Execute as Me, and Only myself access. Keep the existing deployment URL.
4. Reload the app. Complete migration and deployment together; the older code may reject the new Sheet columns. Any rollback must restore matching code and data.

Source was uploaded on 12 October 2026. The migration and versioned web app update were still pending at that time.

Telegram credentials belong only in Apps Script Script Properties as `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`. Keep authentication files, real question lists, and spreadsheet exports out of Git. The owner email and Google resource identifiers in source refer to the existing private installation; they are not authentication credentials.

Git excludes agent notes and Markdown files except this README. The local commit guard checks staged files for common credential patterns and requires the private deployment manifest. GitHub stores the code; pushing to GitHub does not upload or deploy it to Apps Script.
