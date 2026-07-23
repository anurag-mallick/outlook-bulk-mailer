# Outlook Bulk Mailer

A point-and-click Windows tool that sends a personalized email, one at a
time, to every recipient in an Excel file — through your own locally
installed Microsoft Outlook. Each recipient sees their own name, company, or
any other detail you add as a column in the spreadsheet. No coding required
to use it.

## Quick start (no technical knowledge needed)

1. Download this project and open its folder.
2. **Double-click `Send Bulk Mail.bat`.** That's it — the first time you run
   it, it quietly installs everything it needs (a minute or two, with
   messages on screen so you know it's working), then opens the tool. Every
   time after that, it just opens straight away.
3. In the window that opens, follow Steps 1 through 6 from top to bottom. A
   green message near the top always tells you exactly what to do next.

That one file is the only thing you ever need to double-click. If step 2
ever fails, see [Troubleshooting](#troubleshooting) below.

## Features

- **Mail-merge from Excel** — load a spreadsheet, pick the email (and
  optional CC) column, and insert any other column as a `<FieldName>` token
  anywhere in the subject or body.
- **Rich or plain email** — a plain-text mode, or a rich mode with a
  formatting toolbar (bold, italic, underline, colors, alignment, bulleted
  lists, inline tables, and inline images placed directly in the message).
- **Attachments** — attach one or more files to every email sent.
- **Resumable sends** — if the tool is closed and reopened mid-campaign,
  recipients already emailed are not emailed again.
- **Delivery logging** — every send is recorded to a timestamped log file
  with the outcome (and reason, if it failed) next to each recipient.
- **Guided, ordered UI** — the window walks through six steps top to
  bottom; each step's controls stay disabled until the step before it is
  complete, so it's difficult to send something malformed or out of order.

## Requirements

- Windows, with Microsoft Outlook installed and configured with an account.
- Python 3.9+ — if it's missing, `Send Bulk Mail.bat` tells you exactly
  where to get it and stops cleanly; just run it again afterward.

## Using your own recipient list (once you've tried the quick start)

1. **Try it safely first, with `test_data.xlsx`.** Open it, replace the
   placeholder addresses with one or two of your own, and load that file in
   Step 1 the first time — so you see exactly what the email looks like
   before sending to anyone real.
2. **Build your real recipient list.** Copy `recipients_template.xlsx`,
   read the "Instructions" sheet inside it, and fill in the "Recipients" and
   "Custom Fields" sheets. Save it under a new name in this folder.
3. **Set your sending account as Outlook's default.** File > Account
   Settings > Account Settings > select the account > "Set as Default". On
   some Outlook/Exchange setups, mail is sent from whichever account is set
   as default regardless of the in-app selection — this is the only
   reliable way to control the "from" address.
4. Load that file in Step 1 instead of the test data, and continue through
   Steps 2-6: choose the email/CC columns, insert any field tokens you want,
   write the subject and message, attach files if needed, then connect to
   Outlook and send.

Emails are sent using only the Excel file loaded at the moment you click
Send — if you edit that file afterward, click "Load Recipients" again before
sending.

## Troubleshooting

`Send Bulk Mail.bat` sets everything up automatically the first time it
runs. If it ever reports that setup failed:

- Check your internet connection, or ask IT if a company proxy is blocking
  Python package installs.
- Try running `setup.bat` directly — it performs the same install with more
  detailed messages, and can also be used to force a clean reinstall if
  something gets into a bad state.
- If it still fails, share the on-screen error message with your IT support.

## Files

| File | Purpose |
|------|---------|
| `bulk_mailer_gui.py` | The application itself |
| `Send Bulk Mail.bat` | The only file you need to double-click — sets up on first run, then launches |
| `setup.bat` | Manual setup/reinstall, for troubleshooting only |
| `recipients_template.xlsx` | Copy this to build your real recipient list |
| `test_data.xlsx` | Placeholder data for a safe first test run |
| `sent_log_*.txt` | Created automatically per send campaign (git-ignored) |

## License

MIT — see [LICENSE](LICENSE).
