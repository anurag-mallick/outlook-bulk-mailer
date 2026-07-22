# Outlook Bulk Mailer

A point-and-click Windows tool that sends a personalized email, one at a
time, to every recipient in an Excel file — through your own locally
installed Microsoft Outlook. Each recipient sees their own name, company, or
any other detail you add as a column in the spreadsheet. No coding required
to use it.

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
- Python 3.9+.

## Installation

```powershell
git clone <this-repo-url>
cd "Outlook Bulk Mailer"
.\setup.bat
```

`setup.bat` checks for Python, installs the dependencies in
`requirements.txt` (`pywin32`, `pandas`, `openpyxl`), and registers `pywin32`
with Windows (required for Outlook automation).

## Usage

1. **Try it safely first.** Open `test_data.xlsx`, replace the placeholder
   addresses with one or two of your own, and use it as your Excel file the
   first time through the steps below — so you see exactly what the email
   looks like before sending to anyone real.
2. **Build your real recipient list.** Copy `recipients_template.xlsx`,
   read the "Instructions" sheet inside it, and fill in the "Recipients" and
   "Custom Fields" sheets. Save it under a new name in this folder.
3. **Set your sending account as Outlook's default.** File > Account
   Settings > Account Settings > select the account > "Set as Default". On
   some Outlook/Exchange setups, mail is sent from whichever account is set
   as default regardless of the in-app selection — this is the only
   reliable way to control the "from" address.
4. **Double-click `Send Bulk Mail.bat`.** Follow the six steps in the
   window: load the Excel file, choose the email/CC columns, insert any
   field tokens you want, write the subject and message, attach files if
   needed, then connect to Outlook and send.

Emails are sent using only the Excel file loaded at the moment you click
Send — if you edit that file afterward, click "Load Recipients" again before
sending.

## Files

| File | Purpose |
|------|---------|
| `bulk_mailer_gui.py` | The application itself |
| `Send Bulk Mail.bat` | Double-click to launch |
| `setup.bat` | One-time dependency installation |
| `recipients_template.xlsx` | Copy this to build your real recipient list |
| `test_data.xlsx` | Placeholder data for a safe first test run |
| `sent_log_*.txt` | Created automatically per send campaign (git-ignored) |

## License

MIT — see [LICENSE](LICENSE).
