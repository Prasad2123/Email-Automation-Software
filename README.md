# 📧 Email Automation Software

> A configurable batch email automation system built with Python — send emails in smart batches with countdown timers, Excel-based recipient lists, dry-run previews, and campaign state management.

---

## ✨ Features

- 📋 **Excel-based recipients** — load email addresses from `.xlsx` files
- 🔁 **Batch sending** — send N emails, pause, then continue automatically
- ⏱️ **Live countdown timer** — watch the break timer tick down in real time
- 🧪 **Dry-run mode** — preview exactly what will be sent without sending anything
- 💾 **Campaign state** — resumes safely if interrupted; no duplicate sends
- ✅ **Correct counting** — only successful sends count toward the batch limit
- 🔒 **Secure config** — credentials stay in `config.cfg` (gitignored)
- 🖥️ **Clean CLI** — all settings overridable from the command line

---

## 📁 Project Structure

```
Email automation/
│
├── run_automail.py              ← Main script (entry point)
├── config.cfg                   ← Your credentials & settings (gitignored)
├── config.cfg.example           ← Template — copy this to config.cfg
├── requirements.txt             ← Python dependencies
├── template.html                ← Optional HTML email template
├── contacts.csv                 ← Legacy contacts (not used in batch mode)
│
├── sheets/
│   └── test_email_recipients.xlsx   ← Excel file with recipient emails
│
├── state/
│   └── email_campaign_state.json    ← Auto-created progress file (gitignored)
│
└── pyautomail/                  ← Bundled PyAutoMail library
    └── pyautomail/
        └── emailsender.py
```

---

## ⚙️ Configuration

### 1. Copy the example config

```bash
cp config.cfg.example config.cfg
```

### 2. Edit `config.cfg`

```ini
[smtp]
host = smtp.gmail.com
port = 465
is_test = False

[account]
user = your_email@gmail.com
password = your_app_password_here   # Gmail App Password (16 chars)

[log]
file-path = email_sender.log
level = 20

[batch]
batch_size = 3        # How many emails to send per batch
wait_minutes = 5      # How many minutes to wait between batches

[email]
subject = Email Automation Test
body = hello
excel_file = sheets/test_email_recipients.xlsx
```

> **Gmail App Password** — Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) to generate one. Enable 2-Factor Authentication first.

---

## 📊 Excel File Format

Create your recipient list in `sheets/` as an `.xlsx` file with a column named **`email`**:

| email |
|---|
| alice@example.com |
| bob@example.com |
| carol@example.com |

- Blank cells are skipped automatically
- Duplicate addresses are de-duplicated
- Invalid email formats are skipped with a warning

---

## 🚀 Installation

```bash
# 1. Install Python dependencies
pip install pandas openpyxl

# 2. Install the bundled PyAutoMail library
pip install -e pyautomail/

# 3. Copy and fill in your config
cp config.cfg.example config.cfg
# → Edit config.cfg with your Gmail credentials
```

---

## 🖥️ Usage

### Dry Run — preview without sending
```bash
python run_automail.py --dry-run
```

```
====================================================
[DRY RUN] - No emails will be sent
====================================================

  Batch 1
  -------
    1. alice@example.com
    2. bob@example.com
    3. carol@example.com

    --> Would wait 5.0 minute(s) before next batch

  [DRY RUN] No emails were sent. No waiting occurred.
```

### Real Send — uses config.cfg defaults
```bash
python run_automail.py
```

### Fresh Start — ignore previous campaign progress
```bash
python run_automail.py --reset
```

### Override settings from CLI
```bash
python run_automail.py \
  --excel sheets/my_list.xlsx \
  --batch-size 10 \
  --wait-minutes 15 \
  --subject "Monthly Newsletter"
```

---

## 📺 Live Console Output

```
====================================================
EMAIL AUTOMATION STARTED
====================================================

  Total emails in Excel : 6
  Batch size            : 3
  Wait time             : 5.0 minute(s)
  Subject               : Email Automation Test
  Body                  : hello

  Starting campaign...

  1 email sent [OK] -> alice@example.com
  2 email sent [OK] -> bob@example.com
  3 email sent [OK] -> carol@example.com
====================================================
  3 EMAILS SENT
  Taking a break for 5.0 minute(s)...
====================================================
    Waiting: 05:00
    Waiting: 04:59
    ...
    Waiting: 00:01
====================================================
  BREAK COMPLETED
  Starting next batch...
====================================================

  4 email sent [OK] -> dave@example.com
  5 email sent [OK] -> eve@example.com
  6 email sent [OK] -> frank@example.com

====================================================
  ALL 6 EMAILS PROCESSED
  Total records found  : 6
  Successfully sent    : 6
  Failed               : 0
  Remaining            : 0

  Campaign completed successfully.
  Automation stopped.
====================================================
```

---

## 🔢 How Batch Counting Works

The system counts **successful sends only**. Failed emails do not consume a batch slot.

```
batch_sent = 0   ← resets after each batch
total_sent = 0   ← cumulative, never resets

For each recipient:
    → attempt send
    → if FAILED: log it, skip (do NOT increment)
    → if SUCCESS:
        total_sent += 1
        batch_sent += 1
        if batch_sent >= batch_size AND more remain:
            start countdown
            batch_sent = 0
```

**Example with a failure:**
```
1 email sent [OK] -> alice@example.com
2 email FAILED [!] -> bad-address          ← not counted
2 email sent [OK] -> bob@example.com       ← still #2
3 email sent [OK] -> carol@example.com
====================================================
  3 EMAILS SENT — Taking a break...
```

---

## ⚙️ CLI Reference

| Argument | Default | Description |
|---|---|---|
| `--config` | `config.cfg` | Path to config file |
| `--excel` | from config | Path to Excel recipients file |
| `--batch-size` | from config | Emails per batch |
| `--wait-minutes` | from config | Minutes between batches |
| `--subject` | from config | Email subject line |
| `--body` | from config | Email body text |
| `--dry-run` | off | Preview only, no sending |
| `--reset` | off | Clear saved state, start fresh |

> CLI arguments always override `config.cfg` values.

---

## 📁 Changing Settings

All key values live in `config.cfg` — **no code changes needed**.

| What to change | Setting |
|---|---|
| Emails per batch | `batch_size` in `[batch]` |
| Wait time | `wait_minutes` in `[batch]` |
| Input Excel file | `excel_file` in `[email]` |
| Email subject | `subject` in `[email]` |
| Email body | `body` in `[email]` |
| SMTP server | `host` / `port` in `[smtp]` |

---

## 🔒 Security

- `config.cfg` is **gitignored** — credentials are never committed
- `state/*.json` is **gitignored** — campaign progress is never committed
- Never log passwords — the logger only records email addresses and statuses
- Use **Gmail App Passwords**, not your main Gmail password

> ⚠️ If you ever accidentally commit credentials, revoke them immediately at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) and generate a new one.

---

## 📦 Dependencies

```
pandas       — Excel file reading
openpyxl     — .xlsx engine for pandas
jinja2       — Email template rendering (via PyAutoMail)
```

Install with:
```bash
pip install pandas openpyxl
```

---

## 🧾 Logs

Every campaign writes to `email_sender.log` (gitignored):

```
[2026-09-12 12:09:29 - INFO (EmailSender)] : Logged in.
[2026-09-12 12:09:29 - INFO (BatchRunner)] : Campaign sending started.
[2026-09-12 12:09:31 - INFO (BatchRunner)] : [1/6] Sent to alice@example.com
[2026-09-12 12:09:33 - INFO (BatchRunner)] : Batch of 3 complete. Waiting 5 minute(s).
[2026-09-12 12:14:33 - INFO (BatchRunner)] : Batch wait finished. Resuming sending.
```

---

## 📄 License

MIT — free to use and modify.
