# 📧 Email Automation Software

> A powerful, configurable batch email automation system built with Python — send emails in smart batches with live countdown timers, Excel recipient lists, image & file attachment support, dry-run previews, and robust campaign state management.

---

## ✨ Features

- 📋 **Excel-based recipients** — load email addresses dynamically from `.xlsx` spreadsheets
- 📎 **Attachment & Image Support** — attach images (`.png`, `.jpg`), documents, or PDFs to every outgoing email
- 🔁 **Batch sending** — send N emails, pause, then continue automatically to avoid spam filters and rate limits
- ⏱️ **Live countdown timer** — watch the break timer tick down second-by-second in your terminal
- 🧪 **Dry-run mode** — preview batches and attachment paths without sending any actual emails
- 💾 **Campaign state management** — resumes safely from where it left off if interrupted; zero duplicate sends
- ✅ **Accurate counting** — only successful sends count toward the batch limit; failures are logged and retried/skipped cleanly
- 🔒 **Secure configuration** — personal credentials stay safely in `config.cfg` (gitignored)
- 🖥️ **CLI overrides** — override any configuration setting directly from command-line arguments

---

## 📁 Project Structure

```
Email automation/
│
├── run_automail.py              ← Main script (entry point with batching & attachment support)
├── config.cfg                   ← Your credentials & settings (gitignored)
├── config.cfg.example           ← Template — copy this to config.cfg
├── requirements.txt             ← Python dependencies
│
├── assets/                      ← Folder for email attachments (images, PDFs, documents)
│   └── pic1.png                 ← Example image attachment
│
├── sheets/
│   └── test_email_recipients.xlsx   ← Excel file with recipient emails
│
├── state/
│   └── email_campaign_state.json    ← Auto-created progress tracker (gitignored)
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
*(On Windows PowerShell, you can use `copy config.cfg.example config.cfg`)*

### 2. Edit `config.cfg`

```ini
[smtp]
host = smtp.gmail.com
port = 465
is_test = False

[account]
user = your_email@gmail.com
password = your_app_password_here   # Gmail App Password (16 chars, no spaces required)

[log]
file-path = email_sender.log
level = 20

[batch]
batch_size = 3        # How many emails to send per batch
wait_minutes = 2      # How many minutes to wait between batches

[email]
subject = Email Automation Test
body = hello
excel_file = sheets/test_email_recipients.xlsx
attachment = assets/pic1.png         # Optional: path to image or document attachment
```

> **🔑 How to generate a Gmail App Password:**
> 1. Go to your Google Account: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
> 2. Ensure 2-Step Verification is turned **ON**
> 3. Create an App Name (e.g. `Email Automation`)
> 4. Copy the generated 16-character password into `password` in `config.cfg`

---

## 📊 Excel File Format

Place your recipient list in `sheets/` as an `.xlsx` file. Ensure there is a column named **`email`**:

| email |
|---|
| alice@example.com |
| bob@example.com |
| carol@example.com |

- Empty rows or blank cells are automatically ignored
- Duplicate email addresses are de-duplicated
- Invalid email addresses are reported with a warning and skipped

---

## 🚀 Installation & Setup

```bash
# 1. Install required Python packages
pip install pandas openpyxl

# 2. Install the bundled PyAutoMail library in editable mode
pip install -e pyautomail/

# 3. Create your local config file
copy config.cfg.example config.cfg
# -> Fill in your email and App Password in config.cfg
```

---

## 🖥️ Usage

### 1. Dry Run — preview without sending emails
Verify your recipient list, batches, and attachment without sending anything:
```bash
python run_automail.py --dry-run
```

```
====================================================
[DRY RUN] - No emails will be sent
====================================================

  Total emails in Excel : 6
  Batch size            : 3
  Wait time             : 2.0 minute(s)
  Subject               : Email Automation Test
  Attachment            : assets/pic1.png

  Batch 1
  -------
    1. alice@example.com
    2. bob@example.com
    3. carol@example.com

    --> Would wait 2.0 minute(s) before next batch

  Batch 2
  -------
    4. dave@example.com
    5. eve@example.com
    6. frank@example.com

  [DRY RUN] No emails were sent. No waiting occurred.
```

---

### 2. Live Send — start the automated campaign
```bash
python run_automail.py
```

### 3. Fresh Restart — clear previous state
If you stopped the software previously and want to start over from the very first recipient:
```bash
python run_automail.py --reset
```

### 4. Sending with Custom Attachment via CLI
You can specify or override the attachment at runtime:
```bash
python run_automail.py --attachment assets/pic1.png
```

### 5. Override Any Settings from CLI
```bash
python run_automail.py \
  --excel sheets/my_recipients.xlsx \
  --batch-size 5 \
  --wait-minutes 10 \
  --subject "Special Update" \
  --attachment assets/brochure.pdf
```

---

## 📺 Live Console Preview

When running, the software displays formatted live progress and a real-time countdown timer between batches:

```
====================================================
EMAIL AUTOMATION STARTED
====================================================

  Total emails in Excel : 6
  Batch size            : 3
  Wait time             : 2.0 minute(s)
  Subject               : Email Automation Test
  Body                  : hello
  Attachment            : assets/pic1.png

  Starting campaign...

  1 email sent [OK] -> alice@example.com
  2 email sent [OK] -> bob@example.com
  3 email sent [OK] -> carol@example.com
====================================================
  3 EMAILS SENT
  Taking a break for 2.0 minute(s)...
====================================================
    Waiting: 02:00
    Waiting: 01:59
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

The system increments batch counters **only upon confirmed successful delivery**:

```
batch_sent = 0   ← resets to 0 after each batch pause
total_sent = 0   ← cumulative counter across all batches

For each recipient:
    → Attempt SMTP delivery (with attachment if configured)
    → If FAILED:
        Log error, do NOT increment batch_sent
    → If SUCCESS:
        total_sent += 1
        batch_sent += 1
        Record progress to state/email_campaign_state.json
        If batch_sent >= batch_size AND more recipients remain:
            Trigger countdown timer for wait_minutes
            batch_sent = 0
```

---

## ⚙️ CLI Options Reference

| Argument | Default | Description |
|---|---|---|
| `--config` | `config.cfg` | Path to custom `.cfg` configuration file |
| `--excel` | from config | Path to Excel spreadsheet containing recipients |
| `--batch-size` | from config | Number of emails to send before pausing |
| `--wait-minutes` | from config | Pause duration in minutes between batches |
| `--subject` | from config | Subject line for outgoing emails |
| `--body` | from config | Plain text body content |
| `--attachment` | from config | Path to attachment file (e.g. `assets/pic1.png`, PDF, etc.) |
| `--dry-run` | `False` | Simulate the run without sending or waiting |
| `--reset` | `False` | Clear saved campaign state and start from recipient 1 |

> 💡 *Command-line options always take precedence over values in `config.cfg`.*

---

## 🔧 Quick Configuration Reference

All settings can be permanently adjusted in `config.cfg`:

| What you want to change | Location in `config.cfg` |
|---|---|
| Number of emails per batch | `batch_size` under `[batch]` |
| Break duration between batches | `wait_minutes` under `[batch]` |
| Excel recipient list | `excel_file` under `[email]` |
| Email subject | `subject` under `[email]` |
| Email body text | `body` under `[email]` |
| Image or file attachment | `attachment` under `[email]` |
| Gmail address | `user` under `[account]` |
| Gmail App Password | `password` under `[account]` |
| SMTP Host & Port | `host` / `port` under `[smtp]` |

---

## 🔒 Security & Privacy

- `config.cfg` is added to `.gitignore` — your Gmail username and App Password will never be pushed to git.
- `state/*.json` is gitignored to avoid leaking recipient progress or states.
- `*.log` files are gitignored.
- Passwords are never written to log files.
- Always use Google **App Passwords** rather than your main Google account password.

---

## 📦 Requirements

```txt
pandas>=2.0.0
openpyxl>=3.1.0
```

---

## 📄 License

This project is licensed under the MIT License.
