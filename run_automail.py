"""
Batch Email Automation Runner
==============================
Sends emails in configurable batches with a wait period between batches.

Usage examples:
    # Dry run (shows what would be sent, no actual emails)
    python run_automail.py --dry-run

    # Real send using config.cfg defaults
    python run_automail.py

    # Override batch settings from the command line
    python run_automail.py --excel sheets/test_email_recipients.xlsx --batch-size 3 --wait-minutes 5

    # Reset campaign state (ignore previous progress)
    python run_automail.py --reset

Configuration (config.cfg):
    [batch]
    batch_size = 3
    wait_minutes = 5

    [email]
    subject = Email Automation Test
    body = hello
    excel_file = sheets/test_email_recipients.xlsx
"""

import os
import sys
import json
import time
import logging
import argparse
import re
from configparser import ConfigParser
from datetime import datetime, timezone

# Ensure the local pyautomail package is importable.
# The library lives at pyautomail/pyautomail/ relative to this script.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_LIB_PATH = os.path.join(_SCRIPT_DIR, "pyautomail")
if _LIB_PATH not in sys.path:
    sys.path.insert(0, _LIB_PATH)

import pandas as pd
from pyautomail import EmailSender

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
STATE_DIR = "state"
STATE_FILE = os.path.join(STATE_DIR, "email_campaign_state.json")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sep(char="=", width=52):
    return char * width


def _print_sep(char="=", width=52):
    print(_sep(char, width))


def _is_valid_email(addr: str) -> bool:
    """Basic RFC 5322 regex check."""
    pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    return bool(re.match(pattern, addr))


def _load_config(config_path: str) -> ConfigParser:
    """Read and return the ConfigParser object."""
    if not os.path.exists(config_path):
        print(f"\n[ERROR] Configuration file not found: '{config_path}'")
        print("        Copy 'config.cfg.example' to 'config.cfg' and fill in your credentials.")
        sys.exit(1)
    cfg = ConfigParser()
    cfg.read(config_path)
    return cfg


def _validate_batch_config(batch_size: int, wait_minutes: float):
    """Validate batch configuration values."""
    if batch_size <= 0:
        print(f"\n[ERROR] Invalid batch_size={batch_size}. Must be a positive integer (> 0).")
        sys.exit(1)
    if wait_minutes < 0:
        print(f"\n[ERROR] Invalid wait_minutes={wait_minutes}. Must be zero or greater.")
        sys.exit(1)


def _load_recipients(excel_path: str) -> list[str]:
    """
    Load, validate, de-duplicate, and return recipient email addresses from an Excel file.

    Returns a list of clean, unique, valid email addresses.
    """
    if not os.path.exists(excel_path):
        print(f"\n[ERROR] Excel file not found: '{excel_path}'")
        sys.exit(1)

    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as exc:
        print(f"\n[ERROR] Could not read Excel file '{excel_path}': {exc}")
        sys.exit(1)

    if "email" not in df.columns:
        print(f"\n[ERROR] Excel file must contain a column named 'email'.")
        print(f"        Found columns: {list(df.columns)}")
        sys.exit(1)

    raw = df["email"].tolist()
    seen = set()
    valid = []
    skipped_blank = 0
    skipped_invalid = 0
    skipped_duplicate = 0

    for cell in raw:
        addr = str(cell).strip() if cell is not None else ""
        if addr in ("", "nan", "None"):
            skipped_blank += 1
            continue
        if not _is_valid_email(addr):
            print(f"    [WARN] Skipping invalid email: '{addr}'")
            skipped_invalid += 1
            continue
        if addr in seen:
            print(f"    [WARN] Skipping duplicate: '{addr}'")
            skipped_duplicate += 1
            continue
        seen.add(addr)
        valid.append(addr)

    if skipped_blank:
        print(f"    Skipped {skipped_blank} blank cell(s).")
    if skipped_invalid:
        print(f"    Skipped {skipped_invalid} invalid email(s).")
    if skipped_duplicate:
        print(f"    Skipped {skipped_duplicate} duplicate(s).")

    if not valid:
        print(f"\n[ERROR] No valid email addresses found in '{excel_path}'.")
        sys.exit(1)

    return valid


# ---------------------------------------------------------------------------
# State management
# ---------------------------------------------------------------------------

def _load_state(excel_path: str) -> dict | None:
    """Load existing campaign state if it matches the current excel file."""
    if not os.path.exists(STATE_FILE):
        return None
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
        if state.get("excel_file") == excel_path:
            return state
    except Exception:
        pass
    return None


def _save_state(state: dict):
    """Persist campaign state to disk."""
    os.makedirs(STATE_DIR, exist_ok=True)
    state["last_updated"] = datetime.now(timezone.utc).isoformat()
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def _clear_state():
    """Delete the state file to start a fresh campaign."""
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        print("    Campaign state cleared. Starting fresh.\n")


# ---------------------------------------------------------------------------
# Countdown display
# ---------------------------------------------------------------------------

def _countdown(wait_minutes: float, logger: logging.Logger):
    """
    Show a live countdown in the console while waiting between batches.
    Logs only start and end (not every second) to avoid log file noise.
    """
    total_seconds = int(wait_minutes * 60)
    logger.info(f"Batch wait started: {wait_minutes} minute(s) ({total_seconds}s)")

    for remaining in range(total_seconds, 0, -1):
        mins, secs = divmod(remaining, 60)
        sys.stdout.write(f"\r    Waiting: {mins:02d}:{secs:02d}  ")
        sys.stdout.flush()
        time.sleep(1)

    sys.stdout.write("\r" + " " * 30 + "\r")  # clear the line
    sys.stdout.flush()
    logger.info("Batch wait finished. Resuming sending.")


# ---------------------------------------------------------------------------
# Dry-run display
# ---------------------------------------------------------------------------

def _dry_run(recipients: list[str], batch_size: int, wait_minutes: float,
             excel_path: str, subject: str, body: str):
    """Print the dry-run plan with batch boundaries."""
    _print_sep()
    print("[DRY RUN] - No emails will be sent")
    _print_sep()
    print(f"\n  Input file       : {excel_path}")
    print(f"  Total recipients : {len(recipients)}")
    print(f"  Batch size       : {batch_size}")
    print(f"  Wait time        : {wait_minutes} minute(s)")
    print(f"  Subject          : {subject}")
    print(f"  Body             : {body}\n")

    batch_num = 1
    for i, addr in enumerate(recipients):
        pos_in_batch = (i % batch_size) + 1
        if pos_in_batch == 1:
            _print_sep("-", 52)
            print(f"  Batch {batch_num}")
            _print_sep("-", 52)

        print(f"    {i + 1}. {addr}")

        if pos_in_batch == batch_size and i < len(recipients) - 1:
            print(f"\n    --> Would wait {wait_minutes} minute(s) before next batch\n")
            batch_num += 1

    print()
    _print_sep()
    print("  [DRY RUN] No emails were sent. No waiting occurred.")
    _print_sep()


# ---------------------------------------------------------------------------
# Main campaign
# ---------------------------------------------------------------------------

def run_campaign(
    config_path: str,
    excel_path: str,
    batch_size: int,
    wait_minutes: float,
    subject: str,
    body: str,
    attachment: str,
    dry_run: bool,
    reset: bool,
    logger: logging.Logger,
):
    # ---- Header ----
    _print_sep()
    print("EMAIL AUTOMATION STARTED")
    _print_sep()

    logger.info("=" * 52)
    logger.info("Email Automation Campaign Started")
    logger.info(f"Excel file     : {excel_path}")
    logger.info(f"Batch size     : {batch_size}")
    logger.info(f"Wait minutes   : {wait_minutes}")
    logger.info(f"Subject        : {subject}")
    logger.info(f"Attachment     : {attachment or 'None'}")
    logger.info(f"Dry run        : {dry_run}")

    # ---- Load & validate recipients ----
    recipients = _load_recipients(excel_path)
    total_recipients = len(recipients)

    # ---- Validate attachment if provided ----
    if attachment:
        if not os.path.exists(attachment):
            print(f"\n[ERROR] Attachment file not found: '{attachment}'")
            logger.error(f"Attachment not found: {attachment}")
            sys.exit(1)
        logger.info(f"Attachment verified: {attachment}")

    print(f"\n  Total emails in Excel : {total_recipients}")
    print(f"  Batch size            : {batch_size}")
    print(f"  Wait time             : {wait_minutes} minute(s)")
    print(f"  Subject               : {subject}")
    print(f"  Body                  : {body}")
    print(f"  Attachment            : {attachment if attachment else 'None'}\n")

    logger.info(f"Total recipients loaded: {total_recipients}")

    # ---- Dry run shortcut ----
    if dry_run:
        _dry_run(recipients, batch_size, wait_minutes, excel_path, subject, body)
        return

    # ---- Resume / reset state ----
    if reset:
        _clear_state()

    state = _load_state(excel_path) if not reset else None
    already_sent_set: set[str] = set()

    if state:
        already_sent_set = set(state.get("sent", []))
        if already_sent_set:
            print(f"    [RESUME] Found previous campaign state.")
            print(f"             Already sent: {len(already_sent_set)} email(s). Skipping those.\n")
            logger.info(f"Resuming campaign. Already sent: {already_sent_set}")

    # Build ordered list of recipients still to send
    remaining = [r for r in recipients if r not in already_sent_set]

    if not remaining:
        print("  All recipients in the Excel file have already been sent to.")
        print("  Use --reset to start a fresh campaign.")
        logger.info("No remaining recipients. Campaign already complete.")
        return

    # ---- Initialize sender ----
    print("  Initialising email sender...\n")
    try:
        sender = EmailSender(cfg=config_path)
        sender.set_template(plain_temp=body)
    except SystemExit:
        # EmailSender calls exit(1) on auth failure — give a helpful message first
        print("\n[ERROR] Failed to authenticate with the SMTP server.")
        print("        Check [account] user and password in config.cfg.")
        logger.error("SMTP authentication failed. Campaign aborted.")
        sys.exit(1)
    except Exception as exc:
        print(f"\n[ERROR] Could not initialise email sender: {exc}")
        logger.error(f"EmailSender init failed: {exc}")
        sys.exit(1)

    # ---- Campaign state setup ----
    state = {
        "excel_file": excel_path,
        "total_recipients": total_recipients,
        "sent": list(already_sent_set),
        "failed": state.get("failed", []) if state else [],
        "last_updated": "",
    }

    # ---- Campaign counters ----
    total_sent = len(already_sent_set)   # honour previously sent count
    total_failed = 0
    batch_sent = 0                        # resets after each batch completes

    global_num = total_sent              # sequential success counter (display)

    print("  Starting campaign...\n")
    logger.info("Campaign sending started.")

    remaining_iter = list(remaining)  # snapshot for iteration

    for idx, recipient in enumerate(remaining_iter):
        # Determine if more recipients follow *in the remaining list*
        more_after = idx < len(remaining_iter) - 1

        # ---- Attempt send ----
        logger.info(f"Attempting to send to: {recipient}")

        try:
            result = sender.send(recipient, subject, {},
                                 attachment_path=attachment if attachment else None)
        except Exception as exc:
            result = {"err": str(exc)}

        # ---- Check for failure in returned dict ----
        err_val = result.get("err", "none") if isinstance(result, dict) else "none"
        send_failed = (
            isinstance(result, dict)
            and err_val not in ("none", "", None)
            and err_val != "none"
        )
        # Also check if result itself signals test mode (still counts as sent)
        is_test_mode = isinstance(result, dict) and result.get("test") is True

        if send_failed and not is_test_mode:
            total_failed += 1
            state["failed"].append(recipient)
            _save_state(state)
            print(f"  {global_num + 1} email FAILED [!] -> {recipient}  ({err_val})")
            logger.error(f"Failed to send to {recipient}: {err_val}")
            # Do NOT increment counters — continue to next recipient
            continue

        # ---- Success ----
        global_num += 1
        total_sent += 1
        batch_sent += 1
        state["sent"].append(recipient)
        _save_state(state)

        print(f"  {global_num} email sent [OK] -> {recipient}")
        logger.info(f"[{global_num}/{total_recipients}] Sent to {recipient}")

        # ---- Batch boundary check ----
        if batch_sent >= batch_size and more_after:
            _print_sep()
            print(f"  {batch_sent} EMAILS SENT")
            print(f"  Taking a break for {wait_minutes} minute(s)...")
            _print_sep()
            logger.info(f"Batch of {batch_sent} complete. Waiting {wait_minutes} minute(s).")

            _countdown(wait_minutes, logger)

            _print_sep()
            print("  BREAK COMPLETED")
            print("  Starting next batch...")
            _print_sep()
            print()

            batch_sent = 0  # reset batch counter

    # ---- Final summary ----
    print()
    _print_sep()
    remaining_count = total_recipients - total_sent - total_failed
    print(f"  ALL {total_sent} EMAILS PROCESSED")
    print(f"  Total records found  : {total_recipients}")
    print(f"  Successfully sent    : {total_sent}")
    print(f"  Failed               : {total_failed}")
    print(f"  Remaining            : {max(remaining_count, 0)}")
    print()
    print("  Campaign completed successfully.")
    print("  Automation stopped.")
    _print_sep()

    logger.info("Campaign completed.")
    logger.info(f"Total sent: {total_sent} | Failed: {total_failed}")

    # Clean up state file on clean completion
    if total_failed == 0 and max(remaining_count, 0) == 0:
        _clear_state()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Batch Email Automation – sends emails in configurable batches."
    )
    parser.add_argument("--config", default="config.cfg",
                        help="Path to config file (default: config.cfg)")
    parser.add_argument("--excel",
                        help="Path to Excel recipients file (overrides config)")
    parser.add_argument("--batch-size", type=int,
                        help="Number of emails per batch (overrides config)")
    parser.add_argument("--wait-minutes", type=float,
                        help="Minutes to wait between batches (overrides config)")
    parser.add_argument("--subject",
                        help="Email subject (overrides config)")
    parser.add_argument("--body",
                        help="Email body text (overrides config)")
    parser.add_argument("--attachment",
                        help="Path to attachment file, e.g. assets/pic1.png (overrides config)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be sent without sending")
    parser.add_argument("--reset", action="store_true",
                        help="Ignore saved campaign state and start fresh")

    args = parser.parse_args()

    # ---- Load config ----
    cfg = _load_config(args.config)

    # ---- Resolve settings: CLI args override config file ----
    try:
        batch_size = args.batch_size if args.batch_size is not None \
            else cfg.getint("batch", "batch_size", fallback=3)
        wait_minutes = args.wait_minutes if args.wait_minutes is not None \
            else cfg.getfloat("batch", "wait_minutes", fallback=5.0)
        excel_path = args.excel or cfg.get("email", "excel_file",
                                           fallback="sheets/test_email_recipients.xlsx")
        subject = args.subject or cfg.get("email", "subject",
                                          fallback="Email Automation Test")
        body = args.body or cfg.get("email", "body", fallback="hello")
        attachment = args.attachment or cfg.get("email", "attachment", fallback="").strip() or None
    except Exception as exc:
        print(f"\n[ERROR] Could not read configuration: {exc}")
        sys.exit(1)

    # ---- Validate batch configuration ----
    _validate_batch_config(batch_size, wait_minutes)

    # ---- Set up application logger ----
    try:
        log_file = cfg.get("log", "file-path", fallback="email_sender.log")
        log_level = cfg.getint("log", "level", fallback=20)
    except Exception:
        log_file = "email_sender.log"
        log_level = 20

    logging.basicConfig(
        level=log_level,
        format="[%(asctime)s - %(levelname)s (%(name)s) ] : %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")],
    )
    logger = logging.getLogger("BatchRunner")

    # ---- Run ----
    run_campaign(
        config_path=args.config,
        excel_path=excel_path,
        batch_size=batch_size,
        wait_minutes=wait_minutes,
        subject=subject,
        body=body,
        attachment=attachment,
        dry_run=args.dry_run,
        reset=args.reset,
        logger=logger,
    )


if __name__ == "__main__":
    main()
