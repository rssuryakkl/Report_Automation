import os
import sys
import json
import logging
from datetime import date, datetime, time
import pandas as pd
from dotenv import load_dotenv

# Base directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_FILE = os.path.join(BASE_DIR, ".env")
def get_excel_path() -> str:
    for fname in ["Daily_report.xlsx", "daily_report.xlsx"]:
        candidate = os.path.join(BASE_DIR, fname)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(BASE_DIR, "daily_report.xlsx")

EXCEL_FILE = get_excel_path()
LOGS_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE = os.path.join(LOGS_DIR, "scheduler.log")
SENT_RECORD_FILE = os.path.join(LOGS_DIR, "sent_dates.json")

# Ensure logs dir exists
os.makedirs(LOGS_DIR, exist_ok=True)

# Configure dedicated logger for DailyReportScheduler (avoids polluting with external asyncio/server logs)
logger = logging.getLogger("DailyReportScheduler")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    _fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    _sh = logging.StreamHandler(sys.stdout)
    _sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(_fh)
    logger.addHandler(_sh)

# Load environment
load_dotenv(ENV_FILE, override=True)

from email_sender import generate_email_html, generate_email_plain, send_report_email

def format_time_str(val):
    if hasattr(val, "strftime"):
        return val.strftime("%I:%M %p")
    val_str = str(val).strip()
    try:
        dt = datetime.strptime(val_str, "%H:%M:%S")
        return dt.strftime("%I:%M %p")
    except Exception:
        pass
    return val_str

def is_already_sent_today(target_date_str: str) -> bool:
    if not os.path.exists(SENT_RECORD_FILE):
        return False
    try:
        with open(SENT_RECORD_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return target_date_str in data
    except Exception:
        return False

def mark_sent_today(target_date_str: str):
    data = {}
    if os.path.exists(SENT_RECORD_FILE):
        try:
            with open(SENT_RECORD_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    data[target_date_str] = datetime.now().isoformat()
    try:
        with open(SENT_RECORD_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        logger.warning(f"Could not update sent record file: {e}")

def run_scheduled_job(force: bool = False, use_latest: bool = False):
    logger.info("==========================================")
    logger.info("Starting Daily Report 7:00 PM Mail Dispatcher")
    logger.info("==========================================")

    # Check if we should pull latest Excel from GitHub cloud
    gh_token = os.getenv("GITHUB_TOKEN", "")
    gh_repo = os.getenv("GITHUB_REPO", "rssuryakkl/Report_Automation")
    gh_branch = os.getenv("GITHUB_BRANCH", "main")
    if gh_token.strip():
        try:
            from github_sync import pull_excel_from_github
            pull_ok, pull_msg = pull_excel_from_github(EXCEL_FILE, gh_repo, gh_token, gh_branch)
            if pull_ok:
                logger.info(f"Pulled latest Excel from cloud: {pull_msg}")
            else:
                logger.warning(f"Cloud pull notice: {pull_msg}")
        except Exception as e:
            logger.warning(f"Could not pull latest Excel from GitHub: {e}")

    if not os.path.exists(EXCEL_FILE):
        logger.error(f"Excel file not found at: {EXCEL_FILE}")
        return False

    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        logger.error(f"Failed to read {EXCEL_FILE}: {e}")
        return False

    if df.empty:
        logger.warning("daily_report.xlsx is empty. Nothing to send.")
        return False

    today_date = date.today()
    target_row = None
    report_date = today_date

    # Find matching row
    if use_latest:
        target_row = df.iloc[-1]
        try:
            report_date = pd.to_datetime(target_row["Date"]).date()
        except Exception:
            report_date = today_date
        logger.info(f"Using latest recorded row from Excel for date: {report_date}")
    else:
        # Match today's date
        df["_date_parsed"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
        today_matches = df[df["_date_parsed"] == today_date]
        if not today_matches.empty:
            target_row = today_matches.iloc[-1]
            report_date = today_date
            logger.info(f"Found recorded entry for today ({today_date})")
        else:
            # Check if last row was today
            logger.warning(f"No entry found in daily_report.xlsx matching today's date ({today_date}).")
            logger.warning("If you entered your report under a different date, use --latest to force dispatch.")
            return False

    date_str_key = report_date.strftime("%Y-%m-%d")

    # Check for duplicate prevention
    if not force and is_already_sent_today(date_str_key):
        logger.info(f"Report for date {date_str_key} was already sent previously today. Skipping duplicate.")
        logger.info("Use --force if you intentionally wish to resend.")
        return True

    # Check if marked absent ('A') - do not send email
    in_time_check = str(target_row.get("In Time", "")).strip().upper()
    task_check = str(target_row.get("Task📝 ", target_row.get("Task", ""))).strip().upper()
    service_check = str(target_row.get("Service", "")).strip().upper()
    if in_time_check == "A" or (task_check == "A" and service_check == "A"):
        logger.info(f"Entry for {date_str_key} is marked as ABSENT ('A'). Skipping email delivery per configuration.")
        return True

    # Extract row details
    in_time_val = target_row.get("In Time", "09:10 AM")
    out_time_val = target_row.get("Out Time", "06:00 PM")
    in_time_str = format_time_str(in_time_val)
    out_time_str = format_time_str(out_time_val)
    
    task_text = str(target_row.get("Task📝 ", target_row.get("Task", ""))).strip()
    if not task_text or task_text.lower() in ["nan", "none", "a"]:
        task_text = "Completed daily designated tasks and service activities."

    duolingo_val = target_row.get("Dulingo (day,XP,english Score", "")
    if pd.isna(duolingo_val) or not str(duolingo_val).strip() or str(duolingo_val).strip().lower() in ["nan", "none", "a"]:
        duolingo_val = "29 days , 1506XP English Score 60"
    else:
        duolingo_val = str(duolingo_val).strip()

    # Read configuration from .env
    sender_name = os.getenv("SENDER_NAME", "Surya Prakash R")
    sender_email = os.getenv("SENDER_EMAIL", "rsuryaprakashb@gmail.com")
    sender_password = os.getenv("SENDER_PASSWORD", "")
    recipients_raw = os.getenv("RECIPIENTS", "shashikumar.v@supplychainhub.com")
    recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]
    cc_raw = os.getenv("CC_RECIPIENTS", "venkateshwaran.s@supplychainhub.com, Ravikumar.s@supplychainhub.com")
    cc_recipients = [c.strip() for c in cc_raw.split(",") if c.strip()]

    applicant = os.getenv("APPLICANT", "SURYA PRAKASH R")
    reporting_manager = os.getenv("REPORTING_MANAGER", "SHASHIKUMAR V")
    mentor = os.getenv("MENTOR", "N/A")
    module = os.getenv("MODULE", "N/A")
    project_manager = os.getenv("PROJECT_MANAGER", "N/A")
    project_name = os.getenv("PROJECT_NAME", "N/A")

    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", 587))

    if not sender_password:
        logger.error("SENDER_PASSWORD is empty in .env. Cannot send email.")
        return False

    email_subject = f"Daily Report - {report_date.strftime('%d %B %Y')}"
    
    html_content = generate_email_html(
        report_date=report_date,
        applicant=applicant,
        reporting_manager=reporting_manager,
        mentor=mentor,
        module=module,
        project_manager=project_manager,
        project_name=project_name,
        in_time_str=in_time_str,
        out_time_str=out_time_str,
        task_text=task_text,
        duolingo_text=duolingo_val
    )

    plain_content = generate_email_plain(
        report_date=report_date,
        applicant=applicant,
        reporting_manager=reporting_manager,
        mentor=mentor,
        module=module,
        project_manager=project_manager,
        project_name=project_name,
        in_time_str=in_time_str,
        out_time_str=out_time_str,
        task_text=task_text,
        duolingo_text=duolingo_val
    )

    logger.info(f"Dispatching email to: {recipients} (Subject: {email_subject})")
    ok, status_msg = send_report_email(
        subject=email_subject,
        html_content=html_content,
        plain_content=plain_content,
        sender_name=sender_name,
        sender_email=sender_email,
        sender_password=sender_password,
        recipients=recipients,
        cc_recipients=cc_recipients,
        attachment_path=None,
        smtp_server=smtp_server,
        smtp_port=smtp_port
    )

    if ok:
        logger.info("SUCCESS: Email sent successfully!")
        mark_sent_today(date_str_key)
        return True
    else:
        logger.error(f"FAILURE: Email delivery failed: {status_msg}")
        return False

if __name__ == "__main__":
    force_run = "--force" in sys.argv
    use_latest_row = "--latest" in sys.argv
    success = run_scheduled_job(force=force_run, use_latest=use_latest_row)
    sys.exit(0 if success else 1)

