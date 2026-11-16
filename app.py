import os
import re
import html
import smtplib
import ssl
import threading
import time as time_module
from datetime import date, time, datetime, timedelta
import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from dotenv import load_dotenv, set_key

# Import modular components
from email_sender import generate_email_html, generate_email_plain, send_report_email, save_email_to_gmail_drafts
from scheduler_manager import get_task_info, is_task_scheduled, register_task, unregister_task
from send_today_report import mark_sent_today, is_already_sent_today

# Load environment variables
ENV_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
load_dotenv(ENV_FILE, override=True)

def get_config(key: str, default: str = "") -> str:
    """Retrieves configuration from os.environ, .env, or Streamlit Cloud st.secrets."""
    val = os.getenv(key)
    if val is not None and str(val).strip():
        return str(val)
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return default

st.set_page_config(
    page_title="Daily Report Automation",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Universal & Dynamic Theme CSS for BOTH Dark Mode & Light Mode
if "theme_radio" in st.session_state:
    st.session_state["theme_mode"] = "Light" if "Light" in st.session_state["theme_radio"] else "Dark"
elif "theme_mode" not in st.session_state:
    st.session_state["theme_mode"] = get_config("APP_THEME", "Light")

current_theme = st.session_state.get("theme_mode", "Light")

if current_theme == "Light":
    theme_vars = """
    :root {
        --text-color: #0f172a !important;
        --background-color: #f8fafc !important;
        --secondary-background-color: #ffffff !important;
    }
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0 !important;
    }
    [data-testid="stSidebar"] * {
        color: #1e293b !important;
    }
    h1, h2, h3, h4, h5, h6,
    [data-testid="stWidgetLabel"] label, 
    [data-testid="stWidgetLabel"] p, 
    label[data-testid="stWidgetLabel"],
    .stWidgetLabel p,
    .stWidgetLabel {
        color: #0f172a !important;
        font-weight: 600 !important;
    }
    .stMarkdown p, .stMarkdown div, .stMarkdown span {
        color: #1e293b;
    }
    [data-testid="stTextInput"] input, 
    [data-testid="stTextArea"] textarea, 
    [data-testid="stDateInput"] input, 
    [data-testid="stTimeInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
    }
    .stTextInput input:disabled, 
    input:disabled, 
    [data-testid="stTextInput"] input[disabled] {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        opacity: 0.95 !important;
        font-weight: 600 !important;
        border: 1px solid #cbd5e1 !important;
    }
    [data-testid="stCaptionContainer"] p,
    [data-testid="stCaptionContainer"] {
        color: #475569 !important;
    }
    [data-testid="stCaptionContainer"] code,
    .stMarkdown code,
    code {
        color: #15803d !important;
        background-color: #f1f5f9 !important;
        border: 1px solid #cbd5e1 !important;
    }
    [data-testid="stCheckbox"] label, 
    [data-testid="stCheckbox"] label span, 
    [data-testid="stCheckbox"] p,
    [data-testid="stRadio"] label,
    [data-testid="stRadio"] label p,
    [data-testid="stRadio"] div[role="radiogroup"] label {
        color: #0f172a !important;
        font-weight: 600 !important;
    }
    [data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
    }
    [data-testid="stExpander"] details summary span,
    [data-testid="stExpander"] details summary svg {
        color: #0f172a !important;
        fill: #0f172a !important;
    }
    hr, [data-testid="stDivider"] {
        border-color: #e2e8f0 !important;
    }
    [data-testid="stButton"] button {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
    }
    [data-testid="stButton"] button:hover {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        border-color: #94a3b8 !important;
    }
    """
else:
    theme_vars = """
    :root {
        --text-color: #f8fafc !important;
        --background-color: #0e1117 !important;
        --secondary-background-color: #1e293b !important;
    }
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #0e1117 !important;
        color: #f8fafc !important;
    }
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
        background-color: #161b22 !important;
        border-right: 1px solid #30363d !important;
    }
    [data-testid="stSidebar"] * {
        color: #e6edf3 !important;
    }
    h1, h2, h3, h4, h5, h6,
    [data-testid="stWidgetLabel"] label, 
    [data-testid="stWidgetLabel"] p, 
    label[data-testid="stWidgetLabel"],
    .stWidgetLabel p,
    .stWidgetLabel {
        color: #f8fafc !important;
        font-weight: 600 !important;
    }
    .stMarkdown p, .stMarkdown div, .stMarkdown span {
        color: #e2e8f0;
    }
    [data-testid="stTextInput"] input, 
    [data-testid="stTextArea"] textarea, 
    [data-testid="stDateInput"] input, 
    [data-testid="stTimeInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] {
        background-color: #1e293b !important;
        color: #f8fafc !important;
        -webkit-text-fill-color: #f8fafc !important;
        border: 1px solid #334155 !important;
    }
    .stTextInput input:disabled, 
    input:disabled, 
    [data-testid="stTextInput"] input[disabled] {
        background-color: #1e293b !important;
        color: #f8fafc !important;
        -webkit-text-fill-color: #f8fafc !important;
        opacity: 0.95 !important;
        font-weight: 600 !important;
        border: 1px solid #475569 !important;
    }
    [data-testid="stCaptionContainer"] p,
    [data-testid="stCaptionContainer"] {
        color: #94a3b8 !important;
    }
    [data-testid="stCaptionContainer"] code,
    .stMarkdown code,
    code {
        color: #4ade80 !important;
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
    }
    [data-testid="stCheckbox"] label, 
    [data-testid="stCheckbox"] label span, 
    [data-testid="stCheckbox"] p,
    [data-testid="stRadio"] label,
    [data-testid="stRadio"] label p,
    [data-testid="stRadio"] div[role="radiogroup"] label {
        color: #f8fafc !important;
        font-weight: 600 !important;
    }
    [data-testid="stExpander"] {
        background-color: #161b22 !important;
        border: 1px solid #30363d !important;
    }
    [data-testid="stExpander"] details summary span,
    [data-testid="stExpander"] details summary svg {
        color: #f8fafc !important;
        fill: #f8fafc !important;
    }
    hr, [data-testid="stDivider"] {
        border-color: #30363d !important;
    }
    [data-testid="stButton"] button {
        background-color: #1e293b !important;
        color: #f8fafc !important;
        border: 1px solid #334155 !important;
    }
    [data-testid="stButton"] button:hover {
        background-color: #334155 !important;
        color: #ffffff !important;
        border-color: #475569 !important;
    }
    """

st.markdown(f"<style>{theme_vars}</style>", unsafe_allow_html=True)

UNIVERSAL_RESPONSIVE_CSS = """
<style>
/* Common Universal Styles */
[data-testid="stDeployButton"], .stDeployButton {
    display: none !important;
}

[data-testid="stBaseButton-primary"], button[kind="primary"] {
    background-color: #ff4b4b !important;
    color: #ffffff !important;
    border: none !important;
    font-weight: 700 !important;
}
[data-testid="stBaseButton-primary"]:hover, button[kind="primary"]:hover {
    background-color: #e03e3e !important;
    color: #ffffff !important;
}

/* Live Email Preview Card (Always pure white document with crisp black table) */
.email-preview-card,
div[data-testid="stExpander"] div.stHtml,
div[data-testid="stExpander"] div[data-testid="stHtml"] {
    background-color: #ffffff !important;
    color: #000000 !important;
    border-radius: 8px !important;
    padding: 24px 28px;
    border: 1px solid #cbd5e1 !important;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.12) !important;
    width: 100% !important;
    box-sizing: border-box !important;
}

.email-table-scroll-wrapper {
    width: 100% !important;
    overflow-x: auto !important;
    -webkit-overflow-scrolling: touch !important;
}

.email-preview-card p,
.email-preview-card div,
.email-preview-card span,
div[data-testid="stExpander"] div.stHtml p,
div[data-testid="stExpander"] div[data-testid="stHtml"] p {
    color: #000000 !important;
    font-family: Arial, Helvetica, sans-serif !important;
}

.email-preview-card table,
div[data-testid="stExpander"] div.stHtml table,
div[data-testid="stExpander"] div[data-testid="stHtml"] table {
    background-color: #ffffff !important;
    color: #000000 !important;
    border: 1px solid #000000 !important;
    border-collapse: collapse !important;
    width: 100% !important;
    max-width: 680px !important;
}

.email-preview-card td,
div[data-testid="stExpander"] div.stHtml td,
div[data-testid="stExpander"] div[data-testid="stHtml"] td {
    color: #000000 !important;
    border: 1px solid #000000 !important;
    background-color: #ffffff !important;
    padding: 10px 14px !important;
}

.email-preview-card td *,
div[data-testid="stExpander"] div.stHtml td * {
    color: #000000 !important;
}

.email-preview-card .rss-wine,
div[data-testid="stExpander"] div.stHtml .rss-wine,
.email-preview-card p.rss-wine {
    color: #741b47 !important;
    font-weight: bold !important;
    font-family: Georgia, 'Times New Roman', serif !important;
}

.email-preview-card .rss-indigo,
div[data-testid="stExpander"] div.stHtml .rss-indigo,
.email-preview-card p.rss-indigo {
    color: #351c75 !important;
    font-weight: bold !important;
    font-family: Georgia, 'Times New Roman', serif !important;
}

/* ========================================================
   📱 MOBILE RESPONSIVE OPTIMIZATIONS (< 768px and < 640px)
   ======================================================== */
@media (max-width: 768px) {
    /* 1. Eliminate heavy desktop padding on phones & tablets */
    .block-container,
    [data-testid="stMainBlockContainer"],
    [data-testid="stAppViewBlockContainer"] {
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        max-width: 100% !important;
    }

    /* 2. Responsive Headings */
    h1 {
        font-size: 1.6rem !important;
        line-height: 1.25 !important;
    }
    h2 { font-size: 1.35rem !important; }
    h3 { font-size: 1.2rem !important; }
    h4 { font-size: 1.05rem !important; }

    /* 3. Mobile touch-friendly inputs (16px prevents iOS Safari auto-zoom) */
    [data-testid="stTextInput"] input, 
    [data-testid="stTextArea"] textarea, 
    [data-testid="stDateInput"] input, 
    [data-testid="stTimeInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] {
        font-size: 16px !important;
        min-height: 44px !important;
    }

    /* 4. Touch-friendly Buttons (min 46px height) */
    [data-testid="stButton"] button,
    [data-testid="stDownloadButton"] button,
    [data-testid="stBaseButton-primary"] {
        min-height: 46px !important;
        font-size: 15px !important;
        width: 100% !important;
    }

    /* 5. Mobile Sidebar adjustments */
    [data-testid="stSidebar"] {
        max-width: 88vw !important;
    }

    /* 6. Email Preview Card: Responsive padding & font */
    .email-preview-card {
        padding: 16px 14px !important;
        border-radius: 6px !important;
    }
    .email-preview-card table {
        font-size: 12.5px !important;
        min-width: 290px !important;
    }
    .email-preview-card td {
        padding: 7px 8px !important;
    }
    .email-preview-card .rss-indent {
        padding-left: 18px !important;
    }

    /* 7. Responsive Tables & DataFrames */
    [data-testid="stDataFrame"], 
    [data-testid="stTable"] {
        width: 100% !important;
        overflow-x: auto !important;
        -webkit-overflow-scrolling: touch !important;
    }
}

@media (max-width: 640px) {
    /* 8. Stack form columns on narrow phones for 100% full-width touch area */
    [data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
        gap: 10px !important;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        flex: 1 1 100% !important;
        min-width: 100% !important;
        width: 100% !important;
    }

    /* 9. Hide desktop alignment spacer so checkbox is not pushed down */
    .mobile-hide-spacer {
        display: none !important;
    }

    /* 10. Radio theme switcher full-width on mobile */
    [data-testid="stRadio"] div[role="radiogroup"] {
        justify-content: flex-start !important;
    }
}
</style>
"""
st.markdown(UNIVERSAL_RESPONSIVE_CSS, unsafe_allow_html=True)

def get_excel_path() -> str:
    base = os.path.dirname(os.path.abspath(__file__))
    for fname in ["Daily_report.xlsx", "daily_report.xlsx"]:
        candidate = os.path.join(base, fname)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(base, "daily_report.xlsx")

EXCEL_FILE = get_excel_path()
HEADERS = ["Date", "Days", "In Time", "Out Time", "Service", "Task📝 ", "Dulingo (day,XP,english Score"]
DUOLINGO_USERNAME = "rssuryakkl"

def update_env_variable(key: str, value: str):
    """Safely updates a key-value pair in .env file and reloads environment."""
    set_key(ENV_FILE, key, value)
    load_dotenv(ENV_FILE, override=True)

def parse_time_str(time_str: str, default_time: time = time(19, 0)) -> time:
    """Parses a time string (e.g. '19:00' or '07:00 PM') into a datetime.time object."""
    if not time_str or not time_str.strip():
        return default_time
    for fmt in ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"]:
        try:
            return datetime.strptime(time_str.strip(), fmt).time()
        except ValueError:
            pass
    return default_time

def fetch_duolingo_live(username=DUOLINGO_USERNAME):
    """Fetches real-time streak, total XP, and profile name from Duolingo API."""
    url = f"https://www.duolingo.com/2017-06-30/users?username={username}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        r = requests.get(url, headers=headers, timeout=5)
        if r.status_code == 200:
            data = r.json()
            users = data.get("users", [])
            if users:
                user = users[0]
                streak = user.get("streak", 0)
                total_xp = user.get("totalXp", 0)
                name = user.get("name") or username
                return streak, total_xp, name, True
    except Exception:
        pass
    return None, None, username, False

def get_latest_duolingo_history():
    """Reads the last recorded Duolingo entry from daily_report.xlsx to populate defaults."""
    default_days = 29
    default_xp = 1506
    default_score = 60
    has_history = False

    if not os.path.exists(EXCEL_FILE):
        return default_days, default_xp, default_score, False

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        sheet = wb.active
        for r in range(sheet.max_row, 1, -1):
            val = sheet.cell(r, 7).value
            if val and str(val).strip() and str(val).strip().lower() not in ["none", "nil", "a"]:
                val_str = str(val)
                days_match = re.search(r'(\d+)\s*(?:days?|day)', val_str, re.IGNORECASE)
                xp_match = re.search(r'(\d+)\s*xp', val_str, re.IGNORECASE)
                score_match = re.search(r'(?:score\s*|english\s*score\s*)(\d+)', val_str, re.IGNORECASE)

                if days_match:
                    default_days = int(days_match.group(1))
                if xp_match:
                    default_xp = int(xp_match.group(1))
                if score_match:
                    default_score = int(score_match.group(1))
                has_history = True
                break
    except Exception:
        pass

    return default_days, default_xp, default_score, has_history

def save_report_to_excel(data: dict):
    """Appends report entry to the Excel file with the exact matching headers and dd-mm-yyyy date format."""
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            sheet = wb.active
        except Exception:
            wb = openpyxl.Workbook()
            sheet = wb.active
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active

    first_row = [str(cell.value).strip() if cell.value else "" for cell in sheet[1]] if sheet.max_row >= 1 else []
    has_headers = any("Date" in str(v) for v in first_row) and any("Days" in str(v) for v in first_row)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    if not has_headers:
        header_fill = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="000000")
        sheet.insert_rows(1)
        for col_idx, h in enumerate(HEADERS, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        sheet.row_dimensions[1].height = 28

    # Clean up any completely empty rows so no gaps or spaces ever exist
    empty_rows = [
        r for r in range(2, sheet.max_row + 1)
        if all(sheet.cell(r, c).value is None or str(sheet.cell(r, c).value).strip() == "" for c in range(1, 8))
    ]
    for r in sorted(empty_rows, reverse=True):
        sheet.delete_rows(r, 1)

    target_date = data["date_obj"]
    date_val = datetime.combine(target_date, datetime.min.time())

    row_values = [
        date_val,
        data["days"],
        data["in_time"],
        data["out_time"],
        data["service"],
        data["task"],
        data["duolingo"]
    ]

    # Check if entry for this date already exists in the sheet to update it
    target_row_idx = None
    for r in range(2, sheet.max_row + 1):
        c_val = sheet.cell(r, 1).value
        if c_val is not None:
            c_date = None
            if isinstance(c_val, datetime):
                c_date = c_val.date()
            elif isinstance(c_val, date):
                c_date = c_val
            elif isinstance(c_val, str):
                for fmt in ["%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y"]:
                    try:
                        c_date = datetime.strptime(c_val.strip(), fmt).date()
                        break
                    except ValueError:
                        pass
            if c_date == target_date:
                target_row_idx = r
                break

    if target_row_idx is not None:
        target_row = target_row_idx
        for col_idx, val in enumerate(row_values, start=1):
            sheet.cell(target_row, col_idx, value=val)
    else:
        sheet.append(row_values)
        target_row = sheet.max_row

    sheet.row_dimensions[target_row].height = 24

    data_font = Font(name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for col_idx in range(1, len(row_values) + 1):
        cell = sheet.cell(target_row, col_idx)
        cell.font = data_font
        cell.border = thin_border

        # Task (col 6) and Duolingo (col 7) aligned Left; Date, Day, Times, Service centered
        if col_idx in [6, 7]:
            cell.alignment = left_align
        else:
            cell.alignment = center_align

        if col_idx == 1:
            cell.number_format = 'dd-mm-yyyy'

    for col in sheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or "")
            lines = val.split("\n")
            longest_line = max((len(l) for l in lines), default=0)
            if longest_line > max_len:
                max_len = longest_line
        sheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(EXCEL_FILE)

def background_timer_send(delay_seconds: float, email_kwargs: dict, report_date_key: str):
    """Waits delay_seconds and sends the email in a background daemon thread."""
    def worker():
        if delay_seconds > 0:
            time_module.sleep(delay_seconds)
        if is_already_sent_today(report_date_key):
            return
        ok, msg = send_report_email(**email_kwargs)
        if ok:
            mark_sent_today(report_date_key)
    
    t = threading.Thread(target=worker, daemon=True)
    t.start()


# ==========================================
# --- SIDEBAR: Email & Scheduler Settings ---
# ==========================================
with st.sidebar:
    st.header("⚙️ Email & Mentor Setup")
    st.caption("🌓 **Theme:** Toggle between **Light**, **Dark**, or **System** anytime in the top-right **⋮** menu.")
    
    cfg_sender_name = st.text_input(
        "Sender Display Name",
        value=get_config("SENDER_NAME", "Surya Prakash R")
    )
    
    cfg_sender_email = st.text_input(
        "Sender Gmail",
        value=get_config("SENDER_EMAIL", "rsuryaprakashb@gmail.com")
    )
    
    cfg_sender_password = st.text_input(
        "Gmail App Password",
        value=get_config("SENDER_PASSWORD", ""),
        type="password",
        help="16-character Google App Password (spaces will be automatically handled)."
    )
    
    cfg_recipients = st.text_area(
        "Mentor / Recipients (To)",
        value=get_config("RECIPIENTS", "shashikumar.v@supplychainhub.com"),
        height=70,
        help="Enter mentor and manager email addresses separated by commas."
    )
    
    cfg_cc = st.text_input(
        "CC (Optional)",
        value=get_config("CC_RECIPIENTS", "venkateshwaran.s@supplychainhub.com, Ravikumar.s@supplychainhub.com")
    )
    
    cfg_smtp_server = get_config("SMTP_SERVER", "smtp.gmail.com")
    cfg_smtp_port = int(get_config("SMTP_PORT", "587"))
    
    if st.button("💾 Save Settings to .env", use_container_width=True):
        update_env_variable("SENDER_NAME", cfg_sender_name.strip())
        update_env_variable("SENDER_EMAIL", cfg_sender_email.strip())
        update_env_variable("SENDER_PASSWORD", cfg_sender_password.strip())
        update_env_variable("RECIPIENTS", cfg_recipients.strip())
        update_env_variable("CC_RECIPIENTS", cfg_cc.strip())
        st.success("✅ Credentials saved to `.env`!")

    st.divider()

    # --- Windows Task Scheduler Status & Custom Time ---
    st.markdown("### ⏰ Daily Windows Task Scheduler")
    st.caption("Runs silently on Windows even when your browser or PC terminal is closed.")
    
    env_sch_time_str = get_config("SCHEDULED_TIME", "19:00")
    env_sch_time_obj = parse_time_str(env_sch_time_str, time(19, 0))
    
    sidebar_sch_time = st.time_input(
        "Customized Daily Windows Schedule Time:",
        value=env_sch_time_obj,
        step=60,
        help="Choose the daily time when Windows Task Scheduler will run send_today_report.py."
    )
    time_24hr_str = sidebar_sch_time.strftime("%H:%M")

    task_is_active, task_info_str = get_task_info()
    if task_is_active:
        st.success(f"🟢 **Active:** Daily at **{sidebar_sch_time.strftime('%I:%M %p')}**", icon="⏰")
        if task_info_str:
            st.caption(f"Next trigger: `{task_info_str}`")
        
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🔄 Update Time", use_container_width=True):
                update_env_variable("SCHEDULED_TIME", time_24hr_str)
                ok, msg = register_task(time_24hr_str)
                if ok:
                    st.success(f"Updated schedule to {sidebar_sch_time.strftime('%I:%M %p')}!")
                    st.rerun()
                else:
                    st.error(msg)
        with c_btn2:
            if st.button("❌ Disable Task", use_container_width=True):
                ok, msg = unregister_task()
                if ok:
                    st.warning("Disabled scheduled task.")
                    st.rerun()
                else:
                    st.error(msg)
    else:
        st.info("⚪ **Not Registered** in Windows Scheduler")
        if st.button(f"Enable Task at {sidebar_sch_time.strftime('%I:%M %p')}", use_container_width=True):
            update_env_variable("SCHEDULED_TIME", time_24hr_str)
            ok, msg = register_task(time_24hr_str)
            if ok:
                st.success(f"Registered Windows Task at {sidebar_sch_time.strftime('%I:%M %p')}!")
                st.rerun()
            else:
                st.error(msg)

    st.markdown("---")
    if st.button("🧪 Send Test Email", use_container_width=True):
        if not cfg_sender_password.strip():
            st.error("Please enter your Gmail App Password first.")
        else:
            with st.spinner("Sending test email..."):
                t_subj = "Test Email - Daily Report Automation"
                t_html = "<p>Hello! This is a test email verifying that your Daily Report Automation email setup works properly.</p>"
                t_plain = "Hello! This is a test email verifying that your Daily Report Automation email setup works properly."
                t_recips = [r.strip() for r in cfg_recipients.split(",") if r.strip()]
                t_cc = [c.strip() for c in cfg_cc.split(",") if c.strip()]
                ok, msg = send_report_email(
                    subject=t_subj,
                    html_content=t_html,
                    plain_content=t_plain,
                    sender_name=cfg_sender_name,
                    sender_email=cfg_sender_email,
                    sender_password=cfg_sender_password,
                    recipients=t_recips,
                    cc_recipients=t_cc,
                    attachment_path=None,
                    smtp_server=cfg_smtp_server,
                    smtp_port=cfg_smtp_port
                )
                if ok:
                    st.success("✅ Test email sent successfully!")
                else:
                    st.error(f"❌ Failed to send: {msg}")


# ==========================================
# --- MAIN APPLICATION UI ---
# ==========================================
header_col1, header_col2 = st.columns([3.2, 1.4])
with header_col1:
    st.title("📝 Daily Report Automation")
    st.markdown("Record your daily work report into `daily_report.xlsx` and schedule automatic email delivery to your mentor.")
with header_col2:
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    theme_choice = st.radio(
        "Theme",
        options=["☀️ Light", "🌙 Dark"],
        index=0 if current_theme == "Light" else 1,
        horizontal=True,
        label_visibility="collapsed",
        key="theme_radio"
    )
    chosen_theme = "Light" if "Light" in theme_choice else "Dark"
    if chosen_theme != st.session_state.get("theme_mode"):
        st.session_state["theme_mode"] = chosen_theme
        update_env_variable("APP_THEME", chosen_theme)
        st.rerun()

col1, col2, col3 = st.columns([2, 1, 1.4])
with col1:
    report_date = st.date_input("Report Date", date.today(), format="DD-MM-YYYY")
with col2:
    day_name = report_date.strftime("%A")
    st.text_input("Day", value=day_name, disabled=True)
with col3:
    st.markdown("<div class='mobile-hide-spacer' style='height: 28px;'></div>", unsafe_allow_html=True)
    is_absent = st.checkbox("🚫 Mark as ABSENT", value=False, help="Check this if you were absent/on leave. Only Date, Day, and Duolingo will be saved to Excel. No email will be sent.")

if not is_absent:
    col_in, col_out = st.columns(2)
    with col_in:
        in_time = st.time_input("In Time (AM)", time(9, 20))
        in_time_str = in_time.strftime("%I:%M %p")
        st.caption(f"Recorded in Excel as: **`{in_time_str}`**")

    with col_out:
        out_time = st.time_input("Out Time (PM)", time(18, 0))
        out_time_str = out_time.strftime("%I:%M %p")
        st.caption(f"Recorded in Excel as: **`{out_time_str}`**")

    service = st.text_input("Service", value="Nil", placeholder="e.g. Wood Arrangement, Onion Chops, Sukku Tea, Nil...")
    tasks = st.text_area("Task 📝 *", placeholder="Enter tasks completed, topics learned, sessions...", height=120)

# --- Duolingo Tracker Section ---
st.markdown("#### 🦉 Duolingo Progress")
last_days, last_xp, last_score, has_history = get_latest_duolingo_history()

if "duo_synced" not in st.session_state:
    live_streak, live_xp, live_name, success = fetch_duolingo_live(DUOLINGO_USERNAME)
    if success:
        st.session_state["duo_days"] = live_streak
        st.session_state["duo_xp"] = live_xp
        st.session_state["duo_name"] = live_name
        st.session_state["duo_synced"] = True
    else:
        st.session_state["duo_days"] = last_days + 1 if has_history else 29
        st.session_state["duo_xp"] = last_xp if has_history else 1506
        st.session_state["duo_name"] = DUOLINGO_USERNAME
        st.session_state["duo_synced"] = False

if "duo_score" not in st.session_state:
    st.session_state["duo_score"] = last_score if has_history else 60

d_col_info, d_col_btn = st.columns([3, 1])
with d_col_info:
    if st.session_state.get("duo_synced"):
        st.success(f"🟢 Synced with **@{DUOLINGO_USERNAME}** ({st.session_state.get('duo_name')})", icon="🦉")
    elif has_history:
        st.caption(f"📜 Last recorded in Excel: {last_days} days | {last_xp} XP | English Score: {last_score}")
    else:
        st.caption(f"Duolingo user: @{DUOLINGO_USERNAME}")

with d_col_btn:
    if st.button("🔄 Sync Live", use_container_width=True, help="Fetch latest live streak and XP from Duolingo"):
        with st.spinner("Fetching from Duolingo..."):
            live_streak, live_xp, live_name, success = fetch_duolingo_live(DUOLINGO_USERNAME)
            if success:
                st.session_state["duo_days"] = live_streak
                st.session_state["duo_xp"] = live_xp
                st.session_state["duo_name"] = live_name
                st.session_state["duo_synced"] = True
                st.rerun()
            else:
                st.warning("Could not reach Duolingo API.")

duo_col1, duo_col2, duo_col3 = st.columns(3)
with duo_col1:
    duo_days = st.number_input(
        "Days Streak (Daily)",
        min_value=1,
        value=int(st.session_state.get("duo_days", 29)),
        step=1,
        help="Duolingo streak"
    )
with duo_col2:
    duo_xp = st.number_input(
        "XP (Daily)",
        min_value=0,
        value=int(st.session_state.get("duo_xp", 1506)),
        step=10,
        help="Total Duolingo XP"
    )
with duo_col3:
    duo_score = st.number_input(
        "English Score (Rare to change)",
        min_value=0,
        max_value=160,
        value=int(st.session_state.get("duo_score", 60)),
        step=1,
        help="English score"
    )

duolingo_final = f"{duo_days} days , {duo_xp}XP English Score {duo_score}"
st.caption(f"Will be recorded in Excel as: **`{duolingo_final}`**")

if is_absent:
    # Absent Action Button (Saves to Excel only, NO email)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 Save Absent Entry to Excel", type="primary", use_container_width=True):
        entry = {
            "date_obj": report_date,
            "days": day_name,
            "in_time": "A",
            "out_time": "A",
            "service": "A",
            "task": "A",
            "duolingo": duolingo_final.strip()
        }
        excel_saved = False
        try:
            save_report_to_excel(entry)
            excel_saved = True
            st.success(f"✅ Marked **ABSENT ('A')** for **{report_date.strftime('%d-%m-%Y')} ({day_name})** and saved into `daily_report.xlsx`!")
            st.caption("ℹ️ No email or draft was created.")
            st.balloons()
        except PermissionError:
            st.error("⚠️ **Cannot save to Excel:** `daily_report.xlsx` is currently open in Microsoft Excel. Please **close the file in Excel** and try again.")
        except Exception as e:
            st.error(f"❌ An error occurred while saving to Excel: {e}")
else:
    # --- Email Table Metadata Details ---
    with st.expander("📋 Email Table Metadata (Applicant, Manager, Mentor...)", expanded=False):
        st.caption("Customize the fields shown in the email table. Defaults match your official format.")
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            meta_applicant = st.text_input("Applicant", value=get_config("APPLICANT", "SURYA PRAKASH R"))
            meta_rep_manager = st.text_input("Reporting Manager", value=get_config("REPORTING_MANAGER", "SHASHIKUMAR V"))
            meta_mentor = st.text_input("Mentor", value=get_config("MENTOR", "N/A"))
        with m_col2:
            meta_module = st.text_input("Module", value=get_config("MODULE", "N/A"))
            meta_proj_manager = st.text_input("Project Manager", value=get_config("PROJECT_MANAGER", "N/A"))
            meta_proj_name = st.text_input("Project Name", value=get_config("PROJECT_NAME", "N/A"))

    # Generate formatted email content
    email_subject = f"Daily Report - {report_date.strftime('%d %B %Y')}"
    email_html = generate_email_html(
        report_date=report_date,
        applicant=meta_applicant,
        reporting_manager=meta_rep_manager,
        mentor=meta_mentor,
        module=meta_module,
        project_manager=meta_proj_manager,
        project_name=meta_proj_name,
        in_time_str=in_time_str,
        out_time_str=out_time_str,
        task_text=tasks if tasks.strip() else "(Your tasks will appear here)",
        duolingo_text=duolingo_final
    )
    email_plain = generate_email_plain(
        report_date=report_date,
        applicant=meta_applicant,
        reporting_manager=meta_rep_manager,
        mentor=meta_mentor,
        module=meta_module,
        project_manager=meta_proj_manager,
        project_name=meta_proj_name,
        in_time_str=in_time_str,
        out_time_str=out_time_str,
        task_text=tasks if tasks.strip() else "(Your tasks will appear here)",
        duolingo_text=duolingo_final
    )

    # Action & Dispatch Options
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 🚀 Email Action & Dispatch")

    send_now_override = st.checkbox("⚡ Send immediately right now (skip draft)", value=False)

    with st.expander("👁️ Live Email Preview (Strict Format & Alignment)", expanded=False):
        st.markdown(f"**Subject:** `{email_subject}`")
        st.markdown(f"**From:** `{cfg_sender_name} <{cfg_sender_email}>`")
        st.markdown(f"**To:** `{cfg_recipients}`")
        if cfg_cc.strip():
            st.markdown(f"**Cc:** `{cfg_cc}`")
        st.markdown("---")
        preview_card = f"""
        <div class="email-preview-card" style="background-color: #ffffff !important; color: #000000 !important; border-radius: 8px; border: 1px solid #cbd5e1; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15); margin-top: 10px;">
            <div class="email-table-scroll-wrapper">
                {email_html}
            </div>
        </div>
        """
        st.html(preview_card)

    st.markdown("<br>", unsafe_allow_html=True)

    # Main Action Button
    btn_label = "⚡ Generate & Send Immediately Now" if send_now_override else "📥 Generate Report & Save to Gmail Draft"
    if st.button(btn_label, type="primary", use_container_width=True):
        if not tasks.strip():
            st.error("⚠️ Please enter your tasks.")
        else:
            entry = {
                "date_obj": report_date,
                "days": day_name,
                "in_time": in_time.strftime("%I:%M %p"),
                "out_time": out_time.strftime("%I:%M %p"),
                "service": service.strip() or "Nil",
                "task": tasks.strip(),
                "duolingo": duolingo_final.strip()
            }
            excel_saved = False
            try:
                save_report_to_excel(entry)
                excel_saved = True
                st.success(f"✅ Report for **{report_date.strftime('%d-%m-%Y')} ({day_name})** saved into `daily_report.xlsx`!")
            except PermissionError:
                st.error("⚠️ **Cannot save to Excel:** `daily_report.xlsx` is currently open in Microsoft Excel. Please **close the file in Excel** and try again.")
            except Exception as e:
                st.error(f"❌ An error occurred while saving to Excel: {e}")

            if excel_saved:
                recips_list = [r.strip() for r in cfg_recipients.split(",") if r.strip()]
                cc_list = [c.strip() for c in cfg_cc.split(",") if c.strip()]
                date_key = report_date.strftime("%Y-%m-%d")

                if not cfg_sender_password.strip():
                    st.warning("⚠️ App password missing. Please enter your Gmail App Password in the sidebar settings so emails can be sent or saved to draft.")
                elif send_now_override:
                    email_kwargs = {
                        "subject": email_subject,
                        "html_content": email_html,
                        "plain_content": email_plain,
                        "sender_name": cfg_sender_name,
                        "sender_email": cfg_sender_email,
                        "sender_password": cfg_sender_password,
                        "recipients": recips_list,
                        "cc_recipients": cc_list,
                        "attachment_path": None,
                        "smtp_server": cfg_smtp_server,
                        "smtp_port": cfg_smtp_port
                    }
                    with st.spinner("📧 Dispatching report email to mentor..."):
                        ok, send_status = send_report_email(**email_kwargs)
                        if ok:
                            mark_sent_today(date_key)
                            st.success(f"📧 **Email successfully sent to mentor:** `{', '.join(recips_list)}`!")
                            st.balloons()
                        else:
                            st.error(f"❌ **Email delivery failed:** {send_status}")
                else:
                    # Save directly into Gmail Drafts!
                    with st.spinner("📥 Saving formatted report to your Gmail Drafts..."):
                        ok, draft_status = save_email_to_gmail_drafts(
                            subject=email_subject,
                            html_content=email_html,
                            plain_content=email_plain,
                            sender_name=cfg_sender_name,
                            sender_email=cfg_sender_email,
                            sender_password=cfg_sender_password,
                            recipients=recips_list,
                            cc_recipients=cc_list
                        )
                    if ok:
                        st.success("📬 **Email successfully saved to your Gmail Drafts!**")
                        st.balloons()
                    else:
                        st.error(f"❌ **Could not save to Gmail Drafts:** {draft_status}")

# --- Preview Table ---
st.divider()
st.subheader("📋 Registered Reports History")

if os.path.exists(EXCEL_FILE):
    try:
        wb_check = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
        active_sheet_name = wb_check.active.title
        wb_check.close()
        df = pd.read_excel(EXCEL_FILE, sheet_name=active_sheet_name)
        if not df.empty:
            if "Date" in df.columns:
                formatted_dates = pd.to_datetime(df["Date"], errors='coerce').dt.strftime("%d-%m-%Y")
                df["Date"] = formatted_dates.fillna(df["Date"].astype(str))
            
            for time_col in ["In Time", "Out Time"]:
                if time_col in df.columns:
                    df[time_col] = df[time_col].apply(lambda x: x.strftime("%I:%M %p") if hasattr(x, "strftime") else str(x))

            df_display = df.fillna("-").astype(str)
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            with open(EXCEL_FILE, "rb") as f:
                st.download_button(
                    label="📥 Download daily_report.xlsx",
                    data=f.read(),
                    file_name="daily_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("No reports logged yet. Submit the form above to add an entry!")
    except PermissionError:
        st.info("ℹ️ Excel file is open in another program. Close it to view latest saved preview.")
    except Exception as e:
        st.info(f"Could not load preview: {e}")
