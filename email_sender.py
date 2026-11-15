import os
import re
import html
import smtplib
import imaplib
import time
import ssl
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate, make_msgid

def generate_email_html(
    report_date: date,
    applicant: str,
    reporting_manager: str,
    mentor: str,
    module: str,
    project_manager: str,
    project_name: str,
    in_time_str: str,
    out_time_str: str,
    task_text: str,
    duolingo_text: str
) -> str:
    """Generates the exact HTML email matching the required format and strict left-alignment."""
    long_date = report_date.strftime("%d %B %Y")
    delivery_date = report_date.strftime("%d-%b-%Y").upper()
    hours_str = f"{in_time_str} - {out_time_str}"
    
    escaped_task = html.escape(task_text.strip()).replace('\n', '<br>')
    escaped_duo = html.escape(duolingo_text.strip())

    html_code = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body {{
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14px;
      color: #000000;
      line-height: 1.5;
      background-color: #ffffff;
      margin: 0;
      padding: 10px;
    }}
    p {{
      margin: 0 0 16px 0;
      font-size: 14px;
      color: #000000;
    }}
    table.report-table {{
      border-collapse: collapse;
      width: 100%;
      max-width: 680px;
      font-family: Arial, Helvetica, sans-serif;
      font-size: 14px;
      color: #000000;
      border: 1px solid #000000;
      margin-top: 15px;
      margin-bottom: 25px;
    }}
    table.report-table td {{
      border: 1px solid #000000;
      padding: 10px 14px;
      text-align: left;
      vertical-align: top;
    }}
    .label-col {{
      width: 34%;
      font-weight: bold;
    }}
    .value-col {{
      width: 66%;
    }}
  </style>
</head>
<body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #000000; line-height: 1.5; margin: 0; padding: 10px; background-color: #ffffff;">
  <p style="margin: 0 0 16px 0; font-size: 14px; color: #000000;">Dear Sir,</p>
  <p style="margin: 0 0 16px 0; font-size: 14px; color: #000000;">Please find attached my daily work report for {long_date}, including the tasks completed, services provided, and learning activities for the day.</p>
  
  <table style="border-collapse: collapse; width: 100%; max-width: 680px; font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #000000; border: 1px solid #000000; margin-top: 15px; margin-bottom: 25px;">
    <tbody>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; width: 34%; text-align: left; vertical-align: top;">Applicant</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top; width: 66%;">{html.escape(applicant)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Delivery Date</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{delivery_date}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Reporting Manager</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{html.escape(reporting_manager)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Mentor</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{html.escape(mentor)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Module</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{html.escape(module)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Project Manager</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{html.escape(project_manager)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Project Name</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{html.escape(project_name)}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Number of Hours</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top;">{hours_str}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 10px 14px; font-weight: bold; text-align: left; vertical-align: top;">Description of Work</td>
        <td style="border: 1px solid #000000; padding: 10px 14px; text-align: left; vertical-align: top; line-height: 1.6;">
          {escaped_task}<br><br>
          <b>Additional Skills Focused:</b><br>
          &nbsp;&nbsp;<b>Duolingo : {escaped_duo}</b>
        </td>
      </tr>
    </tbody>
  </table>

  <div style="margin-top: 30px; font-family: Georgia, 'Times New Roman', serif; font-size: 13.5px; line-height: 1.5;">
    <p class="rss-wine" style="margin: 0 0 6px 0; color: #741b47; font-weight: bold; font-family: Georgia, 'Times New Roman', serif;">With Regards,</p>
    <div style="padding-left: 65px; font-family: Georgia, 'Times New Roman', serif;">
      <p class="rss-wine" style="margin: 0 0 3px 0; color: #741b47; font-weight: bold;">Surya Prakash R</p>
      <p class="rss-indigo" style="margin: 0 0 3px 0; color: #351c75; font-weight: bold;">B.Tech &ndash; Artificial Intelligence and Machine Learning,</p>
      <p class="rss-indigo" style="margin: 0 0 3px 0; color: #351c75; font-weight: bold;">Vadalur, Cuddalore District</p>
      <p class="rss-indigo" style="margin: 0 0 3px 0; color: #351c75; font-weight: bold;">Tamil Nadu, India (607303).</p>
      <p class="rss-wine" style="margin: 0; color: #741b47; font-weight: bold;">Contact: +91 97151 23929</p>
    </div>
  </div>
</body>
</html>"""
    return html_code

def generate_email_plain(
    report_date: date,
    applicant: str,
    reporting_manager: str,
    mentor: str,
    module: str,
    project_manager: str,
    project_name: str,
    in_time_str: str,
    out_time_str: str,
    task_text: str,
    duolingo_text: str
) -> str:
    """Generates plain text alternative for email clients that do not render HTML."""
    long_date = report_date.strftime("%d %B %Y")
    delivery_date = report_date.strftime("%d-%b-%Y").upper()
    hours_str = f"{in_time_str} - {out_time_str}"
    
    return f"""Dear Sir,

Please find attached my daily work report for {long_date}, including the tasks completed, services provided, and learning activities for the day.

Applicant: {applicant}
Delivery Date: {delivery_date}
Reporting Manager: {reporting_manager}
Mentor: {mentor}
Module: {module}
Project Manager: {project_manager}
Project Name: {project_name}
Number of Hours: {hours_str}

Description of Work:
{task_text}

Additional Skills Focused:
  Duolingo : {duolingo_text}

With Regards,
        Surya Prakash R
        B.Tech - Artificial Intelligence and Machine Learning,
        Vadalur, Cuddalore District
        Tamil Nadu, India (607303).
        Contact: +91 97151 23929
"""

def send_report_email(
    subject: str,
    html_content: str,
    plain_content: str,
    sender_name: str,
    sender_email: str,
    sender_password: str,
    recipients: list,
    cc_recipients: list = None,
    attachment_path: str = None,
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587
) -> tuple[bool, str]:
    """Sends the formatted report email via SMTP."""
    if not sender_email or not sender_email.strip():
        return False, "Sender email is empty. Please enter your email."
    if not sender_password or not sender_password.strip():
        return False, "Gmail App Password is required. Please set it in the sidebar or .env file."
    if not recipients:
        return False, "No recipient email addresses specified."

    clean_recipients = [r.strip() for r in recipients if r and r.strip()]
    clean_cc = [c.strip() for c in (cc_recipients or []) if c and c.strip()]
    all_targets = clean_recipients + clean_cc

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["To"] = ", ".join(clean_recipients)
    if clean_cc:
        msg["Cc"] = ", ".join(clean_cc)

    body_part = MIMEMultipart("alternative")
    body_part.attach(MIMEText(plain_content, "plain", "utf-8"))
    body_part.attach(MIMEText(html_content, "html", "utf-8"))
    msg.attach(body_part)

    if attachment_path and os.path.exists(attachment_path):
        try:
            with open(attachment_path, "rb") as f:
                excel_part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                excel_part.set_payload(f.read())
            encoders.encode_base64(excel_part)
            filename = os.path.basename(attachment_path)
            excel_part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
            msg.attach(excel_part)
        except Exception as e:
            return False, f"Could not read attachment '{attachment_path}': {e}"

    clean_password = sender_password.strip().replace(" ", "")

    try:
        if int(smtp_port) == 465:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_server, int(smtp_port), context=context, timeout=20) as server:
                server.login(sender_email.strip(), clean_password)
                server.sendmail(sender_email.strip(), all_targets, msg.as_string())
        else:
            with smtplib.SMTP(smtp_server, int(smtp_port), timeout=20) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(sender_email.strip(), clean_password)
                server.sendmail(sender_email.strip(), all_targets, msg.as_string())
        return True, "Email successfully sent!"
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. For Gmail, you must use a 16-character App Password (not your normal account password). Enable 2-Step Verification in your Google Account and generate an App Password."
    except Exception as e:
        return False, str(e)

def save_email_to_gmail_drafts(
    subject: str,
    html_content: str,
    plain_content: str,
    sender_name: str,
    sender_email: str,
    sender_password: str,
    recipients: list,
    cc_recipients: list = None
) -> tuple[bool, str]:
    """Deposits the formatted email directly into Gmail Drafts via IMAP, enabling native cloud Schedule Send."""
    if not sender_email or not sender_email.strip():
        return False, "Sender email is empty. Please enter your email."
    if not sender_password or not sender_password.strip():
        return False, "Gmail App Password is required."
    if not recipients:
        return False, "No recipient email addresses specified."

    clean_recipients = [r.strip() for r in recipients if r and r.strip()]
    clean_cc = [c.strip() for c in (cc_recipients or []) if c and c.strip()]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["To"] = ", ".join(clean_recipients)
    if clean_cc:
        msg["Cc"] = ", ".join(clean_cc)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()

    msg.attach(MIMEText(plain_content, "plain", "utf-8"))
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    clean_pwd = sender_password.strip().replace(" ", "")

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993, timeout=20)
        mail.login(sender_email.strip(), clean_pwd)
        status, folders = mail.list()
        drafts_folder = '"[Gmail]/Drafts"'
        for f in folders:
            f_str = f.decode()
            if r"\Drafts" in f_str:
                parts = f_str.split(' "/" ')
                if len(parts) > 1:
                    drafts_folder = parts[1].strip()
                break

        res = mail.append(drafts_folder, r"\Draft", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        mail.logout()
        if res[0] == "OK":
            return True, "Draft successfully created in your Gmail!"
        return False, str(res)
    except Exception as e:
        return False, f"Could not create Gmail Draft: {e}"
