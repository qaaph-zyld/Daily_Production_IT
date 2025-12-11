#!/usr/bin/env python3
"""
Send PVS Report Email

Sends the PVS dashboard snapshot as an HTML email to configured recipients.
Subject: PVS Report (YYYY-MM-DD) where date is yesterday's date.
"""

import os
import sys
import json
import smtplib
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    import win32com.client as win32
except Exception:
    win32 = None

# Add parent directory to path so we can import pvs_server
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

from pvs_server import compute_metrics

# Email configuration
EMAIL_CONFIG = {
    'smtp_server': 'smtp.adient.com',  # Adjust to your company's SMTP
    'smtp_port': 25,                    # Common corporate port (no auth)
    'from_address': 'pvs-dashboard@adient.com',
    'recipients': [
        'nikola.jelacic@adient.com',
    ],
}

# Load any overrides from settings.json
SETTINGS_PATH = os.path.join(PROJECT_DIR, 'config', 'settings.json')
if os.path.exists(SETTINGS_PATH):
    try:
        with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
            settings = json.load(f)
        email_settings = settings.get('email', {})
        if email_settings.get('smtp_server'):
            EMAIL_CONFIG['smtp_server'] = email_settings['smtp_server']
        if email_settings.get('smtp_port'):
            EMAIL_CONFIG['smtp_port'] = email_settings['smtp_port']
        if email_settings.get('from_address'):
            EMAIL_CONFIG['from_address'] = email_settings['from_address']
        if email_settings.get('recipients'):
            EMAIL_CONFIG['recipients'] = email_settings['recipients']
    except Exception as e:
        print(f"[EMAIL] Warning: Could not load email settings: {e}")


def generate_email_html(data: dict) -> str:
    """Generate the HTML email body from PVS data."""
    
    rows = data.get('rows', [])
    report_date = data.get('date', str(date.today() - timedelta(days=1)))
    
    # Build table rows
    table_rows = ""
    for r in rows:
        # Determine colors for adherence cells
        def adh_style(pct):
            deviation = abs(pct)
            max_dev = 50
            intensity = min(deviation / max_dev, 1)
            if pct >= 0:
                r_val = int(20 + (40 - 20) * intensity)
                g_val = int(30 + (167 - 30) * intensity)
                b_val = int(25 + (69 - 25) * intensity)
                text_color = '#d4ffd4' if intensity > 0.3 else '#aaa'
            else:
                r_val = int(30 + (220 - 30) * intensity)
                g_val = int(20 + (53 - 20) * intensity)
                b_val = int(20 + (69 - 20) * intensity)
                text_color = '#fff' if intensity > 0.3 else '#aaa'
            return f"background-color: rgb({r_val}, {g_val}, {b_val}); color: {text_color};"
        
        def delta_style(d):
            if d >= 0:
                return "color: #28a745;"
            else:
                return "color: #dc3545;"
        
        def fmt_int(n):
            return f"{int(n):,}" if n else "0"
        
        def fmt_pct(n):
            return f"{int(n)}%"
        
        mtd = r['mtd']
        wtd = r['wtd']
        daily = r['daily']
        
        table_rows += f"""
        <tr>
            <td style="text-align: left; font-weight: 600; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{r['line']}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd['schedule'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd['production'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {delta_style(mtd['delta'])}">{fmt_int(mtd['delta'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {adh_style(mtd['adherence_pct'])}">{fmt_pct(mtd['adherence_pct'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd['schedule'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd['production'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {delta_style(wtd['delta'])}">{fmt_int(wtd['delta'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {adh_style(wtd['adherence_pct'])}">{fmt_pct(wtd['adherence_pct'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily['schedule'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily['production'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {delta_style(daily['delta'])}">{fmt_int(daily['delta'])}</td>
            <td style="text-align: right; padding: 8px 12px; border-bottom: 1px solid #1a4a50; {adh_style(daily['adherence_pct'])}">{fmt_pct(daily['adherence_pct'])}</td>
        </tr>
        """
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: Segoe UI, Roboto, Arial, sans-serif;
                background: #f5f5f5;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background: #0a3b41;
                border-radius: 12px;
                padding: 20px;
            }}
            h1 {{
                color: #c8e000;
                text-align: center;
                margin: 0 0 5px 0;
                font-size: 24px;
            }}
            .subtitle {{
                color: #a0c0c4;
                text-align: center;
                margin-bottom: 20px;
                font-size: 14px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                background: #08343a;
                border-radius: 8px;
                overflow: hidden;
                font-size: 13px;
            }}
            th {{
                background: #0f5560;
                color: #e3f2f4;
                font-weight: 700;
                padding: 10px 12px;
                text-align: right;
            }}
            th:first-child {{
                text-align: left;
            }}
            th.group {{
                text-align: center;
                font-size: 14px;
                font-weight: 800;
            }}
            th.mtd {{ background: #0f6a70; }}
            th.wtd {{ background: #154f7a; }}
            th.daily {{ background: #6a5b16; }}
            td {{
                background: #0d2d32;
                color: #e8f1f2;
            }}
            tr:nth-child(even) td {{
                background: #0a2428;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>ADIENT LOZNICA PVS</h1>
            <div class="subtitle">Report Date: {report_date}</div>
            <table>
                <thead>
                    <tr>
                        <th rowspan="2" style="text-align: left;">Line</th>
                        <th class="group mtd" colspan="4">MTD</th>
                        <th class="group wtd" colspan="4">WTD</th>
                        <th class="group daily" colspan="4">Daily PVS</th>
                    </tr>
                    <tr>
                        <th class="mtd">Schedule</th><th class="mtd">Production</th><th class="mtd">Delta</th><th class="mtd">Adh %</th>
                        <th class="wtd">Schedule</th><th class="wtd">Production</th><th class="wtd">Delta</th><th class="wtd">Adh %</th>
                        <th class="daily">Sched</th><th class="daily">Prod</th><th class="daily">Delta</th><th class="daily">Adh %</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html


def send_email(subject: str, html_body: str, recipients: list[str]) -> bool:
    """Send HTML email, preferring local Outlook, falling back to SMTP."""

    # 1) Try Outlook (planner's own mailbox)
    if win32 is not None:
        try:
            print("[EMAIL] Using Outlook to create email draft...")
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0 = MailItem
            mail.Subject = subject
            mail.To = '; '.join(recipients)
            mail.HTMLBody = html_body
            # Show email to user so they can review and click Send
            mail.Display(True)
            print("[EMAIL] Draft opened in Outlook. Please review and click Send.")
            return True
        except Exception as e:
            print(f"[EMAIL] WARNING: Outlook path failed, falling back to SMTP: {e}")

    # 2) Fallback: direct SMTP send
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = EMAIL_CONFIG['from_address']
    msg['To'] = ', '.join(recipients)

    html_part = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(html_part)

    try:
        print(f"[EMAIL] Connecting to {EMAIL_CONFIG['smtp_server']}:{EMAIL_CONFIG['smtp_port']}...")
        with smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'], timeout=30) as server:
            server.sendmail(
                EMAIL_CONFIG['from_address'],
                recipients,
                msg.as_string()
            )
        print(f"[EMAIL] Successfully sent via SMTP to: {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"[EMAIL] ERROR: Failed to send via SMTP: {e}")
        return False


def main():
    print("[EMAIL] Generating PVS report for email...")
    
    # Compute metrics (same data as dashboard)
    data = compute_metrics()
    
    if not data.get('success'):
        print("[EMAIL] ERROR: Failed to compute metrics")
        return 1
    
    # Yesterday's date for subject
    yesterday = date.today() - timedelta(days=1)
    subject = f"PVS Report ({yesterday.strftime('%Y-%m-%d')})"
    
    # Generate HTML body
    html_body = generate_email_html(data)
    
    # Send email
    recipients = EMAIL_CONFIG['recipients']
    print(f"[EMAIL] Sending to: {', '.join(recipients)}")
    print(f"[EMAIL] Subject: {subject}")
    
    success = send_email(subject, html_body, recipients)
    
    return 0 if success else 1


if __name__ == '__main__':
    sys.exit(main())
