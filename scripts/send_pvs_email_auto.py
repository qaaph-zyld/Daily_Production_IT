#!/usr/bin/env python3
"""
Automated PVS Report Email (SMTP-only, no Outlook dependency)

This script is designed to run on a VM as a scheduled task at 8:45 AM daily.
It sends the PVS report via SMTP without requiring Outlook.

Usage:
    python send_pvs_email_auto.py [--test]

Options:
    --test    Send to a test recipient instead of the full list
"""

import os
import sys
import json
import smtplib
import logging
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Setup logging
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
LOG_DIR = os.path.join(PROJECT_DIR, 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, 'email_auto.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Add parent directory to path so we can import pvs_server
sys.path.insert(0, PROJECT_DIR)

from pvs_server import compute_metrics

# Email configuration defaults
EMAIL_CONFIG = {
    'smtp_server': 'smtp.adient.com',
    'smtp_port': 25,
    'from_address': 'pvs-dashboard@adient.com',
    'recipients': [],
    'test_recipient': 'nikola.jelacic@adient.com',
}

# Load settings from settings.json
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
        if email_settings.get('test_recipient'):
            EMAIL_CONFIG['test_recipient'] = email_settings['test_recipient']
        logger.info(f"Loaded email settings from {SETTINGS_PATH}")
    except Exception as e:
        logger.warning(f"Could not load email settings: {e}")


def get_category_style(category: str) -> str:
    """Return inline CSS for row based on category."""
    if category == 'SEW':
        return "background-color: rgba(180, 60, 60, 0.25);"
    elif category == 'ASSY':
        return "background-color: rgba(40, 100, 140, 0.25);"
    else:
        return "background-color: rgba(80, 80, 80, 0.20);"


def generate_email_html(data: dict) -> str:
    """Generate the HTML email body from PVS data with category coloring."""
    
    rows = data.get('rows', [])
    totals = data.get('totals', {})
    report_date = data.get('date', str(date.today() - timedelta(days=1)))
    
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
        return "color: #28a745;" if d >= 0 else "color: #dc3545;"
    
    def fmt_int(n):
        return f"{int(n):,}" if n else "0"
    
    def fmt_pct(n):
        return f"{int(n)}%"
    
    # Build table rows with category coloring
    table_rows = ""
    last_category = None
    
    for r in rows:
        category = r.get('category', 'OTHER')
        
        # Add separator between categories
        if last_category and category != last_category:
            table_rows += '<tr style="height: 6px; background: #0a3b41;"><td colspan="13"></td></tr>'
        last_category = category
        
        cat_style = get_category_style(category)
        mtd = r['mtd']
        wtd = r['wtd']
        daily = r['daily']
        
        table_rows += f"""
        <tr style="{cat_style}">
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
    
    # Build totals summary for pie chart section (text-based for email)
    sew_s = totals.get('sew', {}).get('schedule', 0)
    sew_p = totals.get('sew', {}).get('production', 0)
    sew_pct = round((sew_p / sew_s * 100) if sew_s > 0 else 0)
    
    assy_s = totals.get('assy', {}).get('schedule', 0)
    assy_p = totals.get('assy', {}).get('production', 0)
    assy_pct = round((assy_p / assy_s * 100) if assy_s > 0 else 0)
    
    total_s = totals.get('total', {}).get('schedule', 0)
    total_p = totals.get('total', {}).get('production', 0)
    total_pct = round((total_p / total_s * 100) if total_s > 0 else 0)
    
    def pct_color(pct):
        return '#2ecc71' if pct >= 100 else '#e74c3c'
    
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
                margin-bottom: 15px;
                font-size: 14px;
            }}
            .summary-row {{
                display: flex;
                justify-content: center;
                gap: 40px;
                margin-bottom: 20px;
                flex-wrap: wrap;
            }}
            .summary-box {{
                text-align: center;
                padding: 12px 20px;
                border-radius: 8px;
                min-width: 150px;
            }}
            .summary-box.sew {{ background: rgba(180, 60, 60, 0.3); }}
            .summary-box.assy {{ background: rgba(40, 100, 140, 0.3); }}
            .summary-box.total {{ background: rgba(200, 224, 0, 0.2); }}
            .summary-box h4 {{
                margin: 0 0 8px;
                font-size: 14px;
            }}
            .summary-box.sew h4 {{ color: #ff9999; }}
            .summary-box.assy h4 {{ color: #7dc8e8; }}
            .summary-box.total h4 {{ color: #c8e000; }}
            .summary-box .pct {{
                font-size: 28px;
                font-weight: 700;
            }}
            .summary-box .details {{
                font-size: 11px;
                color: #a0c0c4;
                margin-top: 4px;
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
                color: #e8f1f2;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>ADIENT LOZNICA PVS</h1>
            <div class="subtitle">Report Date: {report_date}</div>
            
            <!-- MTD Summary -->
            <table style="width: auto; margin: 0 auto 20px; background: transparent;">
                <tr>
                    <td style="padding: 0 15px;">
                        <div class="summary-box sew">
                            <h4>SEW MTD</h4>
                            <div class="pct" style="color: {pct_color(sew_pct)};">{sew_pct}%</div>
                            <div class="details">Sched: {fmt_int(sew_s)} | Prod: {fmt_int(sew_p)}</div>
                        </div>
                    </td>
                    <td style="padding: 0 15px;">
                        <div class="summary-box assy">
                            <h4>ASSY MTD</h4>
                            <div class="pct" style="color: {pct_color(assy_pct)};">{assy_pct}%</div>
                            <div class="details">Sched: {fmt_int(assy_s)} | Prod: {fmt_int(assy_p)}</div>
                        </div>
                    </td>
                    <td style="padding: 0 15px;">
                        <div class="summary-box total">
                            <h4>TOTAL MTD</h4>
                            <div class="pct" style="color: {pct_color(total_pct)};">{total_pct}%</div>
                            <div class="details">Sched: {fmt_int(total_s)} | Prod: {fmt_int(total_p)}</div>
                        </div>
                    </td>
                </tr>
            </table>
            
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


def send_email_smtp(subject: str, html_body: str, recipients: list[str]) -> bool:
    """Send HTML email via SMTP only (no Outlook)."""
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = EMAIL_CONFIG['from_address']
    msg['To'] = ', '.join(recipients)
    
    html_part = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(html_part)
    
    try:
        logger.info(f"Connecting to SMTP server {EMAIL_CONFIG['smtp_server']}:{EMAIL_CONFIG['smtp_port']}...")
        with smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'], timeout=60) as server:
            server.sendmail(
                EMAIL_CONFIG['from_address'],
                recipients,
                msg.as_string()
            )
        logger.info(f"Email sent successfully to {len(recipients)} recipients")
        return True
    except Exception as e:
        logger.error(f"Failed to send email via SMTP: {e}")
        return False


def main():
    # Check for test mode
    test_mode = '--test' in sys.argv
    
    logger.info("=" * 60)
    logger.info("PVS Automated Email Report")
    logger.info("=" * 60)
    
    if test_mode:
        logger.info("Running in TEST mode - sending to test recipient only")
        recipients = [EMAIL_CONFIG['test_recipient']]
    else:
        recipients = EMAIL_CONFIG['recipients']
        if not recipients:
            logger.error("No recipients configured. Check settings.json")
            return 1
    
    logger.info(f"Recipients: {len(recipients)} addresses")
    
    # Compute metrics
    logger.info("Computing PVS metrics...")
    try:
        data = compute_metrics()
    except Exception as e:
        logger.error(f"Failed to compute metrics: {e}")
        return 1
    
    if not data.get('success'):
        logger.error("compute_metrics returned unsuccessful result")
        return 1
    
    logger.info(f"Metrics computed for date: {data.get('date')}")
    logger.info(f"Total rows: {len(data.get('rows', []))}")
    
    # Generate email HTML
    logger.info("Generating email HTML...")
    html_body = generate_email_html(data)
    
    # Build subject
    yesterday = date.today() - timedelta(days=1)
    subject = f"PVS Report ({yesterday.strftime('%Y-%m-%d')})"
    if test_mode:
        subject = "[TEST] " + subject
    
    logger.info(f"Subject: {subject}")
    
    # Send email
    success = send_email_smtp(subject, html_body, recipients)
    
    if success:
        logger.info("Email sent successfully!")
        return 0
    else:
        logger.error("Email sending failed")
        return 1


if __name__ == '__main__':
    sys.exit(main())
