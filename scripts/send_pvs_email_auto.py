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
import math
import io
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

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

try:
    from PIL import Image, ImageDraw, ImageFont  # type: ignore[import]
    HAVE_PIL = True
except Exception as e:  # pragma: no cover - optional dependency
    Image = ImageDraw = ImageFont = None  # type: ignore[assignment]
    HAVE_PIL = False
    logger.warning(f"Pillow import failed; will try matplotlib fallback for donut PNGs: {e}")

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


def generate_email_html(data: dict) -> tuple[str, dict[str, bytes]]:
    """Generate the HTML email body from PVS data with category coloring."""
    
    rows = data.get('rows', [])
    totals = data.get('totals', {})
    group_totals = data.get('group_totals', [])
    olk_totals = data.get('olk_totals', {})
    report_date = data.get('date', str(date.today() - timedelta(days=1)))
    gauge_images: dict[str, bytes] = {}

    def _adh_color_rgb(abs_pct: float) -> tuple[int, int, int]:
        """Shared adherence color: 100%% green, <100 red→green, >100 green→purple."""
        try:
            pct_val = float(abs_pct)
        except (TypeError, ValueError):
            pct_val = 0.0

        if not math.isfinite(pct_val):
            pct_val = 0.0

        green = (46, 204, 113)
        yellow = (241, 196, 15)
        red = (231, 76, 60)

        diff = abs(pct_val - 100.0)
        if diff <= 3.0:
            return green
        if diff <= 5.0:
            return yellow
        return red

    def adh_style(pct):
        """PVS adherence color: pct is delta/schedule*100 (0 = on plan)."""
        if pct is None:
            return "background-color: rgba(255,255,255,0.08); color: #dddddd;"
        try:
            val = float(pct)
        except (TypeError, ValueError):
            val = 0.0
        if not math.isfinite(val):
            val = 0.0

        green = (46, 204, 113)
        yellow = (241, 196, 15)
        red = (231, 76, 60)

        abs_val = abs(val)
        if abs_val <= 3.0:
            r_val, g_val, b_val = green
            text = '#ffffff'
        elif abs_val < 5.0:
            r_val, g_val, b_val = yellow
            text = '#000000'
        else:
            r_val, g_val, b_val = red
            text = '#ffffff'

        return f"background-color: rgb({r_val}, {g_val}, {b_val}); color: {text};"

    def delta_style(d):
        return "color: #28a745;" if d >= 0 else "color: #dc3545;"

    def fmt_int(n):
        return f"{int(n):,}" if n else "0"

    def fmt_pct(n):
        if n is None:
            return "N/A"
        try:
            f = float(n)
            if f != f or f == float('inf') or f == float('-inf'):
                return "N/A"
            return f"{int(round(f))}%"
        except (TypeError, ValueError):
            return "N/A"

    def olk_adh_style(pct: float) -> str:
        """OLK adherence color: pct is already absolute %% vs OLK."""
        val = pct or 0.0
        r_val, g_val, b_val = _adh_color_rgb(val)
        diff = abs(float(val) - 100.0) if math.isfinite(float(val)) else 999.0
        text = '#000000' if (diff > 3.0 and diff <= 5.0) else '#ffffff'
        return f"background-color: rgb({r_val}, {g_val}, {b_val}); color: {text};"

    def _abs_pct_and_fill(schedule, production):
        try:
            s = float(schedule or 0)
            p = float(production or 0)
        except (TypeError, ValueError):
            return None, 0.0
        if s <= 0:
            return None, 0.0
        abs_pct = (p / s) * 100.0
        fill = min(max(p / s, 0.0), 1.0)
        return abs_pct, fill

    def _donut_svg(schedule, production, abs_pct_override=None) -> str:
        abs_pct, fill = _abs_pct_and_fill(schedule, production)
        if abs_pct_override is not None:
            try:
                abs_pct = float(abs_pct_override)
            except (TypeError, ValueError):
                pass

        r = 50
        stroke = 16
        circ = 2 * math.pi * r
        dash = circ * fill

        if abs_pct is None:
            text = 'N/A'
            color = '#dddddd'
            arc = ''
        else:
            r_val, g_val, b_val = _adh_color_rgb(abs_pct)
            color = f"rgb({r_val}, {g_val}, {b_val})"
            text = f"{int(round(abs_pct))}%"
            arc = (
                f"<circle cx=\"60\" cy=\"60\" r=\"{r}\" fill=\"none\" "
                f"stroke=\"{color}\" stroke-width=\"{stroke}\" "
                f"stroke-dasharray=\"{dash:.1f} {circ:.1f}\" stroke-linecap=\"round\" "
                f"transform=\"rotate(-90 60 60)\"/>"
            )

        return (
            f"<svg width=\"120\" height=\"120\" viewBox=\"0 0 120 120\" xmlns=\"http://www.w3.org/2000/svg\">"
            f"<circle cx=\"60\" cy=\"60\" r=\"{r}\" fill=\"none\" stroke=\"#02252b\" stroke-width=\"{stroke}\"/>"
            f"{arc}"
            f"<text x=\"60\" y=\"60\" text-anchor=\"middle\" dominant-baseline=\"middle\" "
            f"font-family=\"Segoe UI, Arial, sans-serif\" font-size=\"22\" font-weight=\"700\" fill=\"{color}\">{text}</text>"
            f"</svg>"
        )

    chart_index = 0

    def _cid_slug(text: str) -> str:
        s = ''.join(c.lower() if c.isalnum() else '_' for c in str(text))
        s = s.strip('_')
        return s or 'chart'

    def _render_donut_png(schedule, production, abs_pct_override=None) -> bytes:
        abs_pct, fill = _abs_pct_and_fill(schedule, production)
        if abs_pct_override is not None:
            try:
                abs_pct = float(abs_pct_override)
            except (TypeError, ValueError):
                pass

        if not HAVE_PIL or Image is None or ImageDraw is None:  # type: ignore[truthy-function]
            try:
                os.environ.setdefault('MPLBACKEND', 'Agg')
                from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas  # type: ignore[import]
                from matplotlib.figure import Figure  # type: ignore[import]
                from matplotlib.patches import Wedge  # type: ignore[import]
            except Exception as e:
                raise RuntimeError("No renderer available for donut rendering (install Pillow or matplotlib)") from e

            size = 120
            dpi = 100
            fig = Figure(figsize=(size / dpi, size / dpi), dpi=dpi)
            fig.patch.set_facecolor((8 / 255, 52 / 255, 58 / 255, 1))
            canvas = FigureCanvas(fig)
            ax = fig.add_axes([0, 0, 1, 1])
            ax.set_xlim(-1, 1)
            ax.set_ylim(-1, 1)
            ax.set_aspect('equal')
            ax.axis('off')

            base_color = (2 / 255, 37 / 255, 43 / 255, 1)
            ring_width = 0.32
            ax.add_patch(Wedge((0, 0), 1, 0, 360, width=ring_width, facecolor=base_color, edgecolor=base_color))

            if abs_pct is None:
                text = 'N/A'
                arc_color = (221 / 255, 221 / 255, 221 / 255, 1)
                ax.add_patch(Wedge((0, 0), 1, 0, 360, width=ring_width, facecolor=arc_color, edgecolor=arc_color))
            else:
                r_val, g_val, b_val = _adh_color_rgb(abs_pct)
                arc_color = (r_val / 255, g_val / 255, b_val / 255, 1)
                angle = max(0.0, min(fill, 1.0)) * 360.0
                if angle > 0.0:
                    start = -90.0
                    end = start + (359.9 if angle >= 360.0 else angle)
                    ax.add_patch(Wedge((0, 0), 1, start, end, width=ring_width, facecolor=arc_color, edgecolor=arc_color))
                text = f"{int(round(abs_pct))}%"

            ax.text(0, 0, text, ha='center', va='center', fontsize=22, fontweight='bold', color=arc_color)

            buf = io.BytesIO()
            canvas.print_png(buf)
            return buf.getvalue()

        size = 120
        cx = cy = size // 2
        r = 50
        stroke = 16
        bbox = (cx - r, cy - r, cx + r, cy + r)

        bg = (8, 52, 58, 255)
        img = Image.new("RGBA", (size, size), bg)
        draw = ImageDraw.Draw(img)

        base_color = (2, 37, 43, 255)
        draw.arc(bbox, start=0, end=359.9, fill=base_color, width=stroke)

        if abs_pct is None:
            text = 'N/A'
            arc_color = (221, 221, 221, 255)
        else:
            r_val, g_val, b_val = _adh_color_rgb(abs_pct)
            arc_color = (r_val, g_val, b_val, 255)
            angle = max(0.0, min(fill, 1.0)) * 360.0
            if angle > 0.0:
                start = -90.0
                end = start + (359.9 if angle >= 360.0 else angle)
                draw.arc(bbox, start=start, end=end, fill=arc_color, width=stroke)
            text = f"{int(round(abs_pct))}%"

        try:
            font = ImageFont.truetype("segoeui.ttf", 22)  # type: ignore[arg-type]
        except Exception:
            font = ImageFont.load_default()

        try:
            tb = draw.textbbox((0, 0), text, font=font)
            tw, th = tb[2] - tb[0], tb[3] - tb[1]
        except Exception:
            tw, th = draw.textsize(text, font=font)

        tx = cx - tw / 2
        ty = cy - th / 2
        draw.text((tx, ty), text, font=font, fill=arc_color)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    def _donut_img_tag(key: str, schedule, production, abs_pct_override=None) -> str:
        nonlocal chart_index

        abs_pct, _ = _abs_pct_and_fill(schedule, production)
        display_pct = abs_pct
        if abs_pct_override is not None:
            try:
                display_pct = float(abs_pct_override)
            except (TypeError, ValueError):
                pass

        if display_pct is None:
            alt = "N/A"
        else:
            alt = f"{int(round(display_pct))}%"

        chart_index += 1
        cid = f"d_{chart_index}_{_cid_slug(key)}"

        try:
            png_bytes = _render_donut_png(schedule, production, abs_pct_override=abs_pct_override)
        except Exception as e:  # pragma: no cover - defensive
            logger.warning(f"Donut PNG render failed for {key}: {e}")
            return (
                f"<div style=\"font-size:22px; font-weight:700; color:#ffffff; margin:8px 0;\">"
                f"{alt}</div>"
            )

        gauge_images[cid] = png_bytes
        return (
            f"<img src=\"cid:{cid}\" alt=\"{alt}\" width=\"120\" height=\"120\" "
            "style=\"display:block;margin:0 auto;width:120px;height:120px;\" />"
        )

    def _chart_td(title: str, title_color: str, sched_label: str, sched, prod_label: str, prod, abs_pct_override=None) -> str:
        chart = _donut_img_tag(title, sched, prod, abs_pct_override=abs_pct_override)
        return (
            f"<td style=\"width:25%; vertical-align:top; text-align:center; padding:6px 4px;\">"
            f"<div style=\"font-weight:700; color:{title_color}; font-size:13px; margin-bottom:6px;\">{title}</div>"
            f"{chart}"
            f"<div style=\"font-size:11px; color:#a0c0c4; margin-top:4px;\">{sched_label}: {fmt_int(sched)} | {prod_label}: {fmt_int(prod)}</div>"
            f"</td>"
        )

    def _charts_block(tds_html: str) -> str:
        return (
            "<div style=\"background:#08343a; border-radius:8px; padding:10px 8px; margin:6px 0 10px;\">"
            "<table style=\"width:100%; border-collapse:collapse; table-layout:fixed;\"><tr>"
            + tds_html
            + "</tr></table></div>"
        )

    def build_line_row(r: dict) -> str:
        """Build a SEW/ASSY line row with MTD OLK columns and WTD/Daily PVS."""
        category = r.get('category', 'OTHER')
        cat_style = get_category_style(category)
        mtd = r.get('mtd', {})
        wtd = r.get('wtd', {})
        daily = r.get('daily', {})

        mtd_adh = mtd.get('adherence_pct')
        wtd_adh = wtd.get('adherence_pct')
        daily_adh = daily.get('adherence_pct')

        mtd_olk = mtd.get('olk', 0) or 0

        return f"""
        <tr style="{cat_style}">
            <td style="text-align: left; font-weight: 600; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{r.get('line', '')}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd_olk)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(mtd.get('delta', 0))}">{fmt_int(mtd.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(mtd_adh)}">{fmt_pct(mtd_adh)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(wtd.get('delta', 0))}">{fmt_int(wtd.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(wtd_adh)}">{fmt_pct(wtd_adh)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(daily.get('delta', 0))}">{fmt_int(daily.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(daily_adh)}">{fmt_pct(daily_adh)}</td>
        </tr>
        """

    def build_group_row(g: dict) -> str:
        """Build a grouped totals row, including MTD OLK columns."""
        mtd = g.get('mtd', {})
        wtd = g.get('wtd', {})
        daily = g.get('daily', {})

        mtd_adh = mtd.get('adherence_pct')
        wtd_adh = wtd.get('adherence_pct')
        daily_adh = daily.get('adherence_pct')

        mtd_olk = mtd.get('olk', 0) or 0

        return f"""
        <tr>
            <td style="text-align: left; font-weight: 600; padding: 8px 12px; border-bottom: 1px solid #1a4a50;">{g.get('group', '')}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd_olk)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(mtd.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(mtd.get('delta', 0))}">{fmt_int(mtd.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(mtd_adh)}">{fmt_pct(mtd_adh)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(wtd.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(wtd.get('delta', 0))}">{fmt_int(wtd.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(wtd_adh)}">{fmt_pct(wtd_adh)}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily.get('schedule', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50;">{fmt_int(daily.get('production', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {delta_style(daily.get('delta', 0))}">{fmt_int(daily.get('delta', 0))}</td>
            <td style="text-align: right; padding: 8px 10px; border-bottom: 1px solid #1a4a50; {adh_style(daily_adh)}">{fmt_pct(daily_adh)}</td>
        </tr>
        """

    sew_rows = [r for r in rows if r.get('category') == 'SEW']
    assy_rows = [r for r in rows if r.get('category') == 'ASSY']

    sew_table_rows = "".join(build_line_row(r) for r in sew_rows)
    assy_table_rows = "".join(build_line_row(r) for r in assy_rows)
    group_table_rows = "".join(build_group_row(g) for g in group_totals)
    
    # Build totals summary for MTD pie chart section (text-based for email)
    sew_mtd = totals.get('sew', {}).get('mtd', {})
    sew_wtd = totals.get('sew', {}).get('wtd', {})
    sew_daily = totals.get('sew', {}).get('daily', {})

    assy_mtd = totals.get('assy', {}).get('mtd', {})
    assy_wtd = totals.get('assy', {}).get('wtd', {})
    assy_daily = totals.get('assy', {}).get('daily', {})

    sew_olk = olk_totals.get('sew', {}) if isinstance(olk_totals, dict) else {}
    assy_olk = olk_totals.get('assy', {}) if isinstance(olk_totals, dict) else {}

    sew_charts_html = _charts_block(
        _chart_td('SEW MTD', '#ff9999', 'Sched', sew_mtd.get('schedule', 0), 'Prod', sew_mtd.get('production', 0))
        + _chart_td('SEW WTD', '#ff9999', 'Sched', sew_wtd.get('schedule', 0), 'Prod', sew_wtd.get('production', 0))
        + _chart_td('SEW Daily', '#ff9999', 'Sched', sew_daily.get('schedule', 0), 'Prod', sew_daily.get('production', 0))
        + _chart_td(
            'SEW MTD OLK',
            '#c8e000',
            'OLK',
            sew_olk.get('olk', 0),
            'Prod',
            sew_olk.get('production', 0),
            abs_pct_override=sew_olk.get('adh_olk_pct', 0),
        )
    )

    assy_charts_html = _charts_block(
        _chart_td('ASSY MTD', '#7dc8e8', 'Sched', assy_mtd.get('schedule', 0), 'Prod', assy_mtd.get('production', 0))
        + _chart_td('ASSY WTD', '#7dc8e8', 'Sched', assy_wtd.get('schedule', 0), 'Prod', assy_wtd.get('production', 0))
        + _chart_td('ASSY Daily', '#7dc8e8', 'Sched', assy_daily.get('schedule', 0), 'Prod', assy_daily.get('production', 0))
        + _chart_td(
            'ASSY MTD OLK',
            '#c8e000',
            'OLK',
            assy_olk.get('olk', 0),
            'Prod',
            assy_olk.get('production', 0),
            abs_pct_override=assy_olk.get('adh_olk_pct', 0),
        )
    )

    group_charts_html = ''
    group_filtered = [g for g in group_totals if g and g.get('group') and g.get('group') != 'OTHER']
    for i in range(0, len(group_filtered), 4):
        chunk = group_filtered[i:i+4]
        tds = ''
        for g in chunk:
            mtd = g.get('mtd', {}) or {}
            tds += _chart_td(
                g.get('group', ''),
                '#c8e000',
                'Sched',
                mtd.get('schedule', 0),
                'Prod',
                mtd.get('production', 0),
            )
        for _ in range(4 - len(chunk)):
            tds += '<td style="width:25%; padding:6px 4px;"></td>'
        group_charts_html += _charts_block(tds)

    table_header_line = """
                <thead>
                    <tr>
                        <th rowspan="2" style="text-align: left;">Line</th>
                        <th class="group mtd" colspan="5">MTD</th>
                        <th class="group wtd" colspan="4">WTD</th>
                        <th class="group daily" colspan="4">Daily PVS</th>
                    </tr>
                    <tr>
                        <th class="mtd">OLK</th><th class="mtd">Sch</th><th class="mtd">Prod</th><th class="mtd">Delta</th><th class="mtd">Adh %</th>
                        <th class="wtd">Schedule</th><th class="wtd">Production</th><th class="wtd">Delta</th><th class="wtd">Adh %</th>
                        <th class="daily">Sched</th><th class="daily">Prod</th><th class="daily">Delta</th><th class="daily">Adh %</th>
                    </tr>
                </thead>
    """

    table_header_group = """
                <thead>
                    <tr>
                        <th rowspan="2" style="text-align: left;">Group</th>
                        <th class="group mtd" colspan="5">MTD</th>
                        <th class="group wtd" colspan="4">WTD</th>
                        <th class="group daily" colspan="4">Daily PVS</th>
                    </tr>
                    <tr>
                        <th class="mtd">OLK</th><th class="mtd">Sch</th><th class="mtd">Prod</th><th class="mtd">Delta</th><th class="mtd">Adh %</th>
                        <th class="wtd">Schedule</th><th class="wtd">Production</th><th class="wtd">Delta</th><th class="wtd">Adh %</th>
                        <th class="daily">Sched</th><th class="daily">Prod</th><th class="daily">Delta</th><th class="daily">Adh %</th>
                    </tr>
                </thead>
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
                background: #08343a;
                border: 1px solid #0f5560;
            }}
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
            .section-title {{
                color: #c8e000;
                margin: 18px 0 6px 0;
                font-size: 18px;
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
            
            <!-- Page 1: SEW -->
            <h2 class="section-title">SEW Lines</h2>
            {sew_charts_html}
            <table>
                {table_header_line}
                <tbody>
                    {sew_table_rows}
                </tbody>
            </table>

            <!-- Page 2: ASSY -->
            <h2 class="section-title">ASSY Lines</h2>
            {assy_charts_html}
            <table>
                {table_header_line}
                <tbody>
                    {assy_table_rows}
                </tbody>
            </table>

            <!-- Page 3: Grouped Totals -->
            <h2 class="section-title">Grouped Totals</h2>
            {group_charts_html}
            <table>
                {table_header_group}
                <tbody>
                    {group_table_rows}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html, gauge_images


def send_email_smtp(subject: str, html_body: str, recipients: list[str], inline_images: dict[str, bytes] | None = None) -> bool:
    """Send HTML email via SMTP only (no Outlook)."""
    
    msg = MIMEMultipart('related')
    msg['Subject'] = subject
    msg['From'] = EMAIL_CONFIG['from_address']
    msg['To'] = ', '.join(recipients)

    alt = MIMEMultipart('alternative')
    msg.attach(alt)
    
    html_part = MIMEText(html_body, 'html', 'utf-8')
    alt.attach(html_part)

    if inline_images:
        for cid, img_bytes in inline_images.items():
            try:
                img = MIMEImage(img_bytes, _subtype='png')
            except TypeError:
                img = MIMEImage(img_bytes)
            img.add_header('Content-ID', f"<{cid}>")
            img.add_header('Content-Disposition', 'inline', filename=f"{cid}.png")
            img.add_header('X-Attachment-Id', cid)
            msg.attach(img)
    
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
    html_body, inline_images = generate_email_html(data)
    logger.info(f"Inline images generated: {len(inline_images)}")
    
    # Build subject
    yesterday = date.today() - timedelta(days=1)
    subject = f"PVS Report ({yesterday.strftime('%Y-%m-%d')})"
    if test_mode:
        subject = "[TEST] " + subject
    
    logger.info(f"Subject: {subject}")
    
    # Send email
    success = send_email_smtp(subject, html_body, recipients, inline_images=inline_images)
    
    if success:
        logger.info("Email sent successfully!")
        return 0
    else:
        logger.error("Email sending failed")
        return 1


if __name__ == '__main__':
    sys.exit(main())
