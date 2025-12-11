import os
import csv
import json
import pyodbc
import pandas as pd
from flask import Flask, render_template, jsonify
from flask_cors import CORS
from datetime import datetime, date, timedelta
from dotenv import load_dotenv
from decimal import Decimal
try:
    import win32com.client as win32
    import pythoncom
except Exception:
    win32 = None
    pythoncom = None

# Load .env
load_dotenv()

# Load settings.json if available (optional config override)
SETTINGS = {}
SETTINGS_PATH = os.path.join(os.path.dirname(__file__), 'config', 'settings.json')
if os.path.exists(SETTINGS_PATH):
    try:
        with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
            SETTINGS = json.load(f)
        print(f"[CONFIG] Loaded settings from {SETTINGS_PATH}")
    except Exception as e:
        print(f"[CONFIG] Warning: Could not load settings.json: {e}")

app = Flask(__name__)
CORS(app)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.jinja_env.cache = {}

# DB config
DB_CONFIG = {
    'server': os.getenv('DB_SERVER', 'a265m001'),
    'database': os.getenv('DB_DATABASE', 'QADEE2798'),
    'username': os.getenv('DB_USERNAME', 'PowerBI'),
    'password': os.getenv('DB_PASSWORD', 'P0werB1'),
}

# Server config
FLASK_HOST = os.getenv('FLASK_HOST', '0.0.0.0')
PVS_PORT = int(os.getenv('PVS_PORT', '5051'))

# PVS config
_BASE_DIR = os.path.dirname(__file__)
_planned_default = os.path.join('PVS', 'Planned_qtys.xlsx')
_map_default = os.path.join('PVS', 'ProdLine_Project_Map.csv')

PVS_PLANNED_XLSX = os.getenv('PVS_PLANNED_XLSX', _planned_default)
if not os.path.isabs(PVS_PLANNED_XLSX):
    PVS_PLANNED_XLSX = os.path.join(_BASE_DIR, PVS_PLANNED_XLSX)

PVS_RECALC_XLSX = os.getenv('PVS_RECALC_XLSX', 'false').strip().lower() in ('1', 'true', 'yes')
PVS_ADHERENCE_CLAMP = float(os.getenv('PVS_ADHERENCE_CLAMP', '300'))  # percent cap; values beyond are treated as 0%

PVS_MAP_CSV = os.getenv('PVS_MAP_CSV', _map_default)
if not os.path.isabs(PVS_MAP_CSV):
    PVS_MAP_CSV = os.path.join(_BASE_DIR, PVS_MAP_CSV)

# External WH Receipt workbook (used to avoid Excel on VM)
_DATA_SOURCES = SETTINGS.get('dataSources', {}) if isinstance(SETTINGS, dict) else {}
PVS_EXTERNAL_XLSX = _DATA_SOURCES.get(
    'externalSourceExcel',
    r"G:\Logistics\6_Reporting\1_PVS\WH Receipt FY25.xlsx",
)
PVS_EXTERNAL_SHEET = _DATA_SOURCES.get('externalSheetName', 'Daily PVS')
PVS_EXTERNAL_TARGET_LABEL = _DATA_SOURCES.get('externalTargetLabel', 'Target (LTP input)')

_BEHAVIOR = SETTINGS.get('behavior', {}) if isinstance(SETTINGS, dict) else {}
PVS_USE_WH_RECEIPT = bool(_BEHAVIOR.get('useWhReceiptForPlan', True))


def norm_code(code: str) -> str:
    return (code or '').strip().upper()


def load_map_csv(path: str) -> dict[str, str]:
    """Return dict: {prod_line_code -> project_name} with uppercase codes."""
    mapping: dict[str, str] = {}
    if not os.path.exists(path):
        return mapping
    with open(path, 'r', newline='', encoding='utf-8') as f:
        rdr = csv.DictReader(f)
        for row in rdr:
            code = norm_code(row.get('prod_line', ''))
            proj = (row.get('project') or '').strip()
            if code and proj:
                mapping[code] = proj
    return mapping


def _choose_sql_driver():
    try:
        drivers = list(pyodbc.drivers())
    except Exception:
        drivers = []
    for name in (
        'ODBC Driver 18 for SQL Server',
        'ODBC Driver 17 for SQL Server',
        'SQL Server',
    ):
        if name in drivers:
            return name
    return 'SQL Server'


def get_db_connection():
    driver = _choose_sql_driver()
    extra = 'TrustServerCertificate=yes;'
    if 'ODBC Driver 18' in driver:
        extra = 'Encrypt=no;TrustServerCertificate=yes;'
    conn_str = (
        f"DRIVER={{{driver}}};SERVER={DB_CONFIG['server']};DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};PWD={DB_CONFIG['password']};{extra}"
    )
    return pyodbc.connect(conn_str)


def _parse_date_ddmmyyyy(s: str) -> date | None:
    s = s.strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, '%d/%m/%Y').date()
    except Exception:
        return None


def _coerce_header_to_date(col) -> date | None:
    try:
        # Skip NaN/None early
        if col is None or (isinstance(col, float) and pd.isna(col)):
            return None
        # Direct types
        if isinstance(col, date) and not isinstance(col, datetime):
            return col
        if isinstance(col, datetime):
            return col.date()
        # Pandas Timestamp
        if 'Timestamp' in type(col).__name__:
            try:
                if pd.isna(col):
                    return None
                return col.date()
            except Exception:
                pass
        # Excel serial number (float/int) - must be a reasonable date range
        if isinstance(col, (int, float)):
            # Excel dates for 2020-2030 are roughly 43831 to 47848
            if 40000 < col < 50000:
                try:
                    d = pd.to_datetime(col, unit='D', origin='1899-12-30', errors='raise')
                    return d.date()
                except Exception:
                    pass
            return None
        # String formats: dd/mm/yyyy, ISO, or generic
        s = str(col).strip()
        d = _parse_date_ddmmyyyy(s)
        if d:
            return d
        try:
            return datetime.strptime(s, '%Y-%m-%d').date()
        except Exception:
            pass
        try:
            d2 = pd.to_datetime(s, dayfirst=True, errors='coerce')
            if pd.notna(d2):
                return d2.date()
        except Exception:
            pass
    except Exception:
        return None
    return None


def recalc_excel_workbook(path: str) -> bool:
    if not path or not os.path.exists(path):
        return False
    try:
        if win32 is None or pythoncom is None:
            return False
        pythoncom.CoInitialize()
        xl = win32.DispatchEx('Excel.Application')
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            try:
                xl.AskToUpdateLinks = False
            except Exception:
                pass
            wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=3)
            try:
                try:
                    links = wb.LinkSources(1)  # 1 = xlLinkTypeExcelLinks
                    if links:
                        for ln in links:
                            try:
                                wb.UpdateLink(ln, 1)
                            except Exception:
                                pass
                except Exception:
                    pass
                try:
                    wb.RefreshAll()
                except Exception:
                    pass
                try:
                    xl.CalculateUntilAsyncQueriesDone()
                except Exception:
                    pass
                try:
                    xl.CalculateFullRebuild()
                except Exception:
                    try:
                        xl.CalculateFull()
                    except Exception:
                        pass
                try:
                    wb.Save()
                except Exception:
                    pass
            finally:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
        finally:
            try:
                xl.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return True
    except Exception:
        return False


def load_planned_from_wh_receipt(path: str, sheet: str, target_label: str):
    """Build planned schedule directly from WH Receipt workbook.

    Reads the 'Daily PVS' (or configured) sheet and uses rows where the
    classification column equals target_label (e.g. 'Target (LTP input)').

    Returns dict: {prod_line_code: {date: qty_int}} where prod_line_code is the
    production line code (e.g. 'B_FG', 'H_FG', ...).
    """
    planned: dict[str, dict[date, int]] = {}
    if not path or not os.path.exists(path):
        print(f"[LOAD-WH] ERROR: External XLSX not found: {path}")
        return planned

    print(f"[LOAD-WH] Loading WH Receipt data from: {path} (sheet={sheet})")

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception as e:
        print(f"[LOAD-WH] ERROR opening workbook: {e}")
        return planned

    if df.empty:
        print("[LOAD-WH] Sheet is empty")
        return planned

    # 1) Detect header row that contains dates across many columns
    best_row = None
    best_count = 0
    max_check_rows = min(15, df.shape[0])
    for r in range(max_check_rows):
        row = df.iloc[r, :]
        cnt = 0
        for v in row:
            if _coerce_header_to_date(v) is not None:
                cnt += 1
        if cnt > best_count:
            best_count = cnt
            best_row = r

    if best_row is None or best_count < 3:
        print(f"[LOAD-WH] Could not detect date header row (best_row={best_row}, count={best_count})")
        return planned

    print(f"[LOAD-WH] Using row {best_row} as date header with {best_count} date-like columns")
    date_headers = df.iloc[best_row, :]
    date_cols: list[date | None] = []
    for v in date_headers:
        date_cols.append(_coerce_header_to_date(v))

    # 2) For each row where classification column == target_label, collect per-day values
    # We observed layout: col0=Project, col1=Prod line, col2=label (e.g. 'Target (LTP input)')
    label_col_idx = 2
    prod_line_col_idx = 1

    for r in range(df.shape[0]):
        try:
            label_raw = df.iat[r, label_col_idx]
        except Exception:
            continue
        if str(label_raw).strip() != target_label:
            continue

        prod_line_raw = df.iat[r, prod_line_col_idx] if prod_line_col_idx < df.shape[1] else None
        code = norm_code(str(prod_line_raw))
        if not code:
            continue

        per_day: dict[date, int] = {}
        row = df.iloc[r, :]
        for j, d in enumerate(date_cols):
            if d is None:
                continue
            if j >= len(row):
                continue
            val = row.iat[j]
            try:
                if pd.isna(val):
                    q = 0
                else:
                    q = int(round(float(val)))
            except Exception:
                q = 0
            if q:
                per_day[d] = per_day.get(d, 0) + q

        planned[code] = per_day

    print(f"[LOAD-WH] Loaded {len(planned)} production lines from WH Receipt workbook")
    return planned


def load_planned_xlsx(path: str):
    """Return dict: {line_code: {date: qty_int}}. Reads Excel file with formulas."""
    planned: dict[str, dict[date, int]] = {}
    if not os.path.exists(path):
        print(f"[LOAD] ERROR: Planned XLSX not found: {path}")
        return planned
    
    print(f"[LOAD] Loading planned data from: {path}")
    
    try:
        # Read Excel file - pandas will evaluate formulas and return calculated values
        df = pd.read_excel(path, engine='openpyxl')
        
        # First column is 'Project' (line codes), rest are dates
        if df.empty or len(df.columns) < 2:
            return planned
        
        # Parse date columns (skip first column which is 'Project')
        date_cols = []
        for col in df.columns[1:]:
            d = _coerce_header_to_date(col)
            date_cols.append(d)
        
        # Process each row
        for idx, row in df.iterrows():
            code = norm_code(str(row.iloc[0]) if pd.notna(row.iloc[0]) else '')
            if not code:
                continue
            
            per_day: dict[date, int] = {}
            for i, cell_value in enumerate(row.iloc[1:]):
                d = date_cols[i]
                if not d:
                    continue
                try:
                    # Handle NaN, None, or empty values
                    if pd.isna(cell_value):
                        q = 0
                    else:
                        q = int(float(cell_value))
                except Exception:
                    q = 0
                per_day[d] = q
            planned[code] = per_day
    except Exception as e:
        print(f"[LOAD] ERROR loading planned XLSX: {e}")
        import traceback
        traceback.print_exc()
        return planned
    
    print(f"[LOAD] Successfully loaded {len(planned)} production lines")
    return planned


def fetch_produced_by_day(start_d: date, end_d: date):
    """Return dict: {line_code: {date: qty_float}} for [start_d, end_d]."""
    sql = (
        "SELECT pt.pt_prod_line AS line, CAST(tr.tr_effdate AS date) AS d, "
        "SUM(CAST(tr.tr_qty_loc AS DECIMAL(18,2))) AS qty "
        "FROM dbo.tr_hist tr "
        "JOIN dbo.pt_mstr pt ON tr.tr_part = pt.pt_part "
        "WHERE tr.tr_type IN ('RCT-WO','rct-wo') "
        "AND tr.tr_effdate >= ? AND tr.tr_effdate < DATEADD(day, 1, ?) "
        "AND tr.tr_qty_loc > 0 "
        "GROUP BY pt.pt_prod_line, CAST(tr.tr_effdate AS date)"
    )
    data: dict[str, dict[date, float]] = {}
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(sql, start_d, end_d)
        for line_code, d, qty in cur.fetchall():
            if not line_code:
                continue
            code = norm_code(line_code)
            data.setdefault(code, {})[d] = float(qty or 0)
        cur.close()
    finally:
        conn.close()
    return data


def daterange(d0: date, d1: date):
    d = d0
    while d <= d1:
        yield d
        d += timedelta(days=1)


def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())  # Monday


def aggregate(values_by_day: dict[date, float], start_d: date, end_d: date) -> float:
    total = 0.0
    for d in daterange(start_d, end_d):
        total += float(values_by_day.get(d, 0) or 0)
    return total


def compute_metrics():
    today = date.today()
    wd = today.weekday()  # Monday=0 ... Sunday=6
    # For Sun/Mon: include both Friday and Saturday. Others: just yesterday.
    is_weekend_view = wd in (6, 0)  # Sunday or Monday
    if is_weekend_view:
        as_of = today - timedelta(days=1 if wd == 6 else 2)  # Saturday
        daily_start = as_of - timedelta(days=1)  # Friday
    else:
        as_of = today - timedelta(days=1)  # Yesterday
        daily_start = as_of  # Same day

    start_month = as_of.replace(day=1)
    start_week = monday_of_week(as_of)

    print(f"[COMPUTE] Computing metrics for date: {as_of}")
    print(f"[COMPUTE] PVS_USE_WH_RECEIPT={PVS_USE_WH_RECEIPT}, PVS_RECALC_XLSX={PVS_RECALC_XLSX}, PVS_PLANNED_XLSX={PVS_PLANNED_XLSX}")

    if PVS_USE_WH_RECEIPT:
        print("[COMPUTE] Using WH Receipt workbook for planned schedule (no Excel COM)...")
        planned = load_planned_from_wh_receipt(
            PVS_EXTERNAL_XLSX,
            PVS_EXTERNAL_SHEET,
            PVS_EXTERNAL_TARGET_LABEL,
        )
    else:
        if PVS_RECALC_XLSX:
            print("[COMPUTE] Attempting Excel recalculation of Planned_qtys.xlsx...")
            recalc_excel_workbook(PVS_PLANNED_XLSX)
        else:
            print("[COMPUTE] Excel recalc disabled - using cached values in Planned_qtys.xlsx")

        planned = load_planned_xlsx(PVS_PLANNED_XLSX)
    produced = fetch_produced_by_day(start_month, as_of)
    mapping = load_map_csv(PVS_MAP_CSV)

    # Build union of all line codes seen in plan or production
    codes = sorted(set(planned.keys()) | set(produced.keys()))

    rows = []
    for code in codes:
        disp = mapping.get(code, code)
        plan_days = planned.get(code, {})
        prod_days = produced.get(code, {})

        # Plans (previous business day or Fri+Sat)
        plan_day = int(aggregate(plan_days, daily_start, as_of))
        plan_wtd = int(aggregate(plan_days, start_week, as_of))
        plan_mtd = int(aggregate(plan_days, start_month, as_of))

        # Produced (previous business day or Fri+Sat)
        prod_day = float(aggregate(prod_days, daily_start, as_of))
        prod_wtd = float(aggregate(prod_days, start_week, as_of))
        prod_mtd = float(aggregate(prod_days, start_month, as_of))

        # Deltas
        d_day = prod_day - plan_day
        d_wtd = prod_wtd - plan_wtd
        d_mtd = prod_mtd - plan_mtd

        def adherence(delta: float, schedule: float) -> float:
            if schedule <= 0:
                return 0.0
            pct = (delta / schedule) * 100.0
            # Clamp extreme outliers instead of zeroing them
            if pct > PVS_ADHERENCE_CLAMP:
                return PVS_ADHERENCE_CLAMP
            if pct < -PVS_ADHERENCE_CLAMP:
                return -PVS_ADHERENCE_CLAMP
            return pct

        # Determine category (SEW, ASSY, or OTHER)
        disp_upper = disp.upper()
        if 'SEW' in disp_upper:
            category = 'SEW'
        elif 'ASSY' in disp_upper:
            category = 'ASSY'
        else:
            category = 'OTHER'

        rows.append({
            'code': code,
            'line': disp,
            'category': category,
            'mtd': {
                'schedule': plan_mtd,
                'production': round(prod_mtd, 2),
                'delta': round(d_mtd, 2),
                # Use cumulative planned as denominator for MTD adherence
                'adherence_pct': round(adherence(d_mtd, plan_mtd), 1),
            },
            'wtd': {
                'schedule': plan_wtd,
                'production': round(prod_wtd, 2),
                'delta': round(d_wtd, 2),
                # Use cumulative planned as denominator for WTD adherence
                'adherence_pct': round(adherence(d_wtd, plan_wtd), 1),
            },
            'daily': {
                'schedule': plan_day,
                'production': round(prod_day, 2),
                'delta': round(d_day, 2),
                'adherence_pct': round(adherence(d_day, plan_day), 1),
            }
        })

    # Sort: SEW first, then ASSY, then OTHER; within each category sort by line name
    category_order = {'SEW': 0, 'ASSY': 1, 'OTHER': 2}
    rows.sort(key=lambda r: (category_order.get(r['category'], 2), r['line']))

    # Compute totals for pie charts (MTD only)
    totals = {
        'sew': {'schedule': 0, 'production': 0},
        'assy': {'schedule': 0, 'production': 0},
        'total': {'schedule': 0, 'production': 0},
    }
    for r in rows:
        cat = r['category']
        totals['total']['schedule'] += r['mtd']['schedule']
        totals['total']['production'] += r['mtd']['production']
        if cat == 'SEW':
            totals['sew']['schedule'] += r['mtd']['schedule']
            totals['sew']['production'] += r['mtd']['production']
        elif cat == 'ASSY':
            totals['assy']['schedule'] += r['mtd']['schedule']
            totals['assy']['production'] += r['mtd']['production']

    return {
        'success': True,
        'date': as_of.strftime('%Y-%m-%d'),
        'rows': rows,
        'totals': totals,
    }


@app.route('/')
def index():
    return render_template('pvs.html', version=str(int(datetime.now().timestamp())))


@app.route('/api/pvs')
def api_pvs():
    try:
        return jsonify(compute_metrics())
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/health')
def health():
    return jsonify({'status': 'healthy', 'ts': datetime.now().isoformat()})


if __name__ == '__main__':
    from waitress import serve
    print('=' * 70)
    print(f'Running PVS app at http://{FLASK_HOST}:{PVS_PORT}')
    print('=' * 70)
    serve(app, host=FLASK_HOST, port=PVS_PORT, threads=4)
