import os
import csv
import pyodbc
from flask import Flask, render_template, jsonify
from flask_cors import CORS
from datetime import datetime, date, timedelta
from dotenv import load_dotenv
from decimal import Decimal

# Load .env
load_dotenv()

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
PVS_PLANNED_CSV = os.getenv('PVS_PLANNED_CSV', os.path.join('PVS', 'Planned_qtys.csv'))
PVS_ADHERENCE_CLAMP = float(os.getenv('PVS_ADHERENCE_CLAMP', '300'))  # percent cap; values beyond are treated as 0%
PVS_MAP_CSV = os.getenv('PVS_MAP_CSV', os.path.join('PVS', 'ProdLine_Project_Map.csv'))


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


def load_planned_csv(path: str):
    """Return dict: {line_code: {date: qty_int}}"""
    planned: dict[str, dict[date, int]] = {}
    if not os.path.exists(path):
        return planned
    with open(path, 'r', newline='', encoding='utf-8') as f:
        rdr = csv.reader(f)
        rows = list(rdr)
    if not rows:
        return planned
    header = rows[0]
    date_cols = [_parse_date_ddmmyyyy(c) for c in header[1:]]
    for row in rows[1:]:
        if not row:
            continue
        code = norm_code(row[0] if len(row) else '')
        if not code:
            continue
        per_day: dict[date, int] = {}
        for i, cell in enumerate(row[1:]):
            d = date_cols[i]
            if not d:
                continue
            try:
                q = int(float(cell.strip() or '0'))
            except Exception:
                q = 0
            per_day[d] = q
        planned[code] = per_day
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
    if wd in (6, 0):  # Sunday or Monday
        as_of = today - timedelta(days=1 if wd == 6 else 2)  # Saturday
    else:
        as_of = today - timedelta(days=1)  # Yesterday

    start_month = as_of.replace(day=1)
    start_week = monday_of_week(as_of)

    planned = load_planned_csv(PVS_PLANNED_CSV)
    produced = fetch_produced_by_day(start_month, as_of)
    mapping = load_map_csv(PVS_MAP_CSV)

    # Build union of all line codes seen in plan or production
    codes = sorted(set(planned.keys()) | set(produced.keys()))

    rows = []
    for code in codes:
        disp = mapping.get(code, code)
        plan_days = planned.get(code, {})
        prod_days = produced.get(code, {})

        # Plans (previous business day)
        plan_day = int(plan_days.get(as_of, 0) or 0)
        plan_wtd = int(aggregate(plan_days, start_week, as_of))
        plan_mtd = int(aggregate(plan_days, start_month, as_of))

        # Produced (previous business day)
        prod_day = float(prod_days.get(as_of, 0) or 0)
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

        rows.append({
            'code': code,
            'line': disp,
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

    # Sort by display line
    rows.sort(key=lambda r: r['line'])

    return {
        'success': True,
        'date': as_of.strftime('%Y-%m-%d'),
        'rows': rows,
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
