"""
Production Dashboard Server
Serves hourly production data with auto-refresh capability
Configured for shared company drive deployment
"""

import pyodbc
import json
from flask import Flask, render_template, jsonify
from flask_cors import CORS
from datetime import datetime
import os
from dotenv import load_dotenv
from decimal import Decimal

# Load environment variables from .env file
load_dotenv()

app = Flask(__name__)
CORS(app)

# Database connection configuration from environment variables
DB_CONFIG = {
    'server': os.getenv('DB_SERVER', 'a265m001'),
    'database': os.getenv('DB_DATABASE', 'QADEE2798'),
    'username': os.getenv('DB_USERNAME', 'PowerBI'),
    'password': os.getenv('DB_PASSWORD', 'P0werB1')
}

# Flask configuration
FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))
FLASK_HOST = os.getenv('FLASK_HOST', '0.0.0.0')
FLASK_DEBUG = os.getenv('FLASK_DEBUG', 'False').lower() == 'true'

def _choose_sql_driver():
    """Select the best available ODBC driver on this machine."""
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
    # Fallback
    return 'SQL Server'

def get_db_connection():
    """Create database connection with SQL Server authentication"""
    driver = _choose_sql_driver()
    # For newer drivers (18), encryption is on by default which can fail without certs
    # Set Encrypt=no and TrustServerCertificate to ease intranet connections
    extra = 'TrustServerCertificate=yes;'
    if 'ODBC Driver 18' in driver:
        extra = 'Encrypt=no;TrustServerCertificate=yes;'
    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['username']};"
        f"PWD={DB_CONFIG['password']};"
        f"{extra}"
    )
    return pyodbc.connect(conn_str)

def fetch_production_data():
    """Fetch production data from database"""
    query = """
    SELECT
        Project,
        ISNULL([6-7], 0) AS [6-7],
        ISNULL([7-8], 0) AS [7-8],
        ISNULL([8-9], 0) AS [8-9],
        ISNULL([9-10], 0) AS [9-10],
        ISNULL([10-11], 0) AS [10-11],
        ISNULL([11-12], 0) AS [11-12],
        ISNULL([12-13], 0) AS [12-13],
        ISNULL([13-14], 0) AS [13-14],
        ISNULL([14-22], 0) AS [14-22],
        ISNULL([II Shift], 0) AS [II Shift]
    FROM
    (
        SELECT
            th.[tr_qty_loc],
            CASE 
                WHEN pt.[pt_prod_line] = 'H_FG' THEN 'BJA'
                WHEN pt.[pt_prod_line] = 'B_FG' THEN 'BR223 - SEW'
                WHEN pt.[pt_prod_line] = 'C_FG' THEN 'CDPO - ASSY'
                WHEN pt.[pt_prod_line] = 'Z_FG' THEN 'CDPO - SEW'
                WHEN pt.[pt_prod_line] = '0000' THEN 'Pre-production'
                WHEN pt.[pt_prod_line] = 'F_FG' THEN 'FIAT - SEW'
                WHEN pt.[pt_prod_line] = 'K_FG' THEN 'KIA - ASSY'
                WHEN pt.[pt_prod_line] = 'Q_FG' THEN 'KIA - SEW'
                WHEN pt.[pt_prod_line] = 'U_FG' THEN 'MAN'
                WHEN pt.[pt_prod_line] = 'M_FG' THEN 'MMA - ASSY'
                WHEN pt.[pt_prod_line] = 'N_FG' THEN 'MMA - SEW'
                WHEN pt.[pt_prod_line] = 'O_FG' THEN 'OV5X - ASSY'
                WHEN pt.[pt_prod_line] = 'S_FG' THEN 'OV5X - SEW'
                WHEN pt.[pt_prod_line] = 'P_FG' THEN 'PO426 - SEW'
                WHEN pt.[pt_prod_line] = 'G_FG' THEN 'PZ1D'
                WHEN pt.[pt_prod_line] = 'R_FG' THEN 'Renault'
                WHEN pt.[pt_prod_line] = 'E_FG' THEN 'SCANIA'
                WHEN pt.[pt_prod_line] = 'A_FG' THEN 'VOLVO- SEW'
                WHEN pt.[pt_prod_line] = 'V_FG' THEN 'VOLVO- ASSY'
                WHEN pt.[pt_prod_line] = 'T_FG' THEN 'P13A'
                ELSE 'Other' 
            END AS Project,
            CASE
                WHEN (th.[tr_time] / 3600) >= 0 AND (th.[tr_time] / 3600) < 6 THEN '6-7'
                WHEN (th.[tr_time] / 3600) >= 6 AND (th.[tr_time] / 3600) < 7 THEN '6-7'
                WHEN (th.[tr_time] / 3600) >= 7 AND (th.[tr_time] / 3600) < 8 THEN '7-8'
                WHEN (th.[tr_time] / 3600) >= 8 AND (th.[tr_time] / 3600) < 9 THEN '8-9'
                WHEN (th.[tr_time] / 3600) >= 9 AND (th.[tr_time] / 3600) < 10 THEN '9-10'
                WHEN (th.[tr_time] / 3600) >= 10 AND (th.[tr_time] / 3600) < 11 THEN '10-11'
                WHEN (th.[tr_time] / 3600) >= 11 AND (th.[tr_time] / 3600) < 12 THEN '11-12'
                WHEN (th.[tr_time] / 3600) >= 12 AND (th.[tr_time] / 3600) < 13 THEN '12-13'
                WHEN (th.[tr_time] / 3600) >= 13 AND (th.[tr_time] / 3600) < 14 THEN '13-14'
                WHEN (th.[tr_time] / 3600) >= 14 AND (th.[tr_time] / 3600) < 22 THEN '14-22'
                WHEN (th.[tr_time] / 3600) >= 22 AND (th.[tr_time] / 3600) < 24 THEN 'II Shift'
            END AS TimeBucket
        FROM
            [QADEE2798].[dbo].[tr_hist] AS th
        INNER JOIN
            [QADEE2798].[dbo].[pt_mstr] AS pt ON th.[tr_part] = pt.[pt_part]
        WHERE
            th.[tr_type] = 'RCT-WO'
            AND CAST(th.[tr_effdate] AS DATE) = CAST(GETDATE() AS DATE)
    ) AS SourceData
    PIVOT
    (
        SUM(tr_qty_loc)
        FOR TimeBucket IN ([6-7], [7-8], [8-9], [9-10], [10-11], [11-12], [12-13], [13-14], [14-22], [II Shift])
    ) AS PivotTable
    ORDER BY
        Project;
    """
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query)
        
        columns = [column[0] for column in cursor.description]
        rows = cursor.fetchall()
        
        data = []
        for row in rows:
            row_dict = {}
            for i, col in enumerate(columns):
                value = row[i]
                # Convert to float if numeric/decimal strings
                if isinstance(value, (int, float, Decimal)):
                    row_dict[col] = float(value)
                else:
                    row_dict[col] = value
            data.append(row_dict)
        
        cursor.close()
        conn.close()
        
        return {
            'success': True,
            'data': data,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'columns': columns
        }
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

@app.route('/')
def index():
    """Serve the dashboard HTML"""
    return render_template('dashboard.html')

@app.route('/api/production-data')
def get_production_data():
    """API endpoint to fetch production data"""
    return jsonify(fetch_production_data())

@app.route('/api/health')
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    })

if __name__ == '__main__':
    # Create templates directory if it doesn't exist
    os.makedirs('templates', exist_ok=True)
    
    print("=" * 70)
    print("🏭 Adient Production Dashboard Server Starting...")
    print("=" * 70)
    print(f"📊 Dashboard URL: http://localhost:{FLASK_PORT}")
    print(f"🔌 API Endpoint: http://localhost:{FLASK_PORT}/api/production-data")
    print(f"💾 Database: {DB_CONFIG['server']}/{DB_CONFIG['database']}")
    print(f"👤 User: {DB_CONFIG['username']}")
    print("=" * 70)
    print("✅ Server is ready! Open the dashboard URL in your browser.")
    print("🔄 Dashboard will auto-refresh every 15 minutes.")
    print("=" * 70)
    
    app.run(debug=FLASK_DEBUG, host=FLASK_HOST, port=FLASK_PORT)
