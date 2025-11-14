import os
import sys
from waitress import serve

# Resolve paths before importing to ensure .env and CSV resolve next to the EXE
IS_FROZEN = getattr(sys, 'frozen', False)
EXE_DIR = os.path.dirname(sys.executable) if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
ASSET_BASE = getattr(sys, '_MEIPASS', EXE_DIR) if IS_FROZEN else EXE_DIR

os.chdir(EXE_DIR)

import pvs_server as ps  # noqa: E402

# Point Flask to bundled assets
try:
    ps.app.template_folder = os.path.join(ASSET_BASE, 'templates')
    ps.app.static_folder = os.path.join(ASSET_BASE, 'static')
    if ps.app.jinja_loader:
        ps.app.jinja_loader.searchpath = [ps.app.template_folder]
except Exception:
    pass

HOST = os.getenv('FLASK_HOST', '0.0.0.0')
PORT = int(os.getenv('PVS_PORT', '5051'))

if __name__ == '__main__':
    print('=' * 70)
    print(f'Running Adient PVS via Waitress at http://{HOST}:{PORT}')
    print('=' * 70)
    serve(ps.app, host=HOST, port=PORT, threads=4)
