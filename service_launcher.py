import os
import sys
from waitress import serve

# Determine paths before importing the Flask app so .env resolves correctly
IS_FROZEN = getattr(sys, 'frozen', False)
EXE_DIR = os.path.dirname(sys.executable) if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
ASSET_BASE = getattr(sys, '_MEIPASS', EXE_DIR) if IS_FROZEN else EXE_DIR

# Use the executable directory as CWD so dashboard_server.load_dotenv finds .env
os.chdir(EXE_DIR)

# Import app after CWD is set
import dashboard_server as ds  # noqa: E402

# Point Flask to bundled assets (templates/static in _MEIPASS when frozen)
templates_path = os.path.join(ASSET_BASE, 'templates')
static_path = os.path.join(ASSET_BASE, 'static')

try:
    ds.app.template_folder = templates_path
    ds.app.static_folder = static_path
    if ds.app.jinja_loader:
        ds.app.jinja_loader.searchpath = [templates_path]
except Exception:
    pass

HOST = os.getenv('FLASK_HOST', '0.0.0.0')
PORT = int(os.getenv('FLASK_PORT', '5000'))

if __name__ == '__main__':
    print('=' * 70)
    print(f'Running Adient Dashboard via Waitress at http://{HOST}:{PORT}')
    print('=' * 70)
    serve(ds.app, host=HOST, port=PORT, threads=4)
