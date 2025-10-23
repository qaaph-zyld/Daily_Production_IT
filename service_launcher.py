import os
import sys
from waitress import serve

# Ensure we can import the app
import dashboard_server as ds

# If running as a PyInstaller EXE, templates and static will be in sys._MEIPASS
if getattr(sys, 'frozen', False):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Make sure all relative paths (like .env) resolve next to the executable
os.chdir(base_path)

# Point Flask to bundled assets if needed
templates_path = os.path.join(base_path, 'templates')
static_path = os.path.join(base_path, 'static')

# Update app folders and Jinja loader search path
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
