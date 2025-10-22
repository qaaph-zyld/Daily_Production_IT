import sys
print(f"Python: {sys.version}")
print(f"Executable: {sys.executable}")
print("\nTrying to import packages...")

try:
    import flask
    print(f"✓ Flask {flask.__version__}")
except ImportError as e:
    print(f"✗ Flask not installed: {e}")

try:
    import flask_cors
    print(f"✓ Flask-CORS installed")
except ImportError as e:
    print(f"✗ Flask-CORS not installed: {e}")

try:
    import pyodbc
    print(f"✓ pyodbc {pyodbc.version}")
except ImportError as e:
    print(f"✗ pyodbc not installed: {e}")

try:
    import dotenv
    print(f"✓ python-dotenv installed")
except ImportError as e:
    print(f"✗ python-dotenv not installed: {e}")

print("\nAll checks complete!")
