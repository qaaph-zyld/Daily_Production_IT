# 🚀 Quick Start Guide

## First Time Setup (5 minutes)

### Step 1: Run Setup
**Double-click:** `setup_venv.bat`

Wait for completion (installs Python packages with proxy support).

### Step 2: Launch Dashboard
**Double-click:** `start_dashboard.bat`

### Step 3: Open Browser
Navigate to: **http://localhost:5000**

---

## Daily Use

1. **Double-click:** `start_dashboard.bat`
2. **Open browser:** http://localhost:5000
3. **Done!** Dashboard auto-refreshes every 15 minutes

---

## Configuration

Everything is pre-configured:
- ✅ Server: a265m001
- ✅ Database: QADEE2798
- ✅ Username: PowerBI
- ✅ Password: P0werB1
- ✅ Port: 5000
- ✅ Refresh: 15 minutes

**No changes needed!**

---

## Troubleshooting

### "Python is not installed"
Install Python 3.8+ from https://www.python.org/

### "Virtual environment not found"
Run `setup_venv.bat` first

### "Port 5000 already in use"
Edit `.env` file, change `FLASK_PORT=5001`

### "Connection failed"
1. Check if SQL Server is running
2. Verify network connection: `ping a265m001`
3. Test credentials in SQL Server Management Studio

---

## Files Overview

| File | Purpose |
|------|---------|
| `setup_venv.bat` | One-time setup (creates virtual environment) |
| `start_dashboard.bat` | Daily launcher (starts dashboard) |
| `.env` | Database credentials (pre-configured) |
| `dashboard_server.py` | Backend server |
| `templates/dashboard.html` | Frontend with Adient branding |
| `DEPLOYMENT_GUIDE.md` | Complete documentation |

---

## Support

For detailed instructions, see **DEPLOYMENT_GUIDE.md**

---

**That's it! You're ready to go! 🎉**
