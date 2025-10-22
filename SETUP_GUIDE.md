# Quick Setup Guide - Production Dashboard

## Step 1: Install Dependencies

Open Command Prompt or PowerShell in this directory and run:

```bash
pip install -r requirements.txt
```

## Step 2: Configure Database Connection

Open `dashboard_server.py` in a text editor and find this section (around line 19):

```python
DB_CONFIG = {
    'server': 'your_server_name',  # ⚠️ UPDATE THIS
    'database': 'QADEE2798',
    'trusted_connection': 'yes'
}
```

**Replace `'your_server_name'` with your actual SQL Server name.**

### How to find your SQL Server name:

**Option 1: SQL Server Management Studio (SSMS)**
- Open SSMS
- The server name is shown in the connection dialog
- Example: `localhost\SQLEXPRESS` or `SERVER01`

**Option 2: Command Line**
```bash
sqlcmd -L
```

**Option 3: Check existing connection**
- If you have other scripts connecting to the database, check their connection strings

### Common SQL Server names:
- `localhost\SQLEXPRESS` - Local SQL Server Express
- `.\SQLEXPRESS` - Same as above (shorthand)
- `SERVER01` - Named server
- `192.168.1.100` - IP address
- `server.domain.com` - Fully qualified domain name

## Step 3: Start the Dashboard

**Option A: Double-click the batch file**
```
start_dashboard.bat
```

**Option B: Run from command line**
```bash
python dashboard_server.py
```

## Step 4: Open Dashboard in Browser

Once the server starts, open your web browser and go to:

```
http://localhost:5000
```

## Verification Checklist

✅ Python 3.7+ installed
✅ Dependencies installed (`pip install -r requirements.txt`)
✅ SQL Server name configured in `dashboard_server.py`
✅ Database `QADEE2798` is accessible
✅ Server starts without errors
✅ Dashboard loads in browser

## Troubleshooting

### Error: "No module named 'flask'"
**Solution:** Install dependencies
```bash
pip install -r requirements.txt
```

### Error: "Connection failed" or database errors
**Solutions:**
1. Verify SQL Server is running
2. Check server name in `DB_CONFIG`
3. Test connection with SSMS first
4. Ensure you have read permissions on the database

### Error: "Port 5000 already in use"
**Solution:** Change port in `dashboard_server.py` (line ~180):
```python
app.run(debug=True, host='0.0.0.0', port=5001)
```

### Dashboard shows "No data available"
**Solutions:**
1. Check if there's production data for today
2. Run the SQL query directly in SSMS to verify results
3. Check browser console (F12) for errors

## Features Overview

Once running, the dashboard provides:

- **Real-time Statistics** - Total production, active projects, peak hours
- **Interactive Heatmap** - Color-coded production by project and time
- **Bar Chart** - Total production comparison by project
- **Line Chart** - Production trends throughout the day
- **Detailed Table** - Complete data with totals
- **Auto-Refresh** - Updates every 15 minutes automatically

## Next Steps

1. Bookmark `http://localhost:5000` in your browser
2. Consider setting up the server to start automatically (Windows Task Scheduler)
3. For remote access, configure firewall and use server's IP address

## Need Help?

Check the full README.md for detailed documentation and customization options.
