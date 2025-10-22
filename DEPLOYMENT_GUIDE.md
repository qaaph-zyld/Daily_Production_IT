# 🏭 Adient Production Dashboard - Deployment Guide

## 📋 Overview

This dashboard is designed to run on a **shared company drive** and can be accessed by multiple users simultaneously. It displays real-time hourly production data with auto-refresh every 15 minutes.

---

## 🎯 Key Features

- ✅ **Portable Virtual Environment** - Works independently on any computer
- ✅ **Shared Drive Compatible** - Multiple users can access simultaneously
- ✅ **Auto-Refresh** - Updates every 15 minutes automatically
- ✅ **Secure Authentication** - SQL Server authentication with .env file
- ✅ **Adient Branding** - Corporate colors (teal #004851 and lime green #C4D600)
- ✅ **Proxy Support** - Bypasses corporate firewall restrictions

---

## 📁 Project Structure

```
Daily_Production_IT/
├── venv/                      ← Virtual environment (created by setup)
├── templates/
│   └── dashboard.html         ← Frontend with Adient branding
├── .env                       ← Database credentials (CONFIGURED)
├── .env.example               ← Template for new deployments
├── dashboard_server.py        ← Backend Flask server
├── requirements.txt           ← Python dependencies
├── setup_venv.bat            ← One-time setup script
├── start_dashboard.bat       ← Daily launcher script
└── DEPLOYMENT_GUIDE.md       ← This file
```

---

## 🚀 Quick Start (First Time Setup)

### Step 1: Run Setup Script

**Double-click:** `setup_venv.bat`

This will:
1. Check Python installation
2. Create virtual environment in `venv/` folder
3. Install all dependencies with proxy support
4. Display installed packages

**Expected output:**
```
============================================================================
  Adient Production Dashboard - Virtual Environment Setup
============================================================================

[1/5] Checking Python installation...
Python 3.x.x

[2/5] Creating virtual environment...
Virtual environment created successfully!

[3/5] Activating virtual environment...
Virtual environment activated!

[4/5] Upgrading pip...

[5/5] Installing dependencies from requirements.txt...
Using proxy: 104.129.196.38:10563

============================================================================
  Setup Complete!
============================================================================
```

### Step 2: Verify Configuration

The `.env` file is already configured with:
- **Server:** a265m001
- **Database:** QADEE2798
- **Username:** PowerBI
- **Password:** P0werB1

**No changes needed!**

### Step 3: Launch Dashboard

**Double-click:** `start_dashboard.bat`

This will:
1. Activate virtual environment
2. Check configuration
3. Start Flask server

**Expected output:**
```
============================================================================
  Adient Production Dashboard - Starting Server
============================================================================

[1/3] Activating virtual environment...
Virtual environment activated!

[2/3] Checking .env configuration...

[3/3] Starting dashboard server...

======================================================================
🏭 Adient Production Dashboard Server Starting...
======================================================================
📊 Dashboard URL: http://localhost:5000
🔌 API Endpoint: http://localhost:5000/api/production-data
💾 Database: a265m001/QADEE2798
👤 User: PowerBI
======================================================================
✅ Server is ready! Open the dashboard URL in your browser.
🔄 Dashboard will auto-refresh every 15 minutes.
======================================================================
```

### Step 4: Access Dashboard

Open your web browser and navigate to:

**http://localhost:5000**

---

## 🌐 Multi-User Access

### For Users on the Same Network

If other users want to access the dashboard from their computers:

1. Find the server computer's IP address:
   ```cmd
   ipconfig
   ```
   Look for "IPv4 Address" (e.g., 192.168.1.100)

2. Share this URL with other users:
   ```
   http://[SERVER_IP]:5000
   ```
   Example: `http://192.168.1.100:5000`

3. Ensure Windows Firewall allows port 5000:
   ```cmd
   netsh advfirewall firewall add rule name="Adient Dashboard" dir=in action=allow protocol=TCP localport=5000
   ```

### For Shared Drive Access

Multiple users can run the dashboard from the shared drive:

1. Each user double-clicks `start_dashboard.bat`
2. Each user accesses their own instance at `http://localhost:5000`
3. All instances connect to the same database

**Note:** Only one user per computer can run the dashboard at a time (port 5000 limitation).

---

## 🔧 Configuration

### Database Configuration (.env file)

```ini
# Database Configuration
DB_SERVER=a265m001
DB_DATABASE=QADEE2798
DB_USERNAME=PowerBI
DB_PASSWORD=P0werB1

# Server Configuration
FLASK_PORT=5000
FLASK_HOST=0.0.0.0
FLASK_DEBUG=False

# Auto-refresh interval (in seconds)
REFRESH_INTERVAL=900
```

### Changing the Port

If port 5000 is already in use, edit `.env`:

```ini
FLASK_PORT=5001
```

Then restart the dashboard.

### Changing Refresh Interval

To change auto-refresh from 15 minutes to another interval:

1. Edit `.env` file:
   ```ini
   REFRESH_INTERVAL=600  # 10 minutes (in seconds)
   ```

2. Edit `templates/dashboard.html` line 280:
   ```javascript
   const REFRESH_INTERVAL = 10 * 60 * 1000; // 10 minutes in milliseconds
   ```

---

## 🎨 Adient Branding

The dashboard uses official Adient colors:

| Element | Color | Hex Code |
|---------|-------|----------|
| Primary Teal | Dark Teal | #004851 |
| Secondary Teal | Medium Teal | #006B75 |
| Accent | Lime Green | #C4D600 |
| Secondary Accent | Light Lime | #A8C800 |

### Color Usage:
- **Background:** Teal gradient (#004851 → #002b30)
- **Headers:** Lime green (#C4D600)
- **Charts:** Mix of teal and lime variations
- **Heatmap:** Lime green → Teal gradient

---

## 📊 Dashboard Components

### 1. Statistics Cards
- Total Production
- Active Projects
- Peak Hour
- Peak Production Value

### 2. Interactive Heatmap
- Color-coded production by project and time slot
- Hover for details
- Gradient from lime green (low) to teal (high)

### 3. Bar Chart
- Total production by project
- Adient brand colors

### 4. Line Chart
- Production trends throughout the day
- Shows all time slots

### 5. Detailed Table
- Complete data matrix
- Project × Time slot breakdown
- Row and column totals

---

## 🔄 Auto-Refresh System

The dashboard automatically refreshes every **15 minutes** (900 seconds).

**Features:**
- Live countdown timer: "Next refresh in: MM:SS"
- Status indicator: Green pulse when connected
- Seamless updates without page reload
- Manual refresh button available

**How it works:**
1. JavaScript timer counts down from 15:00
2. At 00:00, fetches new data from `/api/production-data`
3. Updates all charts and tables
4. Resets timer to 15:00

---

## 🐛 Troubleshooting

### Issue: "Python is not installed or not in PATH"

**Solution:**
1. Install Python 3.8+ from https://www.python.org/
2. During installation, check "Add Python to PATH"
3. Restart command prompt
4. Run `setup_venv.bat` again

### Issue: "Failed to install dependencies"

**Solution 1 - With Proxy:**
```cmd
venv\Scripts\activate
pip install -r requirements.txt --proxy 104.129.196.38:10563
```

**Solution 2 - Without Proxy:**
```cmd
venv\Scripts\activate
pip install -r requirements.txt
```

**Solution 3 - Manual Installation:**
```cmd
venv\Scripts\activate
pip install Flask==3.0.0 --proxy 104.129.196.38:10563
pip install Flask-CORS==4.0.0 --proxy 104.129.196.38:10563
pip install pyodbc==5.0.1 --proxy 104.129.196.38:10563
pip install python-dotenv==1.0.0 --proxy 104.129.196.38:10563
```

### Issue: "Connection failed" or "Database error"

**Solution:**
1. Verify SQL Server is running:
   - Open SQL Server Management Studio (SSMS)
   - Connect to `a265m001`

2. Test credentials:
   - Server: a265m001
   - Database: QADEE2798
   - Username: PowerBI
   - Password: P0werB1

3. Check network connectivity:
   ```cmd
   ping a265m001
   ```

4. Verify database permissions:
   - User `PowerBI` needs SELECT permission on:
     - `[QADEE2798].[dbo].[tr_hist]`
     - `[QADEE2798].[dbo].[pt_mstr]`

### Issue: "Port 5000 already in use"

**Solution 1 - Change Port:**
Edit `.env` file:
```ini
FLASK_PORT=5001
```

**Solution 2 - Kill Existing Process:**
```cmd
netstat -ano | findstr :5000
taskkill /PID [PID_NUMBER] /F
```

### Issue: "No data available"

**Possible causes:**
1. No production data for today
2. Database connection issue
3. Query returned empty result

**Solution:**
1. Run the SQL query directly in SSMS to verify data exists
2. Check browser console (F12) for JavaScript errors
3. Check Flask server console for Python errors

### Issue: "Virtual environment not found"

**Solution:**
Run `setup_venv.bat` to create the virtual environment.

---

## 🔒 Security Considerations

### Credentials Storage

- ✅ Credentials stored in `.env` file (not in code)
- ✅ `.env` file should be excluded from version control
- ✅ Use `.env.example` as template for new deployments

### Network Security

- Dashboard runs on `0.0.0.0` (all network interfaces)
- Accessible from any computer on the network
- Consider firewall rules for production deployment

### Recommendations

1. **Restrict .env file access:**
   ```cmd
   icacls .env /inheritance:r
   icacls .env /grant:r "%USERNAME%:F"
   ```

2. **Use read-only database account:**
   - Current user `PowerBI` should only have SELECT permissions
   - No INSERT, UPDATE, DELETE permissions needed

3. **Enable HTTPS (optional):**
   - For production, consider using HTTPS
   - Requires SSL certificate configuration

---

## 📈 Performance Optimization

### Database Query

The dashboard runs this query every 15 minutes:

```sql
SELECT Project, [6-7], [7-8], [8-9], [9-10], [10-11], [11-12], 
       [12-13], [13-14], [14-22], [II Shift]
FROM (
    SELECT th.[tr_qty_loc], 
           CASE WHEN pt.[pt_prod_line] = 'H_FG' THEN 'BJA' ... END AS Project,
           CASE WHEN (th.[tr_time] / 3600) >= 6 AND (th.[tr_time] / 3600) < 7 THEN '6-7' ... END AS TimeBucket
    FROM [QADEE2798].[dbo].[tr_hist] AS th
    INNER JOIN [QADEE2798].[dbo].[pt_mstr] AS pt ON th.[tr_part] = pt.[pt_part]
    WHERE th.[tr_type] = 'RCT-WO'
      AND CAST(th.[tr_effdate] AS DATE) = CAST(GETDATE() AS DATE)
) AS SourceData
PIVOT (SUM(tr_qty_loc) FOR TimeBucket IN ([6-7], [7-8], ...)) AS PivotTable
ORDER BY Project;
```

**Performance tips:**
- Query filters by today's date only
- Uses indexed columns (`tr_type`, `tr_effdate`)
- Returns ~20 rows (one per project)
- Execution time: < 1 second

### Browser Performance

- Charts use Chart.js (hardware-accelerated)
- Auto-refresh uses AJAX (no page reload)
- Minimal memory footprint
- Works on older browsers (IE11+)

---

## 🔄 Maintenance

### Daily Operations

1. **Start dashboard:** Double-click `start_dashboard.bat`
2. **Stop dashboard:** Close the command prompt window or press Ctrl+C
3. **Restart dashboard:** Close and run `start_dashboard.bat` again

### Weekly Maintenance

1. **Check for updates:**
   ```cmd
   venv\Scripts\activate
   pip list --outdated
   ```

2. **Update dependencies (if needed):**
   ```cmd
   pip install --upgrade Flask Flask-CORS pyodbc python-dotenv --proxy 104.129.196.38:10563
   ```

### Monthly Maintenance

1. **Review logs:** Check Flask console for errors
2. **Test database connection:** Verify credentials still work
3. **Backup configuration:** Copy `.env` file to safe location

---

## 📞 Support

### Common Questions

**Q: Can multiple users access the dashboard simultaneously?**
A: Yes! Each user can run their own instance from the shared drive, or one user can host it and share the URL.

**Q: Does the dashboard work offline?**
A: No, it requires network access to the SQL Server (a265m001).

**Q: Can I customize the colors?**
A: Yes, edit `templates/dashboard.html` and search for color codes (#004851, #C4D600, etc.).

**Q: Can I change the refresh interval?**
A: Yes, edit `.env` file and `templates/dashboard.html` (see Configuration section).

**Q: How do I add more projects?**
A: Projects are automatically pulled from the database based on `pt_prod_line` values.

### Contact Information

For technical support or questions:
- **IT Department:** [Your IT contact]
- **Database Admin:** [Your DBA contact]
- **Dashboard Developer:** [Your contact]

---

## 📝 Version History

### Version 1.0 (Current)
- ✅ Virtual environment support
- ✅ Shared drive compatibility
- ✅ SQL Server authentication
- ✅ Adient branding
- ✅ Auto-refresh (15 min)
- ✅ Proxy support
- ✅ Multi-user access
- ✅ Interactive charts
- ✅ Heatmap visualization

---

## 🎉 Success Checklist

Before going live, verify:

- [ ] Python 3.8+ installed
- [ ] Virtual environment created (`venv/` folder exists)
- [ ] Dependencies installed (run `pip list` to verify)
- [ ] `.env` file configured with correct credentials
- [ ] Database connection tested (server starts without errors)
- [ ] Dashboard accessible at http://localhost:5000
- [ ] Data displays correctly (no "No data available" message)
- [ ] Auto-refresh works (countdown timer visible)
- [ ] Charts render properly (heatmap, bar chart, line chart)
- [ ] Adient colors applied (teal and lime green theme)
- [ ] Multiple users can access (if shared hosting)

---

## 🚀 You're Ready!

The dashboard is now fully configured and ready for production use on your shared company drive. Enjoy real-time production monitoring with Adient branding!

**Quick Start Commands:**
1. First time: `setup_venv.bat`
2. Daily use: `start_dashboard.bat`
3. Access: http://localhost:5000

---

*Last updated: October 2025*
*Adient Production Dashboard v1.0*
