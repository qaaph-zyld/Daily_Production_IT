# 🏭 Adient Production Dashboard

Real-time hourly production monitoring dashboard with auto-refresh, designed for shared company drive deployment.

![Adient Colors](https://img.shields.io/badge/Teal-%23004851-004851?style=flat-square) ![Adient Colors](https://img.shields.io/badge/Lime-%23C4D600-C4D600?style=flat-square)

---

## ✨ Features

- ✅ **Real-time Production Data** - Displays today's hourly production by project
- ✅ **Auto-Refresh** - Updates every 15 minutes automatically
- ✅ **Interactive Visualizations**:
  - Color-coded heatmap (project × time slot)
  - Bar chart (total production by project)
  - Line chart (production trends by hour)
  - Detailed data table with totals
- ✅ **Adient Branding** - Corporate colors (teal #004851 and lime green #C4D600)
- ✅ **Portable Virtual Environment** - Works on any computer
- ✅ **Shared Drive Compatible** - Multiple users can access simultaneously
- ✅ **Proxy Support** - Bypasses corporate firewall (104.129.196.38:10563)
- ✅ **Secure Authentication** - SQL Server credentials in .env file
- ✅ **Responsive Design** - Works on desktop, tablet, and mobile

---

## 🚀 Quick Start

### First Time Setup (5 minutes)

1. **Run setup script:**
   ```cmd
   setup_venv.bat
   ```
   This creates a virtual environment and installs all dependencies with proxy support.

2. **Launch dashboard:**
   ```cmd
   start_dashboard.bat
   ```

3. **Open browser:**
   ```
   http://localhost:5000
   ```

**That's it!** The dashboard is pre-configured and ready to use.

---

## 📋 Requirements

- **Python:** 3.8 or higher
- **Operating System:** Windows
- **Network Access:** SQL Server a265m001
- **Database:** QADEE2798
- **Credentials:** PowerBI / P0werB1 (configured in .env)

---

## 📁 Project Structure

```
Daily_Production_IT/
├── venv/                      ← Virtual environment (auto-created)
├── templates/
│   └── dashboard.html         ← Frontend with Adient branding
├── .env                       ← Database credentials (CONFIGURED)
├── .env.example               ← Template for new deployments
├── .gitignore                 ← Protects sensitive files
├── dashboard_server.py        ← Backend Flask server
├── requirements.txt           ← Python dependencies
├── setup_venv.bat            ← One-time setup script
├── start_dashboard.bat       ← Daily launcher script
├── DEPLOYMENT_GUIDE.md       ← Complete deployment documentation
├── QUICK_START.md            ← Quick reference guide
└── README.md                 ← This file
```

---

## 🔧 Configuration

### Pre-Configured Settings

The dashboard is **already configured** with:

```ini
# Database
DB_SERVER=a265m001
DB_DATABASE=QADEE2798
DB_USERNAME=PowerBI
DB_PASSWORD=P0werB1
- Hover for exact values

### 4. Production by Time Slot (Line Chart)
- Shows production trends throughout the day
- Identifies peak production hours
- Smooth curve for better visualization

### 5. Detailed Production Table
- Complete data view with all values
- Sortable columns
- Row and column totals

## Time Slots

The dashboard tracks production across these time slots:
- **6-7** - Morning shift start
- **7-8** through **13-14** - Regular hourly slots
- **14-22** - Extended afternoon/evening period
- **II Shift** - Second shift (22:00-24:00)

## Projects Tracked

- BJA
- BR223 - SEW
- CDPO - ASSY
- CDPO - SEW
- FIAT - SEW
- KIA - ASSY
- KIA - SEW
- MAN
- MMA - ASSY
- MMA - SEW
- OV5X - ASSY
- OV5X - SEW
- PO426 - SEW
- PZ1D
- SCANIA
- VOLVO - ASSY
- VOLVO - SEW

## Troubleshooting

### Database Connection Issues

If you see "Connection failed" or database errors:

1. Verify SQL Server is running
2. Check server name in `DB_CONFIG`
3. Ensure you have permissions to access the database
4. Test connection using SQL Server Management Studio first

### Port Already in Use

If port 5000 is already in use, change it in `dashboard_server.py`:

```python
app.run(debug=True, host='0.0.0.0', port=5001)  # Change to any available port
```

### No Data Displayed

If the dashboard shows "No data available":

1. Check if there's production data for today in the database
2. Verify the SQL query is returning results
3. Check browser console (F12) for JavaScript errors

## API Endpoints

The server provides these API endpoints:

- `GET /` - Dashboard HTML page
- `GET /api/production-data` - JSON data endpoint
- `GET /api/health` - Health check endpoint

## Customization

### Change Refresh Interval

Edit `dashboard.html` and modify the `REFRESH_INTERVAL` constant:

```javascript
const REFRESH_INTERVAL = 10 * 60 * 1000; // 10 minutes
```

### Modify Colors

Update the color palette in `dashboard.html`:

```javascript
const colors = [
    '#FF6384', '#36A2EB', '#FFCE56', // Add your colors here
];
```

### Adjust Chart Heights

Modify the `.chart-container` height in the CSS:

```css
.chart-container {
    height: 800px; /* Increase for taller charts */
}
```

## Technical Stack

- **Backend:** Python Flask
- **Database:** SQL Server (via pyodbc)
- **Frontend:** HTML5, CSS3, JavaScript
- **Charts:** Chart.js 4.4.0
- **Auto-refresh:** JavaScript setInterval

## License

Internal use only - Adient IT Department

## Support

For issues or questions, contact the IT Development team.
