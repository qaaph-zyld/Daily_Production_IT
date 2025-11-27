# Adient Loznica PVS Dashboard

Real-time Production vs Schedule (PVS) dashboard displaying MTD, WTD, and Daily adherence metrics on a TV display.

---

## Quick Start

### 1. Install the Windows Service (one-time)
```powershell
# Run as Administrator
.\scripts\install_pvs_service.bat
```

### 2. Install the Scheduled Excel Refresh (one-time)
```powershell
# Run as Administrator with service account
.\scripts\install_scheduled_task.ps1 -ServiceAccount "DOMAIN\ServiceAccount"
```

### 3. Access the Dashboard
Open in browser: `http://VM_HOSTNAME:5051`

---

## Project Structure

```
Daily_Production_IT/
├── config/
│   └── settings.json          # All configuration variables
├── PVS/
│   ├── Planned_qtys.xlsx      # Schedule data (linked to external source)
│   └── ProdLine_Project_Map.csv
├── scripts/
│   ├── install_scheduled_task.ps1  # Creates scheduled task
│   ├── refresh_excel_scheduled.ps1 # Excel refresh (08:45 CET)
│   └── install_pvs_service.bat     # Installs Windows service
├── templates/
│   └── pvs.html               # Dashboard frontend (TV-optimized)
├── logs/                      # Service and refresh logs
├── .env                       # Database credentials
├── pvs_server.py              # Flask backend
└── nssm.exe                   # Service manager
```

---

## Configuration

All settings in `config/settings.json`:

| Setting | Value |
|---------|-------|
| Refresh Time | 08:45 CET, Mon-Fri |
| External Source | `G:\Logistics\6_Reporting\1_PVS\WH Receipt FY25.xlsx` |
| Display | 1920x1080 TV-optimized |
| Server Port | 5051 |

---

## Data Flow

```
08:45 CET (Mon-Fri)
    │
    ▼
[Scheduled Task] → Opens Planned_qtys.xlsx → Updates from WH Receipt FY25.xlsx → Saves
    │
    ▼
[PVS Service] → Reads updated Excel + SQL Server → Serves dashboard
    │
    ▼
[TV Display] → Auto-refreshes at 08:45
```

---

## Dashboard Colors

- **Adherence at 0%**: Dark/neutral (on target)
- **Positive deviation**: Brighter green as % increases
- **Negative deviation**: Brighter red as % decreases
- **Delta cells**: Green for positive, red for negative

---

## Maintenance

```powershell
# Service status
Get-Service AdientPVS

# Task status
Get-ScheduledTask -TaskName "AdientPVS_ExcelRefresh"

# Manual refresh
.\scripts\refresh_excel_scheduled.ps1

# Logs
Get-Content .\logs\excel_refresh.log -Tail 20
```

---

*Adient IT - Loznica Plant*
