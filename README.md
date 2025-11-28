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

## Netlify Static Deployment (No VM Required)

For hosting on Netlify without a VM:

### One-Time Setup

1. **Install Netlify CLI**:
   ```bash
   npm install -g netlify-cli
   netlify login
   ```

2. **Initial deploy** (creates site):
   ```bash
   cd netlify_static
   netlify deploy --prod --dir=.
   ```
   Note the site URL and ID for future deployments.

3. **Install automated task** (optional):
   ```powershell
   # Run as Administrator
   .\scripts\install_netlify_task.ps1 -NetlifySiteId "YOUR_SITE_ID"
   ```

### Manual Deployment

Double-click `deploy_now.bat` or run:
```bash
python scripts/generate_static_pvs.py
cd netlify_static && netlify deploy --prod --dir=.
```

### Automated Daily Deployment

The scheduled task runs at **08:45 CET Mon-Fri**:
1. Generates fresh snapshot from WH Receipt FY25.xlsx
2. Deploys to Netlify automatically

---

## Maintenance

```powershell
# Service status (VM deployment)
Get-Service AdientPVS

# Netlify task status
Get-ScheduledTask -TaskName "PVS_Netlify_Deploy"

# Manual snapshot generation
python scripts\generate_static_pvs.py

# Logs
Get-Content .\logs\netlify_deploy.log -Tail 20
```

---

*Adient IT - Loznica Plant*
