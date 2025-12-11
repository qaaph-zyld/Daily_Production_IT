# PVS Automated Email - Service Account Setup Guide

This guide explains how to set up automated daily PVS email reports on a Windows VM without Outlook.

---

## Overview

The automated email system sends PVS reports at **8:45 AM daily** using:
- **Script:** `scripts/send_pvs_email_auto.py`
- **Method:** Direct SMTP (no Outlook required)
- **Scheduler:** Windows Task Scheduler

---

## Prerequisites

1. **Windows VM** with network access to:
   - SQL Server (`a265m001`)
   - SMTP relay (`smtp.adient.com` or your corporate relay)
   - File share (`G:\Logistics\...`)

2. **Python 3.9+** installed (preferably via virtual environment)

3. **Service Account** (optional but recommended):
   - Domain account with "Log on as a batch job" rights
   - No need for Outlook or email client installed

---

## Step 1: Prepare the Environment

### Option A: Use Virtual Environment (Recommended)

```powershell
# Navigate to project directory
cd G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT

# Create virtual environment
python -m venv .venv

# Activate it
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

### Option B: Use System Python

Ensure all dependencies from `requirements.txt` are installed globally.

---

## Step 2: Configure Email Settings

Edit `config/settings.json`:

```json
{
  "email": {
    "smtp_server": "smtp.adient.com",
    "smtp_port": 25,
    "from_address": "pvs-dashboard@adient.com",
    "recipients": [
      "nikola.jelacic@adient.com",
      "milan.gajic@adient.com",
      ...
    ],
    "test_recipient": "nikola.jelacic@adient.com"
  }
}
```

### SMTP Configuration Notes

| Setting | Description |
|---------|-------------|
| `smtp_server` | Corporate SMTP relay (typically allows unauthenticated relay from internal IPs) |
| `smtp_port` | Usually `25` for internal relay, or `587` for authenticated |
| `from_address` | Can be any valid address; some relays require a real mailbox |
| `test_recipient` | Used when running with `--test` flag |

---

## Step 3: Test the Script Manually

```powershell
# Activate virtual environment first
cd G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT
.\.venv\Scripts\Activate.ps1

# Run in test mode (sends to test_recipient only)
python scripts/send_pvs_email_auto.py --test

# Check the log
Get-Content logs/email_auto.log -Tail 50
```

If successful, you should see:
```
INFO - Email sent successfully to 1 recipients
```

---

## Step 4: Create the Scheduled Task

### Option A: Run as SYSTEM (No Password Required)

```powershell
# Run as Administrator
cd G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT\scripts
powershell -ExecutionPolicy Bypass -File install_email_task.ps1
```

### Option B: Run as Service Account (Recommended for Production)

```powershell
# Run as Administrator
powershell -ExecutionPolicy Bypass -File install_email_task.ps1 -ServiceAccount "DOMAIN\svc_pvs_email"
```

You will be prompted for the service account password.

### Option C: Create Task Manually

1. Open **Task Scheduler** (`taskschd.msc`)
2. Click **Create Task**
3. General tab:
   - Name: `PVS_Daily_Email`
   - Run whether user is logged on or not
   - Configure for: Windows Server 2016/2019/2022
4. Triggers tab:
   - New trigger: Daily at 8:45 AM
5. Actions tab:
   - Action: Start a program
   - Program: `G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT\.venv\Scripts\python.exe`
   - Arguments: `scripts\send_pvs_email_auto.py`
   - Start in: `G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT`
6. Conditions tab:
   - Start only if network is available: ✓
7. Settings tab:
   - Allow task to be run on demand: ✓
   - If task fails, restart every: 5 minutes (up to 3 times)

---

## Step 5: Service Account Permissions

If using a service account, ensure it has:

1. **Local Security Policy** (`secpol.msc`):
   - "Log on as a batch job" right

2. **File System Access**:
   - Read/Write to `G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT\logs\`
   - Read to `G:\Logistics\6_Reporting\1_PVS\WH Receipt FY25.xlsx`

3. **SQL Server Access**:
   - Read access to `QADEE2798` database on `a265m001`

4. **Network Access**:
   - Outbound TCP port 25 to SMTP server

---

## Step 6: Verify the Setup

### Check Task Status

```powershell
schtasks /query /tn "PVS_Daily_Email" /v
```

### Run Task Manually

```powershell
schtasks /run /tn "PVS_Daily_Email"
```

### Check Logs

```powershell
Get-Content G:\Logistics\4_BSA\9_Scripts_Reports\Daily_Production_IT\logs\email_auto.log -Tail 100
```

---

## Troubleshooting

### Email Not Sending

1. **Check SMTP connectivity:**
   ```powershell
   Test-NetConnection smtp.adient.com -Port 25
   ```

2. **Check from_address:**
   Some SMTP relays require the from address to be a real mailbox.

3. **Check logs:**
   ```powershell
   Get-Content logs/email_auto.log -Tail 100
   ```

### Task Not Running

1. **Check task history in Task Scheduler**
2. **Verify service account has "Log on as a batch job" right**
3. **Ensure Python path is correct in task action**

### Database Connection Errors

1. **Verify SQL Server connectivity:**
   ```powershell
   Test-NetConnection a265m001 -Port 1433
   ```

2. **Check credentials in `.env` file**

---

## Uninstall

To remove the scheduled task:

```powershell
powershell -ExecutionPolicy Bypass -File install_email_task.ps1 -Uninstall
```

---

## Files Reference

| File | Purpose |
|------|---------|
| `scripts/send_pvs_email_auto.py` | SMTP-only email sender (no Outlook) |
| `scripts/install_email_task.ps1` | Creates Windows scheduled task |
| `config/settings.json` | Email configuration |
| `logs/email_auto.log` | Email sending logs |

---

## Contact

For issues, contact: nikola.jelacic@adient.com
