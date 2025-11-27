<#
.SYNOPSIS
    Refreshes Planned_qtys.xlsx by opening Excel, updating external links, and saving.
    Designed to run as a scheduled task at 08:45 CET, Monday-Friday.

.DESCRIPTION
    This script:
    1. Opens the Planned_qtys.xlsx file
    2. Updates external links from WH Receipt FY25.xlsx
    3. Recalculates all formulas
    4. Saves and closes the workbook
    
    Run this script via Windows Task Scheduler under the service account.

.NOTES
    Author: Adient IT
    Schedule: 08:45 CET, Monday-Friday
#>

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURATION - Edit these paths as needed
# ============================================================================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

# Load config from settings.json if available
$ConfigPath = Join-Path $ProjectRoot "config\settings.json"
$PlannedXlsx = Join-Path $ProjectRoot "PVS\Planned_qtys.xlsx"
$ExternalSource = "G:\Logistics\6_Reporting\1_PVS\WH Receipt FY25.xlsx"
$LogFile = Join-Path $ProjectRoot "logs\excel_refresh.log"

if (Test-Path $ConfigPath) {
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        if ($Config.dataSources.externalSourceExcel) {
            $ExternalSource = $Config.dataSources.externalSourceExcel
        }
    } catch {
        Write-Warning "Could not load config: $_"
    }
}

# Ensure logs directory exists
$LogDir = Split-Path -Parent $LogFile
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
}

# ============================================================================
# LOGGING FUNCTION
# ============================================================================
function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
}

# ============================================================================
# MAIN REFRESH LOGIC
# ============================================================================
Write-Log "========== Excel Refresh Started =========="
Write-Log "Planned XLSX: $PlannedXlsx"
Write-Log "External Source: $ExternalSource"

# Check files exist
if (-not (Test-Path $PlannedXlsx)) {
    Write-Log "ERROR: Planned_qtys.xlsx not found at: $PlannedXlsx"
    exit 1
}

if (-not (Test-Path $ExternalSource)) {
    Write-Log "WARNING: External source file not found at: $ExternalSource"
    Write-Log "Proceeding anyway - links may not update correctly"
}

$Excel = $null
$Workbook = $null

try {
    Write-Log "Starting Excel COM automation..."
    
    # Create Excel instance
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.AskToUpdateLinks = $false
    
    Write-Log "Opening workbook with link updates..."
    # Open with UpdateLinks=3 (update external references)
    $Workbook = $Excel.Workbooks.Open($PlannedXlsx, 3)
    
    # Update all external links
    Write-Log "Updating external links..."
    try {
        $Links = $Workbook.LinkSources(1)  # 1 = xlExcelLinks
        if ($Links) {
            foreach ($Link in $Links) {
                Write-Log "  Updating link: $Link"
                $Workbook.UpdateLink($Link, 1)
            }
        }
    } catch {
        Write-Log "  No external links found or error updating: $_"
    }
    
    # Refresh all data connections
    Write-Log "Refreshing all data connections..."
    try {
        $Workbook.RefreshAll()
        Start-Sleep -Seconds 2  # Give time for async queries
    } catch {
        Write-Log "  RefreshAll warning: $_"
    }
    
    # Recalculate all formulas
    Write-Log "Recalculating formulas..."
    try {
        $Excel.CalculateFullRebuild()
    } catch {
        try {
            $Excel.CalculateFull()
        } catch {
            Write-Log "  Calculate warning: $_"
        }
    }
    
    # Save the workbook
    Write-Log "Saving workbook..."
    $Workbook.Save()
    
    Write-Log "SUCCESS: Excel refresh completed"
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "Stack: $($_.ScriptStackTrace)"
    exit 1
    
} finally {
    # Cleanup
    if ($Workbook) {
        try {
            $Workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        } catch {}
    }
    if ($Excel) {
        try {
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Log "========== Excel Refresh Finished =========="
}
