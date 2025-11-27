$ErrorActionPreference = "Stop"
Write-Host "`n=== Recalculate Excel Before Deploy ===`n" -ForegroundColor Cyan

$xlsxPath = Join-Path $PSScriptRoot "PVS\Planned_qtys.xlsx"

if (-not (Test-Path $xlsxPath)) {
    Write-Host "ERROR: Excel file not found at: $xlsxPath" -ForegroundColor Red
    exit 1
}

Write-Host "Opening Excel file: $xlsxPath" -ForegroundColor Yellow

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Opening workbook..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($xlsxPath, 3)  # UpdateLinks = 3
    
    Write-Host "Updating external links..." -ForegroundColor Yellow
    try {
        $links = $workbook.LinkSources(1)  # xlLinkTypeExcelLinks
        if ($links) {
            foreach ($link in $links) {
                Write-Host "  Updating: $link" -ForegroundColor Gray
                $workbook.UpdateLink($link, 1)
            }
        }
    } catch {
        Write-Host "  No external links or links already updated" -ForegroundColor Gray
    }
    
    Write-Host "Refreshing all data connections..." -ForegroundColor Yellow
    $workbook.RefreshAll()
    
    Write-Host "Calculating formulas..." -ForegroundColor Yellow
    $excel.CalculateUntilAsyncQueriesDone()
    try {
        $excel.CalculateFullRebuild()
    } catch {
        $excel.CalculateFull()
    }
    
    Write-Host "Saving workbook..." -ForegroundColor Yellow
    $workbook.Save()
    
    Write-Host "Closing workbook..." -ForegroundColor Yellow
    $workbook.Close($false)
    
    Write-Host "`nSUCCESS: Excel file recalculated and saved!" -ForegroundColor Green
    Write-Host "The file now contains cached calculated values that can be used without Excel COM." -ForegroundColor Green
    
} catch {
    Write-Host "`nERROR during Excel recalculation: $_" -ForegroundColor Red
    exit 1
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Copy the entire 'dist' folder to the TV PC" -ForegroundColor White
Write-Host "2. On TV PC, rename .env.tv to .env (to disable COM recalc)" -ForegroundColor White
Write-Host "3. Run Run-PVS.cmd on the TV PC" -ForegroundColor White
Write-Host ""
