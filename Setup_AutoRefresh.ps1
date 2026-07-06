# Setup_AutoRefresh.ps1
# Run this ONCE to register a daily auto-refresh of the LET Dashboard.
# The dashboard HTML will be rebuilt every day at 7:30 AM automatically.

$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe   = (Get-Command py -ErrorAction SilentlyContinue).Source
$RefreshPy   = Join-Path $ScriptDir "refresh_data.py"
$GeneratePy  = Join-Path $ScriptDir "generate_dashboard.py"
$TaskName    = "LET Dashboard Daily Refresh"

if (-not $PythonExe) {
    Write-Host "ERROR: Python (py) not found. Install Python first." -ForegroundColor Red
    exit 1
}

# Build the action: run all scripts in sequence, log output.
# The helper is written with absolute paths baked in so it is fully self-contained
# (avoids variable-scoping problems when launched by Task Scheduler).
$LogFile = Join-Path $ScriptDir "refresh_log.txt"
$Script = @"
`$ScriptDir = "$ScriptDir"
`$LogFile   = "$LogFile"
`$PythonExe = "$PythonExe"

Set-Location `$ScriptDir
`$env:PYTHONUTF8 = '1'
`$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Add-Content -Path `$LogFile -Value "``n===== `$timestamp ====="

# Step 1 - Pull from SharePoint
Add-Content -Path `$LogFile -Value '[1/4] Pulling SharePoint data...'
& `$PythonExe (Join-Path `$ScriptDir "refresh_data.py") 2>&1 | Add-Content -Path `$LogFile

# Step 2 - Rebuild region map from latest LaMP daily download
Add-Content -Path `$LogFile -Value '[2/4] Rebuilding region map...'
& `$PythonExe (Join-Path `$ScriptDir "build_region_map.py") 2>&1 | Add-Content -Path `$LogFile

# Step 3 - Generate dashboard + let_data.json
Add-Content -Path `$LogFile -Value '[3/4] Generating dashboard...'
& `$PythonExe (Join-Path `$ScriptDir "generate_dashboard.py") 2>&1 | Add-Content -Path `$LogFile

# Keep the .aspx copy in sync (for SharePoint embeds)
Copy-Item (Join-Path `$ScriptDir 'LET_Dashboard.html') (Join-Path `$ScriptDir 'LET_Dashboard.aspx') -Force -ErrorAction SilentlyContinue

# Step 4 - Push to GitHub Pages
Add-Content -Path `$LogFile -Value '[4/4] Pushing to GitHub Pages...'
`$RepoDir = "`$env:TEMP\LET-Dashboard-push"
if (-not (Test-Path "`$RepoDir\.git")) {
    git clone https://github.com/ylahat-intc/LET-Dashboard.git "`$RepoDir" 2>&1 | Add-Content -Path `$LogFile
}
git -C "`$RepoDir" pull --quiet 2>&1 | Add-Content -Path `$LogFile
Copy-Item (Join-Path `$ScriptDir 'LET_Dashboard.html') (Join-Path `$RepoDir 'LET_Dashboard.html') -Force
Copy-Item (Join-Path `$ScriptDir 'let_data.json')      (Join-Path `$RepoDir 'let_data.json')      -Force
Copy-Item (Join-Path `$ScriptDir 'Designer (14).png')  (Join-Path `$RepoDir 'favicon.png')        -Force -ErrorAction SilentlyContinue
git -C "`$RepoDir" add LET_Dashboard.html let_data.json favicon.png 2>&1 | Add-Content -Path `$LogFile
`$dateTag = Get-Date -Format 'yyyy-MM-dd'
git -C "`$RepoDir" commit -m "Auto-refresh `$dateTag" 2>&1 | Add-Content -Path `$LogFile
git -C "`$RepoDir" push origin main 2>&1 | Add-Content -Path `$LogFile
Add-Content -Path `$LogFile -Value 'Done.'
"@

# Save as a helper script
$HelperScript = Join-Path $ScriptDir "_auto_refresh_runner.ps1"
$Script | Set-Content $HelperScript -Encoding UTF8

# Register Task Scheduler job — daily at 07:30
$Action  = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -File `"$HelperScript`""

$Trigger = New-ScheduledTaskTrigger -Daily -At "12:00PM"

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

# Remove old task if exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName  $TaskName `
    -Action    $Action `
    -Trigger   $Trigger `
    -Settings  $Settings `
    -Description "Refreshes LET Dashboard HTML from SharePoint list daily at 07:30" | Out-Null

Write-Host ""
Write-Host "✅  Task '$TaskName' registered." -ForegroundColor Green
Write-Host "    Runs every day at 07:30 AM." -ForegroundColor Green
Write-Host "    Log file: $LogFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "IMPORTANT: First time still needs manual login." -ForegroundColor Yellow
Write-Host "   Run this now to do the one-time login:" -ForegroundColor Yellow
Write-Host "   py refresh_data.py" -ForegroundColor White
Write-Host ""
Write-Host "After that, all future refreshes are fully automatic." -ForegroundColor Green
