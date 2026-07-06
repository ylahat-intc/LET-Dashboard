$ScriptDir = "C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint"
$LogFile   = Join-Path $ScriptDir "refresh_log.txt"
$PythonExe = "C:\WINDOWS\py.exe"

Set-Location $ScriptDir
$env:PYTHONUTF8 = '1'
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Add-Content -Path $LogFile -Value "`n===== $timestamp ====="

# Step 1 - Pull from SharePoint
Add-Content -Path $LogFile -Value '[1/4] Pulling SharePoint data...'
& $PythonExe (Join-Path $ScriptDir "refresh_data.py") 2>&1 | Add-Content -Path $LogFile

# Step 2 - Rebuild region map from latest LaMP daily download
Add-Content -Path $LogFile -Value '[2/4] Rebuilding region map...'
& $PythonExe (Join-Path $ScriptDir "build_region_map.py") 2>&1 | Add-Content -Path $LogFile

# Step 3 - Generate dashboard + let_data.json
Add-Content -Path $LogFile -Value '[3/4] Generating dashboard...'
& $PythonExe (Join-Path $ScriptDir "generate_dashboard.py") 2>&1 | Add-Content -Path $LogFile

# Keep the .aspx copy in sync (for SharePoint embeds)
Copy-Item (Join-Path $ScriptDir 'LET_Dashboard.html') (Join-Path $ScriptDir 'LET_Dashboard.aspx') -Force -ErrorAction SilentlyContinue

# Step 4 - Push to GitHub Pages
Add-Content -Path $LogFile -Value '[4/4] Pushing to GitHub Pages...'
$RepoDir = "$env:TEMP\LET-Dashboard-push"
if (-not (Test-Path "$RepoDir\.git")) {
    git clone https://github.com/ylahat-intc/LET-Dashboard.git "$RepoDir" 2>&1 | Add-Content -Path $LogFile
}
git -C "$RepoDir" pull --quiet 2>&1 | Add-Content -Path $LogFile
Copy-Item (Join-Path $ScriptDir 'LET_Dashboard.html') (Join-Path $RepoDir 'LET_Dashboard.html') -Force
Copy-Item (Join-Path $ScriptDir 'let_data.json')      (Join-Path $RepoDir 'let_data.json')      -Force
Copy-Item (Join-Path $ScriptDir 'Designer (14).png')  (Join-Path $RepoDir 'favicon.png')        -Force -ErrorAction SilentlyContinue
git -C "$RepoDir" add LET_Dashboard.html let_data.json favicon.png 2>&1 | Add-Content -Path $LogFile
$dateTag = Get-Date -Format 'yyyy-MM-dd'
git -C "$RepoDir" commit -m "Auto-refresh $dateTag" 2>&1 | Add-Content -Path $LogFile
git -C "$RepoDir" push origin main 2>&1 | Add-Content -Path $LogFile
Add-Content -Path $LogFile -Value 'Done.'

