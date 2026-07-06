$ScriptDir = "C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint"
$LogFile   = Join-Path $ScriptDir "refresh_log.txt"
$PythonExe = "C:\WINDOWS\py.exe"
$RepoDir   = "$env:TEMP\LET-Dashboard-push"

Set-Location $ScriptDir
$env:PYTHONUTF8 = '1'
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Add-Content -Path $LogFile -Value "`n===== $timestamp ====="

# Step 0 - Sync source files from GitHub repo (prevents silent regression
#          if VS Code editor buffer overwrites tool-modified Python files).
#          The repo now tracks generate_dashboard.py — this guarantees we
#          always build from the committed (correct) version.
Add-Content -Path $LogFile -Value '[0/4] Syncing source from GitHub...'
if (-not (Test-Path "$RepoDir\.git")) {
    git clone https://github.com/ylahat-intc/LET-Dashboard.git "$RepoDir" 2>&1 | Add-Content -Path $LogFile
}
git -C "$RepoDir" pull --quiet 2>&1 | Add-Content -Path $LogFile
# Copy source files from repo → working directory (repo is authoritative)
foreach ($src in @('generate_dashboard.py','refresh_data.py')) {
    $repoFile = Join-Path $RepoDir $src
    $localFile = Join-Path $ScriptDir $src
    if (Test-Path $repoFile) {
        $repoHash  = (Get-FileHash $repoFile  -Algorithm SHA256).Hash
        $localHash = if (Test-Path $localFile) { (Get-FileHash $localFile -Algorithm SHA256).Hash } else { '' }
        if ($repoHash -ne $localHash) {
            Copy-Item $repoFile $localFile -Force
            Add-Content -Path $LogFile -Value "  Updated $src from repo (was out of sync)"
        }
    }
}

# Step 1 - Pull from SharePoint
Add-Content -Path $LogFile -Value '[1/4] Pulling SharePoint data...'
& $PythonExe (Join-Path $ScriptDir "refresh_data.py") 2>&1 | Add-Content -Path $LogFile

# Step 2 - Generate dashboard + let_data.json
Add-Content -Path $LogFile -Value '[2/4] Generating dashboard...'
& $PythonExe (Join-Path $ScriptDir "generate_dashboard.py") 2>&1 | Add-Content -Path $LogFile

# Keep the .aspx copy in sync (for SharePoint embeds)
Copy-Item (Join-Path $ScriptDir 'LET_Dashboard.html') (Join-Path $ScriptDir 'LET_Dashboard.aspx') -Force -ErrorAction SilentlyContinue

# Step 3 - Push to GitHub Pages (HTML + data + source)
Add-Content -Path $LogFile -Value '[3/4] Pushing to GitHub Pages...'
Copy-Item (Join-Path $ScriptDir 'LET_Dashboard.html')     (Join-Path $RepoDir 'LET_Dashboard.html')     -Force
Copy-Item (Join-Path $ScriptDir 'let_data.json')           (Join-Path $RepoDir 'let_data.json')           -Force
Copy-Item (Join-Path $ScriptDir 'Designer (14).png')       (Join-Path $RepoDir 'favicon.png')             -Force -ErrorAction SilentlyContinue
# Also push source so repo stays current
Copy-Item (Join-Path $ScriptDir 'generate_dashboard.py')   (Join-Path $RepoDir 'generate_dashboard.py')   -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $ScriptDir 'refresh_data.py')         (Join-Path $RepoDir 'refresh_data.py')         -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $ScriptDir '_auto_refresh_runner.ps1')(Join-Path $RepoDir '_auto_refresh_runner.ps1') -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $ScriptDir 'Refresh_Dashboard.bat')   (Join-Path $RepoDir 'Refresh_Dashboard.bat')   -Force -ErrorAction SilentlyContinue
git -C "$RepoDir" add LET_Dashboard.html let_data.json favicon.png `
    generate_dashboard.py refresh_data.py _auto_refresh_runner.ps1 Refresh_Dashboard.bat 2>&1 | Add-Content -Path $LogFile
$dateTag = Get-Date -Format 'yyyy-MM-dd'
git -C "$RepoDir" commit -m "Auto-refresh $dateTag" 2>&1 | Add-Content -Path $LogFile
git -C "$RepoDir" push origin main 2>&1 | Add-Content -Path $LogFile
Add-Content -Path $LogFile -Value '[4/4] Done.'

