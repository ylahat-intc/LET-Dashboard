# LET Lab Space Dashboard — Project Documentation

**Owner:** Yariv Lahat, Global Lab Planner, Intel REWS/Labs SME  
**Last Updated:** July 6, 2026  
**Project Folder:** `C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint\`  
**Live Dashboard:** https://ylahat-intc.github.io/LET-Dashboard/LET_Dashboard.html

---

## What Is This?

The LET (Lab Efficiency Team) system tracks all Intel lab space requests globally — new lab requests, reuse of existing space, lab returns, and capital projects.

This dashboard replaced a manual Excel-based process. The data now lives in a **SharePoint list** and this project builds a live interactive dashboard from it.

---

## Why It Exists

As Intel's sole Global Lab Planner, Yariv tracks:
- How many lab space requests come in each year
- How many were resolved by **reusing existing space** (instead of building new)
- How much money Intel saved by reusing space ($692/sqft benchmark for new lab construction)
- Cumulative cost avoidance over time (~$484M across all years)

---

## Data Sources

| Source | Records | IDs | Notes |
|--------|---------|-----|-------|
| SharePoint List | 1,053+ | 1474–2529+ | Live, updates daily |
| Excel History (`LET Vetting Working File.xlsm`) | 449 | 1001–1473 | Pre-migration records, static |
| **Total** | **1,502+** | **1001–2529+** | Zero duplicates confirmed |

**SharePoint List URL:**  
`https://intel.sharepoint.com/sites/GlobalLabsandDataCenters/Lists/LET2%20DLD%20Master%20Excel1/LET%20Review%20Meeting%20View.aspx`

**SharePoint Site ID:** `intel.sharepoint.com,8c7b402a-80d5-4df7-9355-5ce084b0e63d,a9bcafb0-ecf2-4d91-87f0-f6b11664510c`  
**List ID:** `781b2a30-b52a-4eda-b558-3f79056885a9`

---

## Key SharePoint Field Mapping

| SharePoint Field | Readable Name | Notes |
|-----------------|---------------|-------|
| `Title` | RequestID | e.g. "2504" |
| `field_2` | DateRequested | ISO format |
| `field_3` | Status | Raw text |
| `field_7` | BU | Business Unit |
| `field_10` | CPA | "YES"/"NO" |
| `field_13` | SqFtRequested | Float |
| `field_17` | Touch | |
| `field_20` | SiteTo | Destination site |
| `field_21` | SiteFrom | Origin site |
| `field_26` | Owner | Planner name |
| `field_28` | StartDate | |
| `field_29` | CloseDate | Used for cost-by-year grouping |
| `field_30` | Workstations | Small count, NOT sqft |
| `field_31` | SqFtAssigned | Actual assigned sqft (confirmed) |
| `field_32` | RoomAssigned | |

---

## SqFt & Cost Calculation Logic

- **$692/sqft** = Intel benchmark cost to build a new lab
- **Pre-2022 records:** Use `SqFtRequested` (Assigned was not tracked)
- **2022+ records:** Use `SqFtAssigned` (field_31), fallback to Requested if blank
- **Cost grouped by:** Close date year (when value was realized, not when requested)
- Implemented in `generate_dashboard.py` as `sqftForCost(r)` JavaScript function

---

## Files in This Project

| File | Purpose |
|------|---------|
| `refresh_data.py` | Pulls data from SharePoint via Graph API, saves JSON cache |
| `build_region_map.py` | Builds `site_region_map.json` from latest LaMP daily download (Spaces sheet) + manual overrides for legacy codes |
| `generate_dashboard.py` | Reads JSON caches + region map, builds self-contained HTML dashboard + `let_data.json` |
| `let_data_cache.json` | Live SharePoint data cache (auto-updated daily) |
| `let_data_cache_excel.json` | Static Excel history (IDs 1001–1473, never changes) |
| `let_data.json` | Cleaned, denormalized dataset consumed by the HTML dashboard |
| `site_region_map.json` | Building/Campus/Site → Region lookup (generated from LaMP) |
| `cresd_campus_map.json` | Reference: CRESD campus code → region/country/city (from `data (2).xlsx`) |
| `LET_Dashboard.html` | Generated dashboard (~1.0 MB, fully self-contained — works on file://) |
| `Refresh_Dashboard.bat` | One-click manual refresh: SharePoint pull → region map → dashboard → GitHub push |
| `Setup_AutoRefresh.ps1` | Registers Windows Task Scheduler job (run once) |
| `_auto_refresh_runner.ps1` | Script run by Task Scheduler daily at 12:00 PM |
| `azure_config.txt` | Azure connection string config (if Azure hosting set up) |
| `refresh_log.txt` | Log of all automatic refresh runs |
| `.msal_token_cache.bin` | MSAL auth token cache (do not delete — needed for silent auth) |

**LaMP source (read-only):**  
`C:\Users\ylahat\OneDrive - Intel Corporation\Documents - Global Lab Planners\General\LaMP Download Archive\*LaMP-daily-download.xlsx`  
`build_region_map.py` picks the latest file by filename (YYYYMMDD prefix) and reads the **Spaces** sheet, columns P–Y.

---

## Daily Automation

**Scheduled Task:** "LET Dashboard Daily Refresh"  
**Runs:** Every day at **12:00 PM**  
**Runner script:** `_auto_refresh_runner.ps1` (4 steps + source sync, logs to `refresh_log.txt`)

**What it does:**
1. **[Step 0]** `git pull` the repo → restores `generate_dashboard.py` / `refresh_data.py` if VS Code accidentally reverted them
2. **[Step 1]** Pulls fresh data from SharePoint (Graph API, MSAL silent auth) → `let_data_cache.json`
3. **[Step 2]** Generates dashboard (`generate_dashboard.py`) → `LET_Dashboard.html` + `let_data.json`
4. **[Step 3]** Pushes to GitHub Pages — HTML, data, **and updated source files** → https://github.com/ylahat-intc/LET-Dashboard

**Live URL** (refreshed automatically each day):  
https://ylahat-intc.github.io/LET-Dashboard/LET_Dashboard.html

**To verify it's scheduled:**
```powershell
Get-ScheduledTaskInfo -TaskName "LET Dashboard Daily Refresh"
```

**To run manually (full pipeline, same as the scheduled run):**
```cmd
Refresh_Dashboard.bat
```

**To run manually (PowerShell version — same as scheduled task):**
```powershell
powershell.exe -NonInteractive -File _auto_refresh_runner.ps1
```

> ⚠ **Important:** the runner script uses *hard-coded absolute paths* at the top (`$ScriptDir`, `$LogFile`, `$PythonExe`). If the project folder ever moves, edit those three lines OR re-run `Setup_AutoRefresh.ps1` (which regenerates the runner with new paths).

---

## Authentication (MSAL)

- Uses Microsoft MSAL library with persistent token cache
- **Scope:** `Sites.ReadWrite.All` (read list + upload HTML back to SharePoint)
- **Token cache file:** `.msal_token_cache.bin` — valid ~90 days
- First run requires browser login via device flow
- After first login: fully silent/automatic

**If token expires (re-login needed):**
```cmd
del .msal_token_cache.bin
py refresh_data.py
```
Follow the device code login prompt, then run `--full` again.

---

## Dashboard Features

### Global Filters (sidebar)
- **Status** — multi-select checkboxes (includes **Hold** as its own status, amber color)
- **Region** — AMER / APAC / EMEA / Unknown (resolved from CRESD campus codes, 937 entries)
- **BU** — multi-select dropdown
- **Planner/Owner** — multi-select dropdown
- **Site** — multi-select dropdown (CRESD-normalized codes)
- **Trade Classification** — Net Zero / Growth / Undefined / Not Classified (shown only when classified records exist)
- **Date range** — Year + Month dropdowns (Start → End)
- **Quarter** — Q1/Q2/Q3/Q4 toggle buttons (**bug-fixed July 2026**: quarter is matched to the date branch that admitted the row, not always the submission quarter)
- **CPA** — toggle

### View Presets (top bar)
Click to filter instantly — can combine multiple: All, Open/Active, Closed, Reuse, Lab Return, Capital/CRE, Withdrawn

### Charts
- Cost Avoidance by Year (bar, grouped by **close** year)
- Closure Categories (donut)
- Monthly New vs Closed (line + bar)
- Quarterly Volume by Year (stacked bar)
- Top BUs — Total vs Reuse
- Planner Workload — Open vs Total
- Pipeline Health: Aging / Touch / SqFt distribution
- Site Demand: Top Destination + Origin sites
- **Trade Net Zero vs Growth** (shown only when classified records exist) — doughnut + SqFt bar, tracked since 2025-12-02

### KPI Cards (all filter-responsive)
💰 Cost Avoided · ♻️ Reuse Rate · 📐 SqFt Reused · 🏗️ Capital/CRE · 📋 Total · ⏱️ Avg Days · ⚡ High Touch

### Bottom Table
- **Per-table filters** (independent of sidebar): free-text search, status dropdown, submit/close date ranges, days-open min/max, Reset button
- **Columns:** ID · Title · Status · BU · **Region** · **Site** · **Trade** · SqFt Req · SqFt Asn · **💰 Cost** · Planner · Submitted · Closed · Q · Days
- **💰 Cost column** (Reuse rows only): 2022+ uses Assigned sqft, pre-2022 uses Requested. All × $692/sqft. Hover tooltip shows the exact sqft field used.
  - Example: Req 2485 → 580 sqft × $692 = **$401,360** (used Assigned, not the 10,000 Requested)
- **ID and Title are clickable links** → open the specific item in SharePoint via DispForm.aspx

### Data Quality Banner
Every build runs sanity checks and shows a **green (✅ clean) or amber/red (⚠️ issues) banner** at the top of the dashboard. Checks:
- `quarter` must match month of `date_req`
- `quarter_close` / `year_close` must match `date_close`
- Closed items should have a `date_close`
- `is_open=True` items should not have a `date_close`
- `sqft_asn` should not be >4× `sqft_req`

As of July 2026: **0 errors, 26 warnings** (all are legacy data gaps in older records).

### Removed (June 2026)
- **Touch** KPI / donut / sidebar filter / table column — requestors not using it consistently; replaced with Region.

---

## Current Hosting Situation

| Location | URL | Status |
|----------|-----|--------|
| Local PC | `...\LET sharepoint\LET_Dashboard.html` | ✅ Opens directly, refreshes at noon |
| SharePoint Documents | `https://intel.sharepoint.com/sites/GlobalLabsandDataCenters/Shared%20Documents/LET%20Dashboard/LET_Dashboard.html` | ⚠️ Downloads file (SP security policy) |
| Azure Static Website | Not configured | ❌ Needs Azure Contributor permission |
| Intel GitHub Pages | Not configured | ❌ No access to github.intel.com |
| Intel IIS/Apache server | Unknown | 🔍 Ask IT team |

**Pending:** Find an Intel internal web server to host the file so it opens directly in browser without downloading.

---

## Data Integrity — Root Cause Analysis and Prevention

### What happened (July 6, 2026 regression)

The dashboard went back ~5 versions because:

1. **`generate_dashboard.py` was not tracked in git.** The GitHub repo only contained built artifacts (HTML, JSON). The Python source lived only on the local machine with no version history.

2. **VS Code editor buffer conflict.** When Copilot used `replace_string_in_file` to modify `generate_dashboard.py` during a coding session, VS Code had an older version of the file open in an editor tab. VS Code detected the file changed on disk and either auto-saved its buffered (old) version or the user clicked “Revert”. This silently overwrote the new content back to the old state.

3. **No detection.** Because the source wasn’t in git, there was no diff, no warning, and no way to know the file had regressed. The built HTML was current but the source that would regenerate it was weeks old.

**Result:** The next `generate_dashboard.py` run (e.g., from the daily Task Scheduler job) rebuilt the HTML from the stale source, overwriting all accumulated features (Hold status, CRESD normalization, Trade tracking, QA checks, per-row cost column, quarter-filter fix) on GitHub Pages.

---

### Prevention (now in place)

| Fix | What it does |
|-----|--------------|
| **Source in git** | `generate_dashboard.py`, `refresh_data.py`, `_auto_refresh_runner.ps1`, `Refresh_Dashboard.bat` are now committed to the GitHub repo. They are the authoritative version. |
| **Auto-refresh pulls source first** | `_auto_refresh_runner.ps1` now runs `git pull` at Step 0 and copies `generate_dashboard.py` / `refresh_data.py` from the repo to the working directory if their hashes differ. VS Code can no longer silently revert them — the next scheduled run will restore the correct version. |
| **VERSION stamp** | `generate_dashboard.py` has `VERSION: YYYY-MM-DD` in its docstring. Check that it matches the latest commit date if something looks wrong. |
| **.gitignore** | Secrets (`azure_config.txt`, `.msal_token_cache.bin`) are explicitly excluded from the repo. |
| **QA banner in HTML** | Every build runs sanity checks; results visible in the live dashboard header. |

### If it happens again — recovery

```powershell
# Restore source files from git repo to working directory
$repo = "$env:TEMP\LET-Dashboard-push"
git -C $repo pull
Copy-Item "$repo\generate_dashboard.py" "C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint\generate_dashboard.py" -Force
Copy-Item "$repo\refresh_data.py"       "C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint\refresh_data.py" -Force
# Then rebuild:
cd "C:\Users\ylahat\OneDrive - Intel Corporation\Documents\CS GLP\Dev soft project\LET sharepoint"
$env:PYTHONUTF8 = '1'; py generate_dashboard.py
```

### Rules going forward

1. **Never ignore "file changed on disk" warnings in VS Code** — click **Load File** (not Revert/Overwrite) when editing Python files that the AI assistant is also modifying.
2. **Run `Refresh_Dashboard.bat` after any Python source change** — this pushes the new source to git alongside the built HTML.
3. **When starting a new coding session**, verify `VERSION:` at the top of `generate_dashboard.py` matches the last commit date. If not, run the recovery command above.

---

## Known Issues & Notes

- **Emoji in Task Scheduler:** Fixed by setting `PYTHONUTF8=1` in the runner script
- **Date parsing:** SharePoint dates are ISO format (`2026-01-29T08:00:00Z`) — parsed by splitting on `T` and taking first 10 chars
- **CPA field:** SharePoint uses `"YES"`/`"NO"` (not `TRUE`/`FALSE` like old system)
- **field_30 ≠ SqFtAssigned:** field_30 = workstation count (small integers). field_31 = SqFtAssigned (confirmed from live data)
- **Status normalization:** "Closed: Filled with existing lab space" → counted as Reuse; two variants of "Withdrawn" both handled

---

## Next Steps / Roadmap

- [ ] **Find internal web hosting** — ask IT team for IIS/Apache server with a folder
- [ ] **Power BI** — connect directly to SharePoint list, replace HTML dashboard
  - Estimated: 3–4 hours setup, can reuse same field mappings
  - No hosting needed — always live via Power BI Service
- [ ] **Azure hosting** — if someone with Contributor access can create a storage account
- [ ] **Expand Excel history** — currently only 1001–1473; older records (pre-1001) if they exist

---

## Quick Reference — Run Commands

```cmd
# Full refresh (pull SP + build HTML + upload to SP)
py refresh_data.py --full

# Just pull data (no HTML rebuild)
py refresh_data.py

# Just rebuild HTML from existing cache
py generate_dashboard.py

# Check scheduled task
powershell -Command "Get-ScheduledTaskInfo -TaskName 'LET Dashboard Daily Refresh'"

# Re-register scheduled task (if needed)
powershell -ExecutionPolicy Bypass -File Setup_AutoRefresh.ps1
```

---

## Contact / Access

- **SharePoint Site:** Global Labs and Data Centers (GLDC)
- **List:** LET2 DLD Master Excel1
- **Owner:** Yariv Lahat (ylahat@intel.com)

---

## Region Mapping (LaMP + Manual Overrides)

Building / campus / site codes in SharePoint are resolved to a Region (AMER / APAC / EMEA) for filtering and charting. The lookup table is rebuilt daily from the LaMP **Spaces** sheet (cols P–Y) plus a curated list of overrides in `build_region_map.py` for legacy or free-text codes that aren't in LaMP.

**Override formats** (in `MANUAL_BUILDING_OVERRIDES`):
1. `"CODE": "campus:XX"` — alias to an existing LaMP campus
2. `"CODE": "site:XX"` — alias to an existing LaMP site
3. `"CODE": {"region": "...", "site": "...", "site_name": "..."}` — explicit (creates a new site code if needed)

**Authoritative reference:** `cresd_campus_map.json` (built from `data (2).xlsx`, CRESD campus export — 937 campus codes with Region/Country/City). Use this to verify any new code before adding it as an override.

**Current state** (~16% Unknown remaining, mostly free-text junk like `CA`, `NEW`, `NA`, blanks):

| Region   | Rows |
| -------- | ---: |
| AMER     | 661  |
| APAC     | 362  |
| EMEA     | 272  |
| Unknown  | 211  |

**New site codes introduced via overrides** (not present in LaMP): `FC` Fort Collins, `TK` Tokyo, `PA` Allentown, `UK` Swindon, `HY` Hyderabad, `SD` San Diego.

---

## Change Log

### June 29, 2026
- **Region** support end-to-end: sidebar filter, bar chart, table column, derived from LaMP + manual overrides (`build_region_map.py`, `site_region_map.json`)
- **CRESD reference** loaded from `data (2).xlsx` → `cresd_campus_map.json` to verify legacy codes
- **SKC / SKC1 corrected** — was Santa Clara, actually Hyderabad India (per CRESD)
- **Bottom-table filters** added (search, status, dates, days-open, Reset)
- **SharePoint links** on ID and Title cells (DispForm.aspx instead of raw `/{id}_.000`)
- **Inline data** embedded in HTML (`window.__LET_DATA__`) so dashboard works on `file://` double-click without CORS errors
- **Touch removed** from KPI / chart / filter / column
- **Auto-refresh fixed** — `_auto_refresh_runner.ps1` was using single-quoted strings around variables (no expansion), task reported success but nothing ran since 2026-06-25. Now uses absolute paths baked in at the top of the script. `Setup_AutoRefresh.ps1` template also fixed.
- **Pipeline now includes region-map rebuild** as step 2 (both `_auto_refresh_runner.ps1` and `Refresh_Dashboard.bat`)
