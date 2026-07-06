# LET Dashboard — Power BI Setup Guide
## Connect Your SharePoint List to a Live, Interactive Power BI Dashboard

---

## STEP 1 — Get Power BI Desktop (Free)
1. Go to: https://powerbi.microsoft.com/desktop
2. Click **Download Free** → Install
3. Sign in with your **Intel (@intel.com)** account
   → Your M365 license includes Power BI Pro

---

## STEP 2 — Connect to the SharePoint List

1. Open Power BI Desktop
2. Click **Home → Get Data → SharePoint Online List**
3. Enter Site URL:
   ```
   https://intel.sharepoint.com/sites/GlobalLabsandDataCenters
   ```
4. Click **OK** → Sign in with Intel account if prompted
5. In the Navigator, check: **LET Lab Space Request/Release**
6. Click **Transform Data** (not Load directly — we need to clean fields)

---

## STEP 3 — Transform Data in Power Query

In the Power Query editor, apply these steps:

### Rename key columns
| Original Field | Rename To |
|---|---|
| Title | RequestID |
| field_2 | DateRequested |
| field_3 | Status |
| field_7 | BU |
| field_8 | Group |
| field_9 | Permanent |
| field_10 | CPA |
| field_13 | SqFtRequested |
| field_17 | Touch |
| field_20 | SiteTo |
| field_21 | SiteFrom |
| field_26 | Planner |
| field_27 | RequestName |
| field_29 | DateClosed |
| field_30 | SqFtAssigned |

### Add Calculated Columns (in Power Query)
```
// Normalized Status
= if [Status] = "Closed: Filled with existing lab space" then "Reuse"
  else if Text.Contains([Status], "return back") then "Lab Return"
  else if Text.Contains([Status], "require CRE") or Text.Contains([Status], "Capital") then "Capital/CRE"
  else if Text.Contains([Status], "ithdrawn") then "Withdrawn"
  else if Text.Contains([Status], "enied") or Text.Contains([Status], "nsupportable") then "Denied"
  else if [Status] = null or [Status] = "" then "Open/Pending"
  else "Other"
```

4. Click **Close & Apply**

---

## STEP 4 — Create DAX Measures

In the **Model** view, create these measures:

```dax
// ── COST AVOIDANCE ──────────────────────────────────────────────────────────
Cost Avoided ($) =
VAR reuseRows = FILTER(ALL_Requests, [Status_Norm] = "Reuse")
VAR sqft = SUMX(reuseRows, MAX([SqFtRequested], 0))
RETURN sqft * 692

// ── REUSE RATE ──────────────────────────────────────────────────────────────
Reuse Rate % =
VAR reuse  = CALCULATE(COUNTROWS(ALL_Requests), [Status_Norm] = "Reuse")
VAR closed = CALCULATE(COUNTROWS(ALL_Requests), [Status_Norm] <> "Open/Pending")
RETURN DIVIDE(reuse, closed, 0)

// ── AVG DAYS TO CLOSE ────────────────────────────────────────────────────────
Avg Days to Close =
AVERAGEX(
    FILTER(ALL_Requests,
        NOT ISBLANK([DateRequested]) && NOT ISBLANK([DateClosed]) &&
        DATEDIFF([DateRequested],[DateClosed],DAY) >= 0
    ),
    DATEDIFF([DateRequested],[DateClosed],DAY)
)

// ── SQFT REUSED ───────────────────────────────────────────────────────────────
SqFt Reused =
CALCULATE(
    SUMX(ALL_Requests, MAX([SqFtRequested],0)),
    [Status_Norm] = "Reuse"
)

// ── SQFT RETURNED ─────────────────────────────────────────────────────────────
SqFt Returned to Inventory =
CALCULATE(
    SUMX(ALL_Requests, ABS(MIN([SqFtRequested],0))),
    [Status_Norm] = "Lab Return"
)

// ── OPEN REQUESTS ─────────────────────────────────────────────────────────────
Open Requests = CALCULATE(COUNTROWS(ALL_Requests), [Status_Norm] = "Open/Pending")

// ── CPA COUNT ─────────────────────────────────────────────────────────────────
CPA Requests = CALCULATE(COUNTROWS(ALL_Requests), [CPA] = "YES")
```

---

## STEP 5 — Build the Report Pages

### Page 1: Executive Summary
- **Card visuals**: Cost Avoided, Reuse Rate, SqFt Reused, Avg Days to Close
- **Donut chart**: Closure Categories (Status_Norm)
- **Bar chart**: Cost Avoided by Year
- **Line chart**: Monthly request volume

### Page 2: BU & Planner Analysis
- **Stacked bar**: Requests per BU (Total vs Reuse)
- **Bar chart**: Planner workload (Total vs Open)
- **Table**: Top 10 BUs with reuse rate

### Page 3: Pipeline Health
- **Bar**: Open request aging buckets
- **Gauge**: Reuse Rate %
- **Map** (if site GPS data available): Site demand heatmap
- **Table**: All open/pending requests

### Page 4: Site Analysis
- **Bar**: Top destination sites
- **Bar**: Top origin sites (lab returns)
- **Scatter**: SqFt Requested vs Assigned

---

## STEP 6 — Add Slicers (Interactive Filters)
Add these slicers to every page:
- **Year** (from DateRequested) — use Range slider slicer
- **BU** — Dropdown or Tile slicer
- **Status** — Checkbox list
- **Planner** — Dropdown
- **Site (To)** — Dropdown
- **Touch Type** — Dropdown (High/Low Touch)

To sync slicers across pages:
→ View → Sync Slicers → check all pages for each slicer

---

## STEP 7 — Publish & Embed in SharePoint

1. **Publish to Power BI Service**:
   → Home → Publish → Select your workspace

2. **Embed in SharePoint Page**:
   - Open your SharePoint page → Edit
   - Add Web Part → **Power BI**
   - Paste your report URL
   - Save the page

3. **Set Automatic Refresh**:
   → In Power BI Service → Datasets → Schedule Refresh
   → Connect with your Intel credentials
   → Set refresh to: Daily or every 4 hours

---

## STEP 8 — Add Historical Excel Data as Second Source

To include the historical data (2014-2025) from the Excel file:
1. Get Data → Excel Workbook → select "LET Vetting Working File.xlsm"
2. Select the "LET DOWNLOAD" sheet
3. In Power Query: append this query to the SharePoint query
4. Remove duplicates by RequestID
5. This gives you the FULL history in one dataset

---

## Cost Avoidance Validation

The $484M / 892K sqft figure from your OKR report uses:
- **SqFtAssigned** (col 55 "Actual sqft" in Excel) — the formally recorded assigned sqft
- Not SqFtRequested — the requested amount may include ranges, text, estimates
- In Power BI, create two measures: one for Requested, one for Assigned, and compare

To see which requests contribute to the $484M:
→ Filter Status = "Closed: Filled with existing lab space"
→ Sum the SqFtAssigned column
→ Multiply by $692

---

## Key Power BI Features You Get That Excel Doesn't Have
- ✅ **Live auto-refresh** from SharePoint (no manual export)
- ✅ **Cross-filter**: click any chart → all other charts filter instantly
- ✅ **Q&A**: type "show me reuse requests in JF by BU" in natural language
- ✅ **Mobile view**: auto-generates mobile-optimized layout
- ✅ **Alerts**: set threshold alerts (e.g., notify if open requests > 20)
- ✅ **Row-level security**: restrict planners to see only their assignments
- ✅ **Embed in Teams**: add as a Teams tab for your team
- ✅ **Export to PDF/PPT**: one-click executive slide export
