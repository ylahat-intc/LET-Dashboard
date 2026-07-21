"""
LET Lab Space Request — Interactive Dashboard Generator v2
Key features:
 - TRUE multi-select filters (checkboxes) for Status, BU, Planner, Site, Region, Trade
 - Month-range date filter + Q1/Q2/Q3/Q4 quarter toggle buttons
 - Quarter filter tied to the date branch that admitted each row (bug-fixed 2026-07-02)
 - CRESD-based campus normalization (937 codes, 5-tier fallback)
 - Trade Net Zero vs Growth tracking (LETClassification field, tracked since 2025-12)
 - Build-time QA sanity checks — surfaced in orange/red banner in the dashboard
 - Per-row Cost Avoided column in table (Reuse rows only)
 - Inline data fallback so HTML works when opened as local file (file:// URL)

VERSION: 2026-07-06
"""

import json, os, re
from datetime import datetime, timedelta

COST_PER_SQFT = 692
OUTPUT_FILE   = os.path.join(os.path.dirname(__file__), "LET_Dashboard.html")
JSON_CACHE    = os.path.join(os.path.dirname(__file__), "let_data_cache.json")
OLE_BASE      = datetime(1899, 12, 30)

BU_CLEAN = {
    "DCG (Data Center Group)": "DCG",
    "Silicon & Platform Engineering Group (SPE)": "SPE",
    "Client Computing Group-CCG": "CCG",
    "CTO&AI/IATG": "CTO&AI",
    "SIG VE grp": "SIG VE", "SIG MPE grp": "SIG MPE", "SIG WI grp": "SIG WI",
    "TBD - needs update": "TBD", "ICG (Intel China Group)": "ICG",
}

STATUS_NORM = {
    "closed: filled with existing lab space":       "Reuse",
    "closed: lab return back to cs":                "Lab Return",
    "closed: lab return back to rews":              "Lab Return",
    "closed: lab return back to cs/rews":           "Lab Return",
    "closed: vetted, require cre":                  "Capital/CRE",
    "closed: lswg capital req. - vetted approved":  "Capital/CRE",
    "closed: mrc approved.  sent to lnc":           "Capital/CRE",
    "closed : withdrawn":                           "Withdrawn",
    "closed: withdrawn":                            "Withdrawn",
    "closed: denied":                               "Denied/Unsupportable",
    "closed: unsupportable":                        "Denied/Unsupportable",
    "closed: new space":                            "Other Closed",
    "closed: placeholder":                          "Other Closed",
    "pending mrc acceptance":                       "Open/Pending",
    "hold":                                         "Hold",
    "open":                                         "Open/Pending",
    "new request":                                  "Open/Pending",
}

STATUS_COLOR = {
    "Reuse":                "#0068B5",
    "Lab Return":           "#00C7FD",
    "Capital/CRE":          "#FF8C00",
    "Withdrawn":            "#9CA3AF",
    "Denied/Unsupportable": "#CC3333",
    "Open/Pending":         "#00A86B",
    "Hold":                 "#F59E0B",
    "Other Closed":         "#AAAAAA",
}

def norm_status(s):
    if not s: return "Unknown"
    return STATUS_NORM.get(s.strip().lower(), "Other Closed")

def clean_bu(b):
    if not b: return "Unknown"
    return BU_CLEAN.get(b.strip(), b.strip())

def parse_sqft(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    m = re.search(r'-?\d[\d,]*', str(v))
    if m:
        try: return int(m.group().replace(",", ""))
        except: pass
    return 0

def ole_to_date(v):
    if v is None: return None
    # OLE numeric date (from Excel)
    if isinstance(v, (int, float)) and v > 1000:
        try: return (OLE_BASE + timedelta(days=float(v))).strftime("%Y-%m-%d")
        except: return None
    if isinstance(v, str):
        v = v.strip()
        if not v: return None
        # SharePoint ISO 8601: "2026-01-29T08:00:00Z" — just take the date part
        if 'T' in v:
            try: return datetime.strptime(v[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
            except: return None
        # US date format: "01/29/2026"
        try: return datetime.strptime(v, "%m/%d/%Y").strftime("%Y-%m-%d")
        except: pass
        # ISO date only: "2026-01-29"
        try: return datetime.strptime(v[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except: pass
    return None

def extract_site(s):
    if not s: return ""
    m = re.search(r'\b([A-Z]{2,4}\d*)\b', s.upper())
    return m.group(1) if m else s.strip()[:10]

# ── CRESD campus normalization ──────────────────────────────────────────────
_JUNK_SITES = {"NONE", "NA", "N/A", "TBD", "0", "", "UNKNOWN", "OTHER"}

def _load_cresd():
    path = os.path.join(os.path.dirname(__file__), "cresd_campus_map.json")
    if not os.path.exists(path):
        return {}, {}
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    items = data if isinstance(data, list) else [
        {k: data[k][i] for k in data} for i in range(len(next(iter(data.values()))))
    ]
    code_map, city_map = {}, {}
    for entry in items:
        code = (entry.get("code") or "").strip().upper()
        if code:
            code_map[code] = entry
        city = (entry.get("city") or "").strip().upper()
        if city:
            city_map.setdefault(city, []).append(code)
    return code_map, city_map

_CRESD, _CITY_TO_CODE = _load_cresd()
if _CRESD:
    print(f"  CRESD master loaded: {len(_CRESD)} campus codes")

_MANUAL_SITE_CODE = {
    "IS": "IDC", "ISR": "IDC", "ISRAEL": "IDC",
    "IDC1": "IDC", "IDC2": "IDC", "IDC3": "IDC",
    "ITK": "ITK",
    "JER": "JM", "JER1": "JM", "JER2": "JM",
    "RNB": "RNB", "JFCC": "JF", "JFS1": "JF",
    "SAN": "SC", "ISH": "YH",
}
_MANUAL_REGION = {
    "IDC": ("EMEA",  "Israel Development Center"),
    "RNB": ("EMEA",  "Raanana, Israel"),
    "JM":  ("EMEA",  "Jerusalem, Israel"),
    "ITK": ("APAC",  "Tsukuba, Japan"),
    "YH":  ("APAC",  "Yokohama, Japan"),
}

def canonical_site(raw: str) -> str:
    if not raw: return raw
    s = raw.strip().upper()
    if s in _JUNK_SITES: return s
    if s in _MANUAL_SITE_CODE: return _MANUAL_SITE_CODE[s]
    if s in _CRESD: return s
    letters = re.sub(r'[^A-Z]', '', s)
    if letters and letters in _CRESD: return letters
    stripped = re.sub(r'\d+$', '', letters)
    if stripped and stripped in _CRESD: return stripped
    if letters in _CITY_TO_CODE and len(_CITY_TO_CODE[letters]) == 1:
        return _CITY_TO_CODE[letters][0]
    for code, entry in _CRESD.items():
        city_up = (entry.get("city") or "").upper().replace(" ", "")
        if city_up and len(city_up) >= 4 and (
                city_up.startswith(letters) or letters.startswith(city_up)):
            return code
    return s

def resolve_region(site_code: str):
    if not site_code or site_code in _JUNK_SITES:
        return ("Unknown", site_code or "", "")
    if site_code in _MANUAL_REGION:
        reg, name = _MANUAL_REGION[site_code]
        return (reg, site_code, name)
    entry = _CRESD.get(site_code)
    if entry:
        reg = entry.get("region") or "Unknown"
        parts = [entry.get("city") or "", entry.get("country") or ""]
        name = ", ".join(p for p in parts if p)
        return (reg, site_code, name)
    return ("Unknown", site_code, "")

EXCEL_CACHE = os.path.join(os.path.dirname(__file__), "let_data_cache_excel.json")

def load_and_clean():
    # --- Load SharePoint (primary source) ---
    with open(JSON_CACHE, encoding="utf-8-sig") as f:
        raw = json.load(f)
    sp_items = raw if isinstance(raw, list) else raw.get("items", [])

    # --- Load Excel history (1001-1473 only, no duplicates) ---
    excel_items = []
    if os.path.exists(EXCEL_CACHE):
        with open(EXCEL_CACHE, encoding="utf-8-sig") as f:
            ex_raw = json.load(f)
        all_excel = ex_raw if isinstance(ex_raw, list) else ex_raw.get("items", [])
        for item in all_excel:
            fld = item.get("fields", item)
            rid = fld.get("RequestID") or fld.get("Title")
            try:
                if int(str(rid).strip()) < 1474:
                    excel_items.append(item)
            except: pass
        print(f"  Excel history loaded: {len(excel_items)} records (IDs 1001-1473)")

    all_items = excel_items + sp_items   # Excel first (oldest), then SP
    rows = []
    for item in all_items:
        fld = item.get("fields", item)
        item_id = item.get("id")  # Get ID from the item itself, not from fields
        
        def g(*keys):
            for k in keys:
                v = fld.get(k)
                if v not in (None, "", 0, False): return v
            return None

        date_req_str   = ole_to_date(g("DateRequested", "field_2"))
        date_close_str = ole_to_date(g("DateClosed", "field_29", "field_28"))

        year = quarter = ym = None
        if date_req_str:
            try:
                dt = datetime.strptime(date_req_str, "%Y-%m-%d")
                year    = dt.year
                quarter = (dt.month - 1) // 3 + 1   # 1-4
                ym      = date_req_str[:7]            # "YYYY-MM"
            except: pass

        # Close-date year/quarter/ym (for cost avoidance charts)
        year_close = quarter_close = ym_close = None
        if date_close_str:
            try:
                dtc = datetime.strptime(date_close_str, "%Y-%m-%d")
                year_close    = dtc.year
                quarter_close = (dtc.month - 1) // 3 + 1
                ym_close      = date_close_str[:7]
            except: pass

        sqft_req = parse_sqft(g("SqFtRequested", "field_13"))
        # field_30 = workstations (small count), field_31 = SqFtAssigned (actual sqft)
        sqft_asn = parse_sqft(g("SqFtAssigned", "field_31"))
        sqft_ret_act = parse_sqft(g("ActualReturnSqft", "Actual_x0020_Return_x0020_sqft"))

        days_open = 0
        if date_req_str and date_close_str:
            try:
                d = (datetime.strptime(date_close_str, "%Y-%m-%d") -
                     datetime.strptime(date_req_str,   "%Y-%m-%d")).days
                if 0 <= d <= 1500: days_open = d
            except: pass
        if not days_open:
            raw_do = g("DaysOpen", "field_44")
            if raw_do:
                try: days_open = int(float(str(raw_do)))
                except: pass

        planner = str(g("Owner", "Planner", "field_26") or "").strip()
        # Normalize "Last, First" → "First Last"
        if "," in planner:
            parts = [p.strip() for p in planner.split(",", 1)]
            planner = (parts[1] + " " + parts[0]).strip() if len(parts) == 2 else planner
        planner = planner.strip().title() if planner else "Unassigned"

        status_raw = str(g("Status", "field_3") or "")
        status     = norm_status(status_raw)
        is_open    = status in ("Open/Pending", "Hold")
        cpa_val    = str(g("CPA", "field_10") or "").strip().upper()
        cpa        = cpa_val in ("TRUE", "YES", "1", "Y")

        # Normalize site codes via CRESD master
        site_to_raw   = canonical_site(extract_site(str(g("SiteTo",   "field_20") or "")))
        site_from_raw = canonical_site(extract_site(str(g("SiteFrom", "field_21") or "")))
        region, site_code, site_name = resolve_region(site_to_raw)

        # Trade classification from LETClassification field (added Dec 2025)
        cls_raw = str(g("LETClassification") or "").strip()
        cls_l   = cls_raw.lower()
        if "net zero" in cls_l or ("trade" in cls_l and "zero" in cls_l):
            trade_type = "Net Zero"
        elif "growth" in cls_l:
            trade_type = "Growth"
        elif cls_raw and "defined" in cls_l:
            trade_type = "Undefined"
        else:
            trade_type = ""

        rows.append({
            "id":           str(g("RequestID", "Title") or ""),
            "title":        str(g("Title", "field_27") or "")[:80],
            "status_raw":   status_raw,
            "status":       status,
            "is_open":      is_open,
            "bu":           clean_bu(str(g("SuperGroup", "field_7") or "")),
            "group":        str(g("Group", "field_8") or "").strip()[:30],
            "sqft_req":     sqft_req,
            "sqft_asn":     sqft_asn,
            "sqft_ret_act": sqft_ret_act,
            "planner":      planner,
            "touch":        str(g("Touch", "field_17") or "").strip(),
            "permanent":    str(g("Permanent", "field_9") or "").strip(),
            "cpa":          cpa,
            "site_to":      site_to_raw,
            "site_from":    site_from_raw,
            "region":       region,
            "site_code":    site_code,
            "site_name":    site_name,
            "trade_type":   trade_type,
            "trade_raw":    cls_raw,
            "date_req":     date_req_str   or "",
            "date_close":   date_close_str or "",
            "year":         year,
            "quarter":      quarter,
            "ym":           ym or "",
            "year_close":   year_close,
            "quarter_close":quarter_close,
            "ym_close":     ym_close or "",
            "days_open":    days_open,
            "web_url":      f"https://intel.sharepoint.com/sites/GlobalLabsandDataCenters/Lists/LET2%20DLD%20Master%20Excel1/DispForm.aspx?ID={item_id}" if item_id else "",
        })
    return rows


def generate_html(rows, generated_at, qa_issues=None):
    rows_js      = json.dumps(rows, ensure_ascii=False)
    color_map_js = json.dumps(STATUS_COLOR)

    all_bus      = sorted({r["bu"]      for r in rows if r["bu"]      not in ("Unknown","")})
    all_planners = sorted({r["planner"] for r in rows if r["planner"] not in ("Unassigned","")})
    all_sites    = sorted({r["site_to"] for r in rows
                           if r["site_to"] and len(r["site_to"]) >= 2
                           and r["site_to"] not in ("NA","N/A","0","TBD")})
    all_touches  = sorted({r["touch"]   for r in rows if r["touch"]})
    all_regions  = sorted({r.get("region","Unknown") for r in rows if r.get("region")})
    all_statuses = sorted(STATUS_COLOR.keys())

    all_bus_js      = json.dumps(all_bus)
    all_planners_js = json.dumps(all_planners)
    all_sites_js    = json.dumps(all_sites)

    all_yms  = sorted({r["ym"] for r in rows if r["ym"]})
    min_ym   = all_yms[0]  if all_yms else "2015-01"
    max_ym   = all_yms[-1] if all_yms else "2026-12"
    all_years = sorted({r["year"] for r in rows if r["year"]})
    min_year  = all_years[0]  if all_years else 2015
    max_year  = all_years[-1] if all_years else 2026
    MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    def year_opts(selected):
        return "\n".join(
            '<option value="{y}"{sel}>{y}</option>'.format(
                y=y, sel=' selected' if y == selected else '')
            for y in all_years)

    def month_opts(selected_num):
        return "\n".join(
            '<option value="{n:02d}"{sel}>{m}</option>'.format(
                n=i+1, m=MONTHS[i], sel=' selected' if i+1 == selected_num else '')
            for i in range(12))

    def chk_list(items, css_class, show_search=False):
        """Build a searchable checkbox list."""
        search_id = f"srch-{css_class}"
        search_html = ""
        if show_search:
            search_html = (f'<input type="text" id="{search_id}" placeholder="Search…" '
                           f'oninput="filterSearch(\'{search_id}\',\'{css_class}\')" '
                           f'style="width:100%;padding:5px 8px;border:1px solid #E5E7EB;'
                           f'border-radius:5px;font-size:11px;margin-bottom:5px">')
        lines = [f'<label><input type="checkbox" class="{css_class}" value="{v}" '
                 f'checked onchange="applyFilters()"><span>{v}</span></label>'
                 for v in items]
        return (f'<div class="fl-top"><button class="mini-btn" '
                f'onclick="selectAll(\'{css_class}\',true)">All</button>'
                f'<button class="mini-btn" onclick="selectAll(\'{css_class}\',false)">None</button>'
                f'</div>'
                + search_html
                + '<div class="checkbox-group">' + "\n".join(lines) + '</div>')

    status_chks = "\n".join(
        '<label><input type="checkbox" class="f-status" value="{s}" checked '
        'onchange="applyFilters()">'
        '<span class="dot" style="background:{col}"></span>'
        '{s}</label>'.format(s=s, col=STATUS_COLOR.get(s, "#999"))
        for s in all_statuses)

    touch_chks   = chk_list(all_touches,  "f-touch")
    region_chks  = chk_list(all_regions,  "f-region")
    bu_chks      = chk_list(all_bus,      "f-bu",      show_search=True)
    planner_chks = chk_list(all_planners, "f-planner", show_search=True)
    site_chks    = chk_list(all_sites,    "f-site",    show_search=True)

    # Trade classification filter (only if any classified records exist)
    TRADE_LABELS = ["Net Zero", "Growth", "Undefined"]
    TRADE_COLORS = {"Net Zero": "#0A7C3A", "Growth": "#FF8C00", "Undefined": "#9CA3AF"}
    trade_dates = [r["date_req"] for r in rows if r.get("trade_type") and r.get("date_req")]
    trade_start = min(trade_dates) if trade_dates else ""
    trade_count = sum(1 for r in rows if r.get("trade_type"))
    trade_chks = ""
    if trade_count:
        trade_chks = (
            f'<div class="fl-top"><button class="mini-btn" onclick="selectAll(\'f-trade\',true)">All</button>'
            f'<button class="mini-btn" onclick="selectAll(\'f-trade\',false)">None</button></div>'
            f'<div class="checkbox-group">'
        ) + "\n".join(
            f'<label><input type="checkbox" class="f-trade" value="{t}" checked '
            f'onchange="applyFilters()"><span class="dot" style="background:{TRADE_COLORS[t]}"></span>{t}</label>'
            for t in TRADE_LABELS
        ) + (
            f'<label><input type="checkbox" class="f-trade" value="" checked '
            f'onchange="applyFilters()"><span class="dot" style="background:#E5E7EB"></span>'
            f'<span style="color:var(--gray)">Not Classified (legacy)</span></label>'
            f'</div>'
        )
    trade_colors_js = json.dumps(TRADE_COLORS)
    trade_section = ""
    if trade_count:
        trade_section = f"""
  <div class="section-title">🌱 Trade Net Zero vs Growth <span style="font-size:11px;font-weight:400;color:var(--gray)">(new workflow, tracked since {trade_start} · {trade_count} classified total · chart reflects active filters)</span></div>
  <div class="chart-grid g2">
    <div class="card"><div class="card-title">🌱 Trade Classification Split</div><div class="chart-wrap h240"><canvas id="c-trade"></canvas></div></div>
    <div class="card"><div class="card-title">📊 SqFt Requested by Trade Type</div><div class="chart-wrap h240"><canvas id="c-trade-sqft"></canvas></div></div>
  </div>"""
    trade_filter_section = f"""
  <div class="fs">
    <span class="fs-label">🌱 Trade Classification</span>
    {trade_chks if trade_chks else '<span class="note">No classified requests yet</span>'}
    {f'<div class="note">Tracked since {trade_start}</div>' if trade_start else ''}
  </div>""" if trade_count else ""

    # QA banner
    qa_issues = qa_issues or []
    qa_errors = [(l,m) for l,m in qa_issues if l=="ERROR"]
    qa_warns  = [(l,m) for l,m in qa_issues if l=="WARN"]
    if qa_issues:
        qa_color = "#CC3333" if qa_errors else "#F59E0B"
        qa_icon  = "🔴" if qa_errors else "⚠️"
        qa_label = (f"{len(qa_errors)} error(s)" if qa_errors else "") + \
                   (" · " if qa_errors and qa_warns else "") + \
                   (f"{len(qa_warns)} warning(s)" if qa_warns else "")
        qa_rows_html = "".join(
            f'<tr><td style="color:{"#CC3333" if lv=="ERROR" else "#F59E0B"};font-weight:600;'
            f'padding:2px 8px;white-space:nowrap">{lv}</td>'
            f'<td style="padding:2px 8px;font-size:12px">{msg}</td></tr>'
            for lv, msg in qa_issues[:50]
        )
        extra = f'<tr><td colspan="2" style="color:#6B7280;padding:2px 8px">… and {len(qa_issues)-50} more</td></tr>' if len(qa_issues) > 50 else ""
        qa_banner = f"""
  <div id="qa-banner" style="margin:8px 16px 0;border:1px solid {qa_color};border-radius:8px;overflow:hidden">
    <div onclick="document.getElementById('qa-detail').style.display=document.getElementById('qa-detail').style.display==='none'?'block':'none'"
         style="background:{qa_color}15;padding:8px 14px;cursor:pointer;display:flex;justify-content:space-between;align-items:center">
      <span style="font-weight:700;font-size:13px">{qa_icon} Data Quality: {qa_label} &nbsp;<span style="font-size:11px;opacity:.7">(click to expand)</span></span>
      <span style="font-size:11px;color:#6B7280">Built {generated_at}</span>
    </div>
    <div id="qa-detail" style="display:none;padding:8px 4px;overflow-x:auto">
      <table style="border-collapse:collapse;width:100%;font-family:monospace">{qa_rows_html}{extra}</table>
      <div style="padding:4px 8px;font-size:11px;color:#6B7280">Fix in SharePoint list or normalization rules in generate_dashboard.py</div>
    </div>
  </div>"""
    else:
        qa_banner = f'<div style="margin:6px 16px 0;padding:6px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;font-size:12px;color:#166534">✅ Data quality: all checks passed · Built {generated_at}</div>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>LET Lab Space Request — Live Dashboard</title>
<link rel="icon" type="image/png" href="favicon.png">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>
:root{{--blue:#0068B5;--lt:#00C7FD;--dark:#1A1A1A;--green:#00A86B;--orange:#FF8C00;
      --red:#CC3333;--gray:#6B7280;--bg:#F0F4FA;--card:#FFF;--border:#E5E7EB;--sw:290px}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--dark);display:flex;flex-direction:column;min-height:100vh}}
header{{background:linear-gradient(135deg,var(--blue) 0%,#004A82 100%);color:#fff;padding:14px 24px;
  display:flex;justify-content:space-between;align-items:center;box-shadow:0 2px 8px rgba(0,0,0,.2);
  position:sticky;top:0;z-index:100}}
header h1{{font-size:18px;font-weight:700}}
header .meta{{font-size:11px;opacity:.75;text-align:right;line-height:1.6}}
.layout{{display:flex;flex:1}}
/* SIDEBAR */
.sidebar{{width:var(--sw);background:var(--card);border-right:1px solid var(--border);
  padding:14px;overflow-y:auto;flex-shrink:0;position:sticky;top:56px;height:calc(100vh - 56px)}}
.sidebar h3{{font-size:12px;text-transform:uppercase;letter-spacing:.8px;color:var(--gray);
  margin-bottom:10px;display:flex;align-items:center;gap:6px}}
.fs{{margin-bottom:14px}}
.fs-label{{font-size:12px;font-weight:700;color:var(--dark);margin-bottom:5px;display:block}}
.checkbox-group{{max-height:130px;overflow-y:auto;border:1px solid var(--border);
  border-radius:6px;padding:7px;display:flex;flex-direction:column;gap:3px}}
.checkbox-group.tall{{max-height:180px}}
.checkbox-group label{{display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer;padding:1px 0;white-space:nowrap}}
.checkbox-group label:hover{{color:var(--blue)}}
.checkbox-group input[type=checkbox]{{cursor:pointer;flex-shrink:0}}
.dot{{display:inline-block;width:9px;height:9px;border-radius:50%;flex-shrink:0}}
.fl-top{{display:flex;gap:5px;margin-bottom:4px}}
.mini-btn{{padding:3px 8px;font-size:11px;border:1px solid var(--border);border-radius:4px;
  background:#fff;cursor:pointer;color:var(--blue);font-weight:600}}
.mini-btn:hover{{background:var(--blue);color:#fff}}
/* Month range */
.mrange{{display:flex;gap:6px;align-items:center;flex-wrap:wrap}}
.mrange select{{flex:1;padding:5px 7px;border:1px solid var(--border);
  border-radius:6px;font-size:12px;background:#fff;cursor:pointer}}
.mrange span{{font-size:11px;color:var(--gray)}}
/* Quarter buttons */
.q-btns{{display:flex;gap:4px;margin-top:6px}}
.q-btn{{flex:1;padding:5px 2px;font-size:11px;font-weight:700;border:2px solid var(--border);
  border-radius:5px;background:#fff;cursor:pointer;color:var(--gray);text-align:center}}
.q-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
/* SqFt toggle */
.sqft-toggle{{display:flex;flex-direction:column;gap:5px}}
.sqft-toggle label{{display:flex;align-items:center;gap:7px;font-size:12px;cursor:pointer}}
.sqft-note{{font-size:10px;color:var(--gray);font-style:italic;margin-top:2px}}
/* View preset bar */
.view-bar{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;padding:10px 14px;
  background:var(--card);border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);align-items:center}}
.view-bar-label{{font-size:11px;font-weight:700;color:var(--gray);text-transform:uppercase;
  letter-spacing:.5px;margin-right:4px}}
.view-btn{{padding:6px 14px;font-size:12px;font-weight:600;border:2px solid var(--border);
  border-radius:20px;background:#fff;cursor:pointer;color:var(--gray);transition:all .15s;
  white-space:nowrap;user-select:none}}
.view-btn:hover{{border-color:var(--blue);color:var(--blue)}}
.view-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.view-btn-note{{font-size:10px;color:var(--gray);font-style:italic;margin-left:4px}}
.btn-clear{{width:100%;padding:8px;background:var(--blue);color:#fff;border:none;border-radius:6px;
  font-size:13px;font-weight:600;cursor:pointer;margin-top:6px}}
.btn-clear:hover{{background:#004A82}}
.active-filters{{font-size:11px;color:var(--blue);font-weight:600;margin-top:8px;
  padding:6px 8px;background:#EFF6FF;border-radius:6px;display:none;line-height:1.8;
  word-break:break-word;white-space:normal}}
/* MAIN */
.main{{flex:1;padding:20px;overflow-y:auto;max-width:1400px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:20px}}
.kpi{{background:var(--card);border-radius:10px;padding:16px 18px;border-left:4px solid var(--blue);
  box-shadow:0 1px 4px rgba(0,0,0,.07);transition:transform .15s}}
.kpi:hover{{transform:translateY(-2px)}}
.kpi.green{{border-color:var(--green)}}.kpi.orange{{border-color:var(--orange)}}.kpi.red{{border-color:var(--red)}}.kpi.lt{{border-color:var(--lt)}}
.kpi-label{{font-size:11px;color:var(--gray);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}}
.kpi-value{{font-size:26px;font-weight:700}}
.kpi-sub{{font-size:11px;color:var(--gray);margin-top:3px}}
.section-title{{font-size:14px;font-weight:700;color:var(--blue);margin-bottom:12px;
  display:flex;align-items:center;gap:8px;border-bottom:2px solid #E3EEFF;padding-bottom:6px}}
.chart-grid{{display:grid;gap:16px;margin-bottom:16px}}
.g2{{grid-template-columns:1fr 1fr}}.g3{{grid-template-columns:1fr 1fr 1fr}}
.g13{{grid-template-columns:2fr 1fr}}.g31{{grid-template-columns:1fr 2fr}}
@media(max-width:900px){{.g2,.g3,.g13,.g31{{grid-template-columns:1fr}}}}
.card{{background:var(--card);border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:0}}
.card-title{{font-size:12px;font-weight:600;color:var(--gray);text-transform:uppercase;letter-spacing:.4px;margin-bottom:12px}}
.chart-wrap{{position:relative}}.h240{{height:240px}}.h280{{height:280px}}.h300{{height:300px}}.h320{{height:320px}}
.tbl-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:var(--blue);color:#fff;padding:9px 12px;text-align:left;font-size:11px;
  font-weight:600;cursor:pointer;user-select:none;white-space:nowrap}}
th:hover{{background:#004A82}}
td{{padding:8px 12px;border-bottom:1px solid var(--border);vertical-align:middle}}
tr:hover td{{background:#F0F7FF}}
.badge{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700}}
.tbl-footer{{font-size:11px;color:var(--gray);padding:8px 0;text-align:right}}
.insight{{background:#EFF6FF;border-left:4px solid var(--blue);padding:10px 14px;
  border-radius:0 8px 8px 0;font-size:13px;margin-bottom:16px}}
.insight strong{{color:var(--blue)}}
.note{{font-size:11px;color:var(--gray);font-style:italic;margin-top:4px}}
/* Table filter bar */
.tbl-filter-bar{{display:flex;flex-wrap:wrap;gap:10px 12px;align-items:center;padding:12px;background:#F9FAFB;border:1px solid var(--border);border-radius:8px;margin-bottom:12px}}
.tbl-filter-bar input[type=text]{{flex:1;min-width:200px;padding:6px 10px;border:1px solid var(--border);border-radius:5px;font-size:12px}}
.tbl-filter-bar input[type=month],.tbl-filter-bar input[type=number]{{padding:5px 7px;border:1px solid var(--border);border-radius:5px;font-size:12px}}
.tf-group{{display:flex;gap:6px;align-items:center}}
.tf-label{{font-size:11px;font-weight:600;color:var(--gray);min-width:50px;text-align:right}}
.tf-multi{{position:relative;display:inline-block}}
.tf-multi-btn{{padding:5px 10px;border:1px solid var(--border);border-radius:5px;background:#fff;font-size:12px;cursor:pointer;min-width:100px;text-align:left;font-weight:600;color:var(--blue)}}
.tf-multi-btn:hover{{border-color:var(--blue)}}
.tf-multi-pop{{position:absolute;top:100%;left:0;background:#fff;border:1px solid var(--border);border-radius:6px;margin-top:4px;min-width:180px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:1000;max-height:200px;overflow-y:auto;display:none;padding:8px}}
.tf-multi-pop.open{{display:block}}
.tf-status-cb{{display:flex;align-items:center;gap:6px;padding:4px 0;font-size:12px;cursor:pointer;width:100%}}
.sp-link{{color:var(--blue);text-decoration:none;cursor:pointer;font-weight:500}}
.sp-link:hover{{text-decoration:underline}}
</style>
</head>
<body>

<header>
  <div style="display:flex;align-items:center;gap:14px">
    <img src="data:image/png;base64,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" alt="LET" style="height:36px">
    <div>
      <h1>LET Lab Space Request — Live Dashboard</h1>
      <div style="font-size:11px;opacity:.8;margin-top:2px">Global Lab Planner · Intel Lab Efficiency Team · All filters apply simultaneously to all charts</div>
    </div>
  </div>
  <div class="meta">Refreshed: {generated_at}<br><span id="hdr-count"></span></div>
</header>

<div class="layout">

<!-- ═══ FILTER SIDEBAR ═══ -->
<aside class="sidebar">
  <h3>🎛️ Filters</h3>
  <div id="active-filter-badge" class="active-filters"></div>

  <!-- DATE / QUARTER -->
  <div class="fs">
    <span class="fs-label">📅 From</span>
    <div class="mrange">
      <select id="f-year-from" onchange="applyFilters()">{year_opts(min_year)}</select>
      <select id="f-month-from" onchange="applyFilters()">{month_opts(1)}</select>
    </div>
    <span class="fs-label" style="margin-top:6px">To</span>
    <div class="mrange">
      <select id="f-year-to" onchange="applyFilters()">{year_opts(max_year)}</select>
      <select id="f-month-to" onchange="applyFilters()">{month_opts(12)}</select>
    </div>
    <div class="q-btns">
      <button class="q-btn active" data-q="1" onclick="toggleQ(this)">Q1<br><span style="font-size:9px;font-weight:400">Jan–Mar</span></button>
      <button class="q-btn active" data-q="2" onclick="toggleQ(this)">Q2<br><span style="font-size:9px;font-weight:400">Apr–Jun</span></button>
      <button class="q-btn active" data-q="3" onclick="toggleQ(this)">Q3<br><span style="font-size:9px;font-weight:400">Jul–Sep</span></button>
      <button class="q-btn active" data-q="4" onclick="toggleQ(this)">Q4<br><span style="font-size:9px;font-weight:400">Oct–Dec</span></button>
    </div>
    <div class="note" style="margin-top:4px">Click quarter to toggle on/off</div>
  </div>

  <!-- STATUS -->
  <div class="fs">
    <span class="fs-label">📊 Status</span>
    <div class="fl-top">
      <button class="mini-btn" onclick="selectAll('f-status',true)">All</button>
      <button class="mini-btn" onclick="selectAll('f-status',false)">None</button>
    </div>
    <div class="checkbox-group">
      {status_chks}
    </div>
  </div>

  <!-- BU -->
  <div class="fs">
    <span class="fs-label">🏢 Business Unit</span>
    {bu_chks}
  </div>

  <!-- PLANNER -->
  <div class="fs">
    <span class="fs-label">👤 Planner</span>
    {planner_chks}
  </div>

  <!-- REGION -->
  <div class="fs">
    <span class="fs-label">🌍 Region</span>
    {region_chks}
  </div>

  <!-- SITE TO -->
  <div class="fs">
    <span class="fs-label">📍 Site (Destination)</span>
    {site_chks}
  </div>

  <!-- TOUCH -->
  <div class="fs">
    <span class="fs-label">⚡ Touch Type</span>
    {touch_chks}
  </div>

  {trade_filter_section}

  <!-- SQFT NOTE -->
  <div class="fs">
    <span class="fs-label">📐 Cost Avoidance Basis</span>
    <div class="note" style="background:rgba(0,104,181,.06);padding:7px 9px;border-radius:10px;border-left:3px solid var(--accent);font-style:normal;color:var(--text-dim)">
      <b>Auto rule:</b><br>
      • Pre-2022 → SqFt <b>Requested</b><br>
      • 2022 onward → SqFt <b>Assigned</b><br>
      (falls back to Requested if Assigned missing)<br>
      • Lab Return → use <b>Actual Return sqft</b> (falls back to Requested)
    </div>
  </div>

  <!-- CPA -->
  <div class="fs">
    <span class="fs-label" title="CPA = Critical Path Activity: marks requests on critical project timelines">🔴 Critical Path Activity (CPA)</span>
    <div class="fl-top">
      <button class="mini-btn" onclick="setCPA('')" id="cpa-all" style="background:var(--blue);color:#fff">All</button>
      <button class="mini-btn" onclick="setCPA('yes')" id="cpa-yes">CPA Only</button>
      <button class="mini-btn" onclick="setCPA('no')"  id="cpa-no">Non-CPA</button>
    </div>
    <div class="note">Only 32 of {len(rows)} records are CPA-flagged.<br>CPA = request blocking a critical program timeline.</div>
  </div>

  <button class="btn-clear" onclick="clearFilters()">↺ Clear All Filters</button>
</aside>

<!-- ═══ MAIN CONTENT ═══ -->
<main class="main">

  <!-- VIEW PRESET BAR — multi-select, click to toggle -->
  <div class="view-bar">
    <span class="view-bar-label">📋 View:</span>
    <button class="view-btn active" id="vb-all"    onclick="toggleView('all',this)">All Requests</button>
    <button class="view-btn"        id="vb-open"   onclick="toggleView('open',this)">🔄 Open / Active <span id="vb-open-count" style="font-size:10px;opacity:.7"></span></button>
    <button class="view-btn"        id="vb-closed" onclick="toggleView('closed',this)">✅ Closed (All)</button>
    <button class="view-btn"        id="vb-reuse"  onclick="toggleView('reuse',this)">♻️ Reuse</button>
    <button class="view-btn"        id="vb-return" onclick="toggleView('return',this)">🏭 Lab Return</button>
    <button class="view-btn"        id="vb-capital" onclick="toggleView('capital',this)">🏗️ Capital/CRE</button>
    <button class="view-btn"        id="vb-withdrawn" onclick="toggleView('withdrawn',this)">↩️ Withdrawn</button>
    <span class="view-btn-note" id="view-desc">Click to toggle · multiple allowed</span>
  </div>

  {qa_banner}
  <div class="insight" id="insight-bar">Loading...</div>

  <div class="section-title">📊 Executive KPIs <span style="font-size:11px;font-weight:400;color:var(--gray)">(all update with filters)</span></div>
  <div class="kpi-grid">
    <div class="kpi green"><div class="kpi-label">💰 Cost Avoided</div><div class="kpi-value" id="kpi-cost">—</div><div class="kpi-sub" id="kpi-cost-sub">Reuse × ${COST_PER_SQFT}/sqft</div></div>
    <div class="kpi"><div class="kpi-label">♻️ Reuse Rate</div><div class="kpi-value" id="kpi-reuse-rate">—</div><div class="kpi-sub" id="kpi-reuse-sub">of closed</div></div>
    <div class="kpi lt"><div class="kpi-label">📐 SqFt (Reuse)</div><div class="kpi-value" id="kpi-sqft">—</div><div class="kpi-sub" id="kpi-sqft-sub">—</div></div>
    <div class="kpi orange"><div class="kpi-label">🏭 SqFt Returned</div><div class="kpi-value" id="kpi-returned">—</div><div class="kpi-sub">Freed to inventory</div></div>
    <div class="kpi"><div class="kpi-label">📋 Total Requests</div><div class="kpi-value" id="kpi-total">—</div><div class="kpi-sub" id="kpi-open-sub">—</div></div>
    <div class="kpi"><div class="kpi-label">⏱️ Avg Days/Close</div><div class="kpi-value" id="kpi-days">—</div><div class="kpi-sub">Submit → resolution</div></div>
    <div class="kpi red"><div class="kpi-label">⚡ High Touch</div><div class="kpi-value" id="kpi-touch">—</div><div class="kpi-sub">of filtered</div></div>
    <div class="kpi orange"><div class="kpi-label">🏗️ Capital/CRE</div><div class="kpi-value" id="kpi-capital">—</div><div class="kpi-sub">Escalated to capital</div></div>
  </div>

  <div class="section-title">💰 Cost Avoidance & Closure Analysis</div>
  <div class="chart-grid g2">
    <div class="card"><div class="card-title">📅 Cost Avoided by Year <span style="font-size:10px;color:var(--gray)">(by close date)</span></div><div class="chart-wrap h280"><canvas id="c-cost-year"></canvas></div></div>
    <div class="card"><div class="card-title">🍩 Closure Categories</div><div class="chart-wrap h280"><canvas id="c-status"></canvas></div></div>
  </div>

  <div class="section-title">📈 Request Volume Trends</div>
  <div class="chart-grid g2">
    <div class="card"><div class="card-title">📅 Monthly New (submitted) vs Closed</div><div class="chart-wrap h280"><canvas id="c-monthly"></canvas></div></div>
    <div class="card"><div class="card-title">📆 Quarterly Volume (by Year)</div><div class="chart-wrap h280"><canvas id="c-quarterly"></canvas></div></div>
  </div>

  <div class="section-title">🏢 BU & Planner Analysis</div>
  <div class="chart-grid g2">
    <div class="card"><div class="card-title">📊 Top BUs — Total vs Reuse</div><div class="chart-wrap h300"><canvas id="c-bu"></canvas></div></div>
    <div class="card"><div class="card-title">👤 Planner Workload — Open vs Total</div><div class="chart-wrap h300"><canvas id="c-planner"></canvas></div></div>
  </div>

  <div class="section-title">🚦 Pipeline Health</div>
  <div class="chart-grid g3">
    <div class="card"><div class="card-title">⏳ Open Request Aging</div><div class="chart-wrap h240"><canvas id="c-aging"></canvas></div></div>
    <div class="card"><div class="card-title">🔧 High vs Low Touch</div><div class="chart-wrap h240"><canvas id="c-touch"></canvas></div></div>
    <div class="card"><div class="card-title">📐 SqFt Buckets (Requested)</div><div class="chart-wrap h240"><canvas id="c-sqft-dist"></canvas></div></div>
  </div>

  <div class="section-title">🗺️ Site Demand</div>
  <div class="chart-grid g2">
    <div class="card"><div class="card-title">🎯 Top Destination Sites</div><div class="chart-wrap h240"><canvas id="c-site-to"></canvas></div></div>
    <div class="card"><div class="card-title">🏃 Top Origin Sites (Vacating)</div><div class="chart-wrap h240"><canvas id="c-site-from"></canvas></div></div>
  </div>

  {trade_section}

  <div class="section-title">📋 Request Detail Table <span style="font-size:11px;font-weight:400;color:var(--gray)">(click column header to sort)</span></div>
  <div class="card">
    <div class="tbl-filter-bar">
      <input type="text" id="tf-search" placeholder="🔍 Search ID, title, planner, BU…" oninput="renderTable(filteredData)">
      <div class="tf-group">
        <label class="tf-label">Status:</label>
        <div class="tf-multi" id="tf-status-wrap">
          <button class="tf-multi-btn" id="tf-status-btn" onclick="toggleStatusDropdown(event)">All ▾</button>
          <div class="tf-multi-pop" id="tf-status-pop"></div>
        </div>
      </div>
      <div class="tf-group">
        <label class="tf-label">Sub:</label>
        <input type="month" id="tf-sub-from" placeholder="From" oninput="renderTable(filteredData)">
        <input type="month" id="tf-sub-to" placeholder="To" oninput="renderTable(filteredData)">
      </div>
      <div class="tf-group">
        <label class="tf-label">Close:</label>
        <input type="month" id="tf-close-from" placeholder="From" oninput="renderTable(filteredData)">
        <input type="month" id="tf-close-to" placeholder="To" oninput="renderTable(filteredData)">
      </div>
      <div class="tf-group">
        <label class="tf-label">Days Open:</label>
        <input type="number" id="tf-days-min" placeholder="Min" min="0" style="width:70px" oninput="renderTable(filteredData)">
        <input type="number" id="tf-days-max" placeholder="Max" min="0" style="width:70px" oninput="renderTable(filteredData)">
      </div>
      <button style="padding:6px 14px;font-size:12px;background:var(--blue);color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:600" onclick="clearTableFilters()">↺ Reset</button>
    </div>
    <div class="tbl-wrap">
      <table id="detail-table">
        <thead><tr>
          <th onclick="sortTable('id')">ID ↕</th>
          <th onclick="sortTable('title')">Title ↕</th>
          <th onclick="sortTable('status')">Status ↕</th>
          <th onclick="sortTable('bu')">BU ↕</th>
          <th onclick="sortTable('region')">Region ↕</th>
          <th onclick="sortTable('site_to')">Site ↕</th>
          <th onclick="sortTable('trade_type')" title="Trade Net Zero / Growth classification">Trade ↕</th>
          <th onclick="sortTable('sqft_req')" style="text-align:right">SqFt Req ↕</th>
          <th onclick="sortTable('sqft_asn')" style="text-align:right" title="Used for cost calc 2022+">SqFt Asn ↕</th>
          <th onclick="sortTable('sqft_ret_act')" style="text-align:right" title="Used for Lab Return KPI (fallback to SqFt Req when missing)">SqFt Ret Act ↕</th>
          <th onclick="sortTable('cost_avoided')" style="text-align:right" title="Cost avoided (Reuse only): 2022+ uses Assigned sqft, pre-2022 uses Requested. Both × ${COST_PER_SQFT}/sqft">💰 Cost ↕</th>
          <th onclick="sortTable('planner')">Planner ↕</th>
          <th onclick="sortTable('ym')">Submitted ↕</th>
          <th onclick="sortTable('ym_close')">Closed ↕</th>
          <th onclick="sortTable('quarter')">Q ↕</th>
          <th onclick="sortTable('days_open')" style="text-align:right">Days ↕</th>
        </tr></thead>
        <tbody id="tbl-body"></tbody>
      </table>
    </div>
    <div class="tbl-footer" id="tbl-footer"></div>
  </div>

  <div style="text-align:center;padding:16px;font-size:11px;color:var(--gray)">
    Intel LET Lab Space Dashboard · {generated_at} · Data: SharePoint LET Lab Space Request/Release
  </div>
</main>
</div>

<script>
const COST_SQFT  = {COST_PER_SQFT};
const COLOR_MAP  = {color_map_js};
const TRADE_COLORS = {trade_colors_js};
const KNOWN_BUS      = new Set({all_bus_js});
const KNOWN_PLANNERS = new Set({all_planners_js});
const KNOWN_SITES    = new Set({all_sites_js});

const charts = {{}};
let sortKey = 'date_req', sortDir = -1;
let ALL_DATA = [];
let filteredData = [];
let cpaMode = '';          // '' | 'yes' | 'no'
let currentSqftMode = 'asn';

const fmt$ = v => {{
  if (v >= 1e9) return '$' + (v/1e9).toFixed(2) + 'B';
  if (v >= 1e6) return '$' + (v/1e6).toFixed(1) + 'M';
  return '$' + Math.round(v).toLocaleString();
}};
const fmtN = v => Math.round(v).toLocaleString();
const GRIDCOLOR = 'rgba(0,0,0,0.05)';
const FONT = {{ family:"'Segoe UI',system-ui,sans-serif", size:12 }};
Chart.defaults.font = FONT;
Chart.defaults.color = undefined;
Chart.defaults.plugins.legend.labels.boxWidth = 14;

function mkChart(id, cfg) {{
  const ctx = document.getElementById(id);
  if (!ctx) return;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(ctx, cfg);
}}
function counter(arr, key) {{
  const m = {{}};
  arr.forEach(r => {{ const v = r[key] || 'Unknown'; m[v] = (m[v]||0)+1; }});
  return m;
}}
function topN(obj, n=10) {{
  return Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,n);
}}
function parseNum(v) {{
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g,''));
  return isNaN(n) ? 0 : n;
}}

const BADGE_STYLE = {{
  'Reuse':               'background:#E3F2FD;color:#0068B5',
  'Lab Return':          'background:#E0F7FA;color:#007799',
  'Capital/CRE':         'background:#FFF3E0;color:#CC6600',
  'Withdrawn':           'background:#F3F4F6;color:#6B7280',
  'Denied/Unsupportable':'background:#FFEBEE;color:#CC3333',
  'Open/Pending':        'background:#E8F5E9;color:#00A86B',
  'Other Closed':        'background:#F9FAFB;color:#9CA3AF',
}};
function badge(s) {{
  return `<span class="badge" style="${{BADGE_STYLE[s]||''}}">${{s}}</span>`;
}}

// ── VIEW PRESETS — multi-select toggles ─────────────────
const VIEW_STATUSES = {{
  all:      null,   // null = no status filter (show everything)
  open:     ['Open/Pending'],
  closed:   ['Reuse','Lab Return','Capital/CRE','Withdrawn','Denied/Unsupportable','Other Closed'],
  reuse:    ['Reuse'],
  return:   ['Lab Return'],
  capital:  ['Capital/CRE'],
  withdrawn:['Withdrawn','Denied/Unsupportable'],
}};
let activeViews = new Set(['all']);

function toggleView(v, btn) {{
  if (v === 'all') {{
    // All = reset everything
    activeViews = new Set(['all']);
    document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }} else {{
    // Deactivate "All" when picking specifics
    activeViews.delete('all');
    document.getElementById('vb-all').classList.remove('active');
    if (activeViews.has(v)) {{
      activeViews.delete(v);
      btn.classList.remove('active');
      if (activeViews.size === 0) {{
        activeViews.add('all');
        document.getElementById('vb-all').classList.add('active');
      }}
    }} else {{
      activeViews.add(v);
      btn.classList.add('active');
    }}
  }}
  // Build combined status list from all active views
  let combined = null;
  if (!activeViews.has('all')) {{
    combined = [];
    activeViews.forEach(v => {{
      const s = VIEW_STATUSES[v];
      if (s) s.forEach(x => {{ if (!combined.includes(x)) combined.push(x); }});
    }});
  }}
  // Sync status checkboxes
  document.querySelectorAll('.f-status').forEach(cb => {{
    cb.checked = combined === null || combined.includes(cb.value);
  }});
  applyFilters();
}}

// ── FILTER HELPERS ────────────────────────────────────────
function selectAll(cls, state) {{
  document.querySelectorAll('.' + cls).forEach(cb => cb.checked = state);
  if (cls === 'f-status') {{
    activeViews = new Set(['all']);
    document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('vb-all').classList.add('active');
  }}
  applyFilters();
}}
function filterSearch(inputId, cls) {{
  const q = document.getElementById(inputId).value.toLowerCase();
  document.querySelectorAll('.' + cls).forEach(cb => {{
    const lbl = cb.closest('label');
    if (lbl) lbl.style.display = cb.value.toLowerCase().includes(q) ? '' : 'none';
  }});
}}
function toggleQ(btn) {{
  btn.classList.toggle('active');
  applyFilters();
}}
function setCPA(mode) {{
  cpaMode = mode;
  document.querySelectorAll('#cpa-all,#cpa-yes,#cpa-no').forEach(b=>b.style.background='');
  document.querySelectorAll('#cpa-all,#cpa-yes,#cpa-no').forEach(b=>b.style.color='');
  const active = document.getElementById('cpa-' + (mode||'all'));
  if (active) {{ active.style.background='var(--blue)'; active.style.color='#fff'; }}
  applyFilters();
}}

// ── MAIN FILTER LOGIC ─────────────────────────────────────
function getFilters() {{
  const yFrom  = parseInt(document.getElementById('f-year-from').value)  || {min_year};
  const mFrom  = parseInt(document.getElementById('f-month-from').value) || 1;
  const yTo    = parseInt(document.getElementById('f-year-to').value)    || {max_year};
  const mTo    = parseInt(document.getElementById('f-month-to').value)   || 12;
  const ymFrom = yFrom + '-' + String(mFrom).padStart(2,'0');
  const ymTo   = yTo   + '-' + String(mTo).padStart(2,'0');
  const quarters= [...document.querySelectorAll('.q-btn.active')].map(b=>parseInt(b.dataset.q));
  const statuses= [...document.querySelectorAll('.f-status:checked')].map(e=>e.value);
  const bus     = [...document.querySelectorAll('.f-bu:checked')].map(e=>e.value);
  const planners= [...document.querySelectorAll('.f-planner:checked')].map(e=>e.value);
  const sites   = [...document.querySelectorAll('.f-site:checked')].map(e=>e.value);
  const regions = [...document.querySelectorAll('.f-region:checked')].map(e=>e.value);
  const touches = [...document.querySelectorAll('.f-touch:checked')].map(e=>e.value);
  const tradeNodes = document.querySelectorAll('.f-trade');
  const trades  = tradeNodes.length ? [...document.querySelectorAll('.f-trade:checked')].map(e=>e.value) : null;
  currentSqftMode = 'auto';
  return {{ ymFrom, ymTo, quarters, statuses, bus, planners, sites, regions, touches, trades }};
}}

function applyFilters() {{
  const f = getFilters();
  const TODAY_YM = new Date().toISOString().slice(0,7);
  filteredData = ALL_DATA.filter(r => {{
    // DATE: include if submission OR close date falls in range; always include open/pending if range covers today
    const ymReq   = r.ym       || '';
    const ymClose = r.ym_close || '';
    let reqIn = false, closeIn = false, isOpenNow = false;
    if (ymReq || ymClose) {{
      reqIn     = !!(ymReq   && ymReq   >= f.ymFrom && ymReq   <= f.ymTo);
      closeIn   = !!(ymClose && ymClose >= f.ymFrom && ymClose <= f.ymTo);
      isOpenNow = !!(r.is_open && f.ymTo >= TODAY_YM);
      if (!reqIn && !closeIn && !isOpenNow) return false;
    }}
    // QUARTER: tie to the date branch that admitted the row (prevents 2025-Q3
    // request closed 2026-Q2 from appearing under "2026 Q3")
    if (f.quarters.length < 4) {{
      const qc = [];
      if (reqIn     && r.quarter)        qc.push(r.quarter);
      if (closeIn   && r.quarter_close)  qc.push(r.quarter_close);
      if (isOpenNow && r.quarter)        qc.push(r.quarter);
      if (!qc.length && r.quarter)       qc.push(r.quarter);
      if (qc.length && !qc.some(q => f.quarters.includes(q))) return false;
    }}
    if (f.statuses.length && !f.statuses.includes(r.status)) return false;
    // BU/Planner/Site: only filter rows whose value appears in the checkbox list.
    // Rows with values not in the list (blank, Unassigned, NA, etc.) bypass.
    if (f.bus.length     && KNOWN_BUS.has(r.bu)           && !f.bus.includes(r.bu))           return false;
    if (f.planners.length && KNOWN_PLANNERS.has(r.planner) && !f.planners.includes(r.planner)) return false;
    if (f.sites.length   && KNOWN_SITES.has(r.site_to)    && !f.sites.includes(r.site_to))    return false;
    if (f.regions.length  && !f.regions.includes(r.region || 'Unknown')) return false;
    if (f.touches.length && r.touch && !f.touches.includes(r.touch)) return false;
    if (f.trades && f.trades.length && !f.trades.includes(r.trade_type || '')) return false;
    if (cpaMode === 'yes' && !r.cpa) return false;
    if (cpaMode === 'no'  &&  r.cpa) return false;
    return true;
  }});

  // Active filter badge
  const active = [];
  if (f.ymFrom !== '{min_year}-01' || f.ymTo !== '{max_year}-12') active.push(`📅 ${{f.ymFrom.slice(0,7)}} → ${{f.ymTo.slice(0,7)}}`);
  if (f.quarters.length < 4) active.push('Q' + f.quarters.join(',Q'));
  if (f.statuses.length < {len(all_statuses)}) active.push(`${{f.statuses.length}} of {len(all_statuses)} statuses`);
  if (f.bus.length     < {len(all_bus)})      active.push(`${{f.bus.length}} of {len(all_bus)} BUs`);
  if (f.planners.length < {len(all_planners)}) active.push(`${{f.planners.length}} planners`);
  if (f.sites.length   < {len(all_sites)})    active.push(`${{f.sites.length}} sites`);
  if (f.regions.length < {len(all_regions)})  active.push(`${{f.regions.length}} regions`);
  if (f.touches.length < {len(all_touches)})  active.push(f.touches.join('/'));
  if (cpaMode) active.push('CPA:' + cpaMode);
  const el = document.getElementById('active-filter-badge');
  if (active.length) {{ el.style.display='block'; el.textContent='Active: '+active.join(' · '); }}
  else {{ el.style.display='none'; }}

  renderAll();
}}

function clearFilters() {{
  const _cy = new Date().getFullYear();
  document.getElementById('f-year-from').value  = _cy;
  document.getElementById('f-month-from').value = '01';
  document.getElementById('f-year-to').value    = _cy;
  document.getElementById('f-month-to').value   = '12';
  document.querySelectorAll('.q-btn').forEach(b=>b.classList.add('active'));
  document.querySelectorAll('.f-status,.f-bu,.f-planner,.f-site,.f-region,.f-touch,.f-trade').forEach(cb=>cb.checked=true);
  currentSqftMode = 'auto';
  setCPA('');
  applyFilters();
}}

// Pre-2022 → use sqft_req; 2022+ → use sqft_asn if available, else sqft_req
function sqftForCost(r) {{
  const yr = r.year_close || r.year || 0;
  if (yr >= 2022 && r.sqft_asn > 0) return r.sqft_asn;
  return r.sqft_req;
}}

function sqftForReturn(r) {{
  const actual = parseNum(r.sqft_ret_act);
  if (actual !== 0) return Math.abs(actual);
  return Math.abs(parseNum(r.sqft_req));
}}

// ── RENDER ALL ────────────────────────────────────────────
function renderAll() {{
  const d = filteredData;
  const n = d.length;
  const reuse   = d.filter(r => r.status === 'Reuse');
  const returns = d.filter(r => r.status === 'Lab Return');
  const capital = d.filter(r => r.status === 'Capital/CRE');
  const open    = d.filter(r => r.status === 'Open/Pending');
  const closed  = d.filter(r => r.status !== 'Open/Pending');

  const sqftModeLabel = 'Pre-2022: Requested · 2022+: Assigned';

  const sqftReused   = reuse.reduce((s,r)=>s+Math.max(sqftForCost(r),0), 0);
  const sqftReturned = returns.reduce((s,r)=>s+Math.max(sqftForReturn(r),0), 0);
  const costAvoided  = sqftReused * COST_SQFT;
  const reuseRate    = closed.length ? (reuse.length/closed.length*100).toFixed(1)+'%' : '0%';

  const validDays = closed.filter(r=>r.days_open>0&&r.days_open<1500);
  const avgDays   = validDays.length ? Math.round(validDays.reduce((s,r)=>s+r.days_open,0)/validDays.length) : 0;
  const hiTouch   = d.filter(r=>r.touch==='High Touch').length;

  document.getElementById('kpi-cost').textContent     = fmt$(costAvoided);
  document.getElementById('kpi-cost-sub').textContent = `${{sqftModeLabel}}`;
  document.getElementById('kpi-reuse-rate').textContent = reuseRate;
  document.getElementById('kpi-reuse-sub').textContent  = `${{reuse.length}} of ${{closed.length}} closed`;
  document.getElementById('kpi-sqft').textContent     = fmtN(sqftReused);
  document.getElementById('kpi-sqft-sub').textContent = sqftModeLabel;
  document.getElementById('kpi-returned').textContent = fmtN(sqftReturned);
  document.getElementById('kpi-total').textContent    = fmtN(n);
  document.getElementById('kpi-open-sub').textContent = `${{open.length}} open / pending`;
  document.getElementById('kpi-days').textContent     = avgDays || '—';
  document.getElementById('kpi-touch').textContent    = n ? Math.round(100*hiTouch/n)+'%' : '0%';
  document.getElementById('kpi-capital').textContent  = capital.length;
  document.getElementById('hdr-count').textContent    = `${{fmtN(n)}} of ${{fmtN(ALL_DATA.length)}} records`;
  const openBadge = document.getElementById('vb-open-count');
  if (openBadge) openBadge.textContent = `(${{ALL_DATA.filter(r=>r.is_open).length}})`;

  document.getElementById('insight-bar').innerHTML =
    `<strong>💡 Key Insight:</strong> ${{reuse.length}} requests resolved via reuse (${{reuseRate}} rate),
    saving <strong>${{fmt$(costAvoided)}}</strong> (${{fmtN(sqftReused)}} ${{sqftModeLabel}} × $${{COST_SQFT}}/sqft).
    ${{returns.length}} returns freed <strong>${{fmtN(sqftReturned)}} sqft</strong>.
    ${{open.length}} requests open/pending. ${{capital.length}} escalated to Capital/CRE.`;

  renderCostByYear(d, reuse);
  renderStatusDonut(d);
  renderMonthly(d);
  renderQuarterly(d);
  renderBU(d, reuse);
  renderPlanner(d, open);
  renderAging(open);
  renderTouch(d);
  renderSqftDist(d);
  renderSiteTo(d);
  renderSiteFrom(returns);
  renderTrade(d);
  renderTable(d);
}}

// ── CHARTS ────────────────────────────────────────────────
function renderCostByYear(d, reuse) {{
  // Cost avoidance grouped by CLOSE year (when value was realized)
  // Rule: pre-2022 = SqFtRequested, 2022+ = SqFtAssigned (fallback to Requested)
  const byYear = {{}};
  reuse.forEach(r => {{
    const yr = r.year_close || r.year;
    if (yr) byYear[yr] = (byYear[yr]||0) + Math.max(sqftForCost(r),0);
  }});
  const years = Object.keys(byYear).sort();
  mkChart('c-cost-year', {{ type:'bar', data:{{
    labels: years,
    datasets:[{{ label:'Cost Avoided (by close year)', data:years.map(y=>byYear[y]*COST_SQFT),
      backgroundColor:'rgba(0,104,181,0.75)', borderRadius:5 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label: c=>'  $'+Math.round(c.parsed.y).toLocaleString() }} }} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, ticks:{{ callback:v=>v>=1e6?'$'+(v/1e6).toFixed(1)+'M':'$'+(v/1e3).toFixed(0)+'K' }}, beginAtZero:true }},
      x:{{grid:{{display:false}}}} }} }} }});
}}

function renderStatusDonut(d) {{
  const counts = counter(d,'status');
  const labels = Object.keys(counts);
  mkChart('c-status', {{ type:'doughnut', data:{{
    labels, datasets:[{{ data:labels.map(l=>counts[l]),
      backgroundColor:labels.map(l=>COLOR_MAP[l]||'#CCC'), borderWidth:2, hoverOffset:8 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'right', labels:{{ padding:12 }} }} }} }} }});
}}

function renderMonthly(d) {{
  const newM={{}}, closedM={{}};
  d.forEach(r => {{
    if (r.ym) newM[r.ym] = (newM[r.ym]||0)+1;
    if (r.date_close && r.status !== 'Open/Pending') {{
      const m = r.date_close.slice(0,7);
      if (m) closedM[m] = (closedM[m]||0)+1;
    }}
  }});
  const allM = [...new Set([...Object.keys(newM),...Object.keys(closedM)])].sort().slice(-30);
  mkChart('c-monthly', {{ type:'bar', data:{{
    labels: allM,
    datasets:[
      {{ label:'New', type:'line', data:allM.map(m=>newM[m]||0),
        borderColor:'#0068B5', backgroundColor:'rgba(0,104,181,.1)', fill:true, tension:.3, pointRadius:2 }},
      {{ label:'Closed', type:'bar', data:allM.map(m=>closedM[m]||0),
        backgroundColor:'rgba(0,199,253,.55)', borderRadius:4 }}
    ]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
}}

function renderQuarterly(d) {{
  // Count requests per quarter per year
  const qMap = {{}};
  d.forEach(r => {{
    if (!r.year || !r.quarter) return;
    const key = String(r.year);
    if (!qMap[key]) qMap[key] = {{1:0,2:0,3:0,4:0}};
    qMap[key][r.quarter]++;
  }});
  const years = Object.keys(qMap).sort().slice(-8);  // last 8 years
  const colors = ['#0068B5','#00A86B','#FF8C00','#CC3333'];
  mkChart('c-quarterly', {{ type:'bar', data:{{
    labels: years,
    datasets: [1,2,3,4].map((q,i) => ({{
      label: 'Q'+q,
      data: years.map(y => qMap[y] ? qMap[y][q] : 0),
      backgroundColor: colors[i], borderRadius: 3,
    }}))
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ x:{{ stacked:true, grid:{{display:false}} }},
              y:{{ stacked:true, grid:{{color:GRIDCOLOR}}, beginAtZero:true }} }} }} }});
}}

function renderBU(d, reuse) {{
  const total = counter(d,'bu'), reuseBU = counter(reuse,'bu');
  const top = topN(total,10).map(e=>e[0]);
  mkChart('c-bu', {{ type:'bar', data:{{
    labels: top,
    datasets:[
      {{ label:'Total',       data:top.map(b=>total[b]||0),   backgroundColor:'rgba(0,104,181,.55)', borderRadius:4 }},
      {{ label:'Reuse Fills', data:top.map(b=>reuseBU[b]||0), backgroundColor:'#00A86B',             borderRadius:4 }}
    ]
  }}, options:{{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ x:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, y:{{grid:{{display:false}}}} }} }} }});
}}

function renderPlanner(d, open) {{
  const total = counter(d,'planner'), openP = counter(open,'planner');
  const top = topN(total,8).map(e=>e[0]);
  mkChart('c-planner', {{ type:'bar', data:{{
    labels: top,
    datasets:[
      {{ label:'Total Handled', data:top.map(p=>total[p]||0), backgroundColor:'rgba(0,104,181,.5)', borderRadius:4 }},
      {{ label:'Open Now',      data:top.map(p=>openP[p]||0), backgroundColor:'#FF8C00',            borderRadius:4 }}
    ]
  }}, options:{{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'top' }} }},
    scales:{{ x:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, y:{{grid:{{display:false}}}} }} }} }});
}}

function renderAging(open) {{
  const now = new Date();
  const b = {{'0–30d':0,'31–60d':0,'61–90d':0,'90+d':0}};
  open.forEach(r => {{
    if (!r.date_req) return;
    const age = Math.floor((now-new Date(r.date_req))/86400000);
    if (age<=30) b['0–30d']++; else if(age<=60) b['31–60d']++; else if(age<=90) b['61–90d']++; else b['90+d']++;
  }});
  mkChart('c-aging', {{ type:'bar', data:{{
    labels: Object.keys(b),
    datasets:[{{ label:'Open', data:Object.values(b),
      backgroundColor:['#00A86B','#FF8C00','#FF4500','#CC3333'], borderRadius:5 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{display:false}} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
}}

function renderTouch(d) {{
  const tc = counter(d,'touch'); delete tc[''];
  mkChart('c-touch', {{ type:'doughnut', data:{{
    labels: Object.keys(tc),
    datasets:[{{ data:Object.values(tc), backgroundColor:['#0068B5','#00C7FD','#9CA3AF'], borderWidth:2, hoverOffset:6 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{ position:'bottom' }} }} }} }});
}}

function renderSqftDist(d) {{
  const b={{'<500':0,'500–1K':0,'1K–2.5K':0,'2.5K–5K':0,'5K–10K':0,'>10K':0}};
  d.forEach(r => {{
    const s = parseNum(r.sqft_req);
    if (s<=0) return;
    if(s<500)b['<500']++;else if(s<1000)b['500–1K']++;else if(s<2500)b['1K–2.5K']++;
    else if(s<5000)b['2.5K–5K']++;else if(s<10000)b['5K–10K']++;else b['>10K']++;
  }});
  mkChart('c-sqft-dist', {{ type:'bar', data:{{
    labels: Object.keys(b),
    datasets:[{{ label:'Requests', data:Object.values(b), backgroundColor:'#00C7FD', borderRadius:5 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{display:false}} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
}}

function renderSiteTo(d) {{
  const sc={{}};
  d.forEach(r=>{{ const s=r.site_to; if(s&&s.length>=2&&!['NA','N/A','0','TBD'].includes(s)) sc[s]=(sc[s]||0)+1; }});
  const top=topN(sc,10);
  mkChart('c-site-to', {{ type:'bar', data:{{
    labels:top.map(e=>e[0]), datasets:[{{ label:'Requests', data:top.map(e=>e[1]), backgroundColor:'rgba(0,104,181,.6)', borderRadius:5 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{display:false}} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
}}

function renderSiteFrom(returns) {{
  const sc={{}};
  returns.forEach(r=>{{ const s=r.site_from; if(s&&s.length>=2&&!['NA','N/A','0','TBD'].includes(s)) sc[s]=(sc[s]||0)+1; }});
  const top=topN(sc,10);
  mkChart('c-site-from', {{ type:'bar', data:{{
    labels:top.map(e=>e[0]), datasets:[{{ label:'Returns', data:top.map(e=>e[1]), backgroundColor:'rgba(0,199,253,.55)', borderRadius:5 }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false,
    plugins:{{ legend:{{display:false}} }},
    scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
}}

function renderTrade(d) {{
  if (!document.getElementById('c-trade')) return;
  const labels = ['Net Zero', 'Growth', 'Undefined'];
  const counts = {{'Net Zero':0,'Growth':0,'Undefined':0}};
  const sqft   = {{'Net Zero':0,'Growth':0,'Undefined':0}};
  d.forEach(r => {{
    const t = r.trade_type;
    if (!t) return;
    counts[t] = (counts[t]||0) + 1;
    sqft[t]   = (sqft[t]||0) + Math.max(parseNum(r.sqft_req), 0);
  }});
  const colors = labels.map(l => TRADE_COLORS[l] || '#9CA3AF');
  mkChart('c-trade', {{ type:'doughnut', data:{{
    labels, datasets:[{{ data:labels.map(l=>counts[l]), backgroundColor:colors, borderWidth:2, borderColor:'#fff' }}]
  }}, options:{{ responsive:true, maintainAspectRatio:false, cutout:'55%',
    plugins:{{ legend:{{position:'bottom'}}, tooltip:{{ callbacks:{{ label:c=>`${{c.label}}: ${{c.parsed}} request${{c.parsed===1?'':'s'}}` }} }} }} }} }});
  if (document.getElementById('c-trade-sqft')) {{
    mkChart('c-trade-sqft', {{ type:'bar', data:{{
      labels, datasets:[{{ label:'SqFt Requested', data:labels.map(l=>sqft[l]), backgroundColor:colors, borderRadius:5 }}]
    }}, options:{{ responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label:c=>fmtN(c.parsed)+' sqft' }} }} }},
      scales:{{ y:{{ grid:{{color:GRIDCOLOR}}, beginAtZero:true }}, x:{{grid:{{display:false}}}} }} }} }});
  }}
}}

// ── TABLE FILTERING ────────────────────────────────────────
function escapeHtml(text) {{
  if (!text) return '';
  const map = {{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}};
  return String(text).replace(/[&<>"']/g, m => map[m]);
}}

function getTableFilters() {{
  const q = (document.getElementById('tf-search')?.value || '').trim().toLowerCase();
  const sf = document.getElementById('tf-sub-from')?.value || '';
  const st = document.getElementById('tf-sub-to')?.value   || '';
  const cf = document.getElementById('tf-close-from')?.value || '';
  const ct = document.getElementById('tf-close-to')?.value   || '';
  const dmin = parseInt(document.getElementById('tf-days-min')?.value);
  const dmax = parseInt(document.getElementById('tf-days-max')?.value);
  const statuses = [...document.querySelectorAll('.tf-status-cb:checked')].map(e=>e.value);
  return {{ q, sf, st, cf, ct, dmin, dmax, statuses }};
}}

function applyTableFilters(rows) {{
  const tf = getTableFilters();
  return rows.filter(r => {{
    if (tf.q) {{
      const hay = (r.id+' '+r.title+' '+r.planner+' '+r.bu).toLowerCase();
      if (!hay.includes(tf.q)) return false;
    }}
    if (tf.statuses.length && !tf.statuses.includes(r.status)) return false;
    if (tf.sf && (!r.ym || r.ym < tf.sf)) return false;
    if (tf.st && (!r.ym || r.ym > tf.st)) return false;
    if (tf.cf && (!r.ym_close || r.ym_close < tf.cf)) return false;
    if (tf.ct && (!r.ym_close || r.ym_close > tf.ct)) return false;
    if (!isNaN(tf.dmin) && (r.days_open||0) < tf.dmin) return false;
    if (!isNaN(tf.dmax) && (r.days_open||0) > tf.dmax) return false;
    return true;
  }});
}}

function refreshStatusBtnLabel() {{
  const btn = document.getElementById('tf-status-btn');
  if (!btn) return;
  const cbs = [...document.querySelectorAll('.tf-status-cb')];
  const checkedVals = cbs.filter(c=>c.checked).map(c=>c.value);
  if (checkedVals.length === 0 || checkedVals.length === cbs.length) {{
    btn.textContent = 'All ▾';
  }} else if (checkedVals.length === 1) {{
    btn.textContent = `${{checkedVals[0]}} ▾`;
  }} else {{
    btn.textContent = `${{checkedVals.length}} selected ▾`;
  }}
}}

function toggleStatusDropdown(e) {{
  if (e) e.stopPropagation();
  document.getElementById('tf-status-pop').classList.toggle('open');
}}

function clearTableFilters(rerender) {{
  ['tf-search','tf-sub-from','tf-sub-to','tf-close-from','tf-close-to','tf-days-min','tf-days-max']
    .forEach(id => {{ const el = document.getElementById(id); if (el) el.value = ''; }});
  const defaults = new Set(['Open/Pending', 'Hold']);
  document.querySelectorAll('.tf-status-cb').forEach(cb => cb.checked = defaults.has(cb.value));
  refreshStatusBtnLabel();
  if (rerender !== false) renderTable(filteredData);
}}

function renderTable(d) {{
  const filtered = applyTableFilters(d);
  const sorted = [...filtered].sort((a,b) => {{
    let av=a[sortKey]??'', bv=b[sortKey]??'';
    if (typeof av==='number') return sortDir*(av-bv);
    return sortDir*String(av).localeCompare(String(bv));
  }});
  const show = sorted.slice(0,200);
  document.getElementById('tbl-body').innerHTML = show.map(r => {{
    const idCell    = r.web_url
      ? `<a class="sp-link" href="${{escapeHtml(r.web_url)}}" target="_blank" rel="noopener" title="Open in SharePoint">${{escapeHtml(r.id)}}</a>`
      : `<strong>${{escapeHtml(r.id)}}</strong>`;
    const titleCell = r.web_url
      ? `<a class="sp-link" href="${{escapeHtml(r.web_url)}}" target="_blank" rel="noopener" title="Open in SharePoint">${{escapeHtml(r.title)}}</a>`
      : escapeHtml(r.title);
    const siteLabel = r.site_to ? escapeHtml(r.site_to) : '—';
    const siteTitle = r.site_name ? ` title="${{escapeHtml(r.site_name)}}"` : '';
    const tradeCell = r.trade_type
      ? `<span style="display:inline-block;padding:2px 7px;border-radius:9px;font-size:10px;font-weight:600;background:${{TRADE_COLORS[r.trade_type]||'#9CA3AF'}};color:#fff" title="${{escapeHtml(r.trade_raw||r.trade_type)}}">${{escapeHtml(r.trade_type)}}</span>`
      : '<span style="color:#CCC">—</span>';
    // Cost avoided: Reuse only; 2022+ uses sqft_asn if > 0, else sqft_req
    let costCell = '—';
    if (r.status === 'Reuse') {{
      const sqftUsed = sqftForCost(r);
      const cost = sqftUsed * COST_SQFT;
      const src  = ((r.year_close||r.year||0) >= 2022 && r.sqft_asn > 0) ? 'asn' : 'req';
      costCell = `<span style="color:#00A86B;font-weight:600" title="${{sqftUsed.toLocaleString()}} sqft (${{src}}) × $${{COST_SQFT}}">${{fmt$(cost)}}</span>`;
    }}
    r.cost_avoided = (r.status === 'Reuse') ? sqftForCost(r) * COST_SQFT : 0; // enable sort
    return `<tr>
      <td>${{idCell}}</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{escapeHtml(r.title)}}">${{titleCell}}</td>
      <td>${{badge(r.status)}}</td>
      <td>${{escapeHtml(r.bu)}}</td>
      <td style="font-size:11px">${{escapeHtml(r.region||'—')}}</td>
      <td${{siteTitle}} style="font-size:11px">${{siteLabel}}</td>
      <td style="text-align:center">${{tradeCell}}</td>
      <td style="text-align:right">${{r.sqft_req>0?r.sqft_req.toLocaleString():'—'}}</td>
      <td style="text-align:right;color:${{(r.sqft_asn>0&&(r.year_close||r.year||0)>=2022)?'#00A86B':'#CCC'}}">${{r.sqft_asn>0?r.sqft_asn.toLocaleString():'—'}}</td>
      <td style="text-align:right;color:${{r.sqft_ret_act>0?'#00A86B':'#CCC'}}">${{r.sqft_ret_act>0?r.sqft_ret_act.toLocaleString():'—'}}</td>
      <td style="text-align:right">${{costCell}}</td>
      <td>${{escapeHtml(r.planner)}}</td>
      <td>${{r.ym||'—'}}</td>
      <td style="color:var(--gray)">${{r.ym_close||'—'}}</td>
      <td style="text-align:center">${{r.quarter?'Q'+r.quarter:'—'}}</td>
      <td style="text-align:right">${{r.days_open>0?r.days_open:'—'}}</td>
    </tr>`;
  }}).join('');
  document.getElementById('tbl-footer').textContent =
    `Showing ${{show.length}} of ${{filtered.length}} filtered, ${{d.length}} total (max 200). Use filters to narrow.`;
}}

function sortTable(key) {{
  if (sortKey===key) sortDir*=-1; else {{ sortKey=key; sortDir=-1; }}
  renderTable(filteredData);
}}

// Init — try fetch (GitHub Pages / web server), fall back to inline data (local file open)
const INLINE_DATA = {rows_js};
(function() {{
  const bar = document.getElementById('insight-bar');
  if (bar) bar.textContent = '⏳ Loading dashboard data…';

  function boot(data) {{
    ALL_DATA = data;
    filteredData = [...ALL_DATA];
    const statusSet = new Set(ALL_DATA.map(r => r.status).filter(s => s));
    const statuses = Array.from(statusSet).sort();
    const statusPop = document.getElementById('tf-status-pop');
    if (statusPop) {{
      statusPop.innerHTML = statuses.map(s => `
        <label class="tf-status-cb">
          <input type="checkbox" class="tf-status-cb" value="${{escapeHtml(s)}}" checked onchange="renderTable(filteredData)">
          <span>${{escapeHtml(s)}}</span>
        </label>
      `).join('');
    }}
    refreshStatusBtnLabel();
    const defaults = new Set(['Open/Pending', 'Hold']);
    document.querySelectorAll('.tf-status-cb').forEach(cb => cb.checked = defaults.has(cb.value));
    refreshStatusBtnLabel();
    const _cy = new Date().getFullYear();
    document.getElementById('f-year-from').value = _cy;
    document.getElementById('f-year-to').value   = _cy;
    applyFilters();
  }}

  // 1) Try fetch (GitHub Pages — always gets latest data)
  fetch('let_data.json')
    .then(r => {{ if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); }})
    .then(boot)
    .catch(() => {{
      // 2) Fallback: inline data baked into HTML (works when opened as local file)
      if (INLINE_DATA && INLINE_DATA.length) {{
        boot(INLINE_DATA);
      }} else {{
        if (bar) bar.innerHTML =
          '<strong style="color:#CC3333">⚠ Could not load data. Re-run generate_dashboard.py.</strong>';
      }}
    }});
}})();
</script>
</body>
</html>"""


def run_sanity_checks(rows):
    """Return list of (level, message). level = 'ERROR' or 'WARN'."""
    issues = []
    known_statuses = set(STATUS_COLOR.keys()) | {"Unknown"}
    for r in rows:
        rid = r.get("id", "?")
        if r.get("date_req") and r.get("quarter"):
            try:
                m = int(r["date_req"][5:7])
                exp = (m - 1) // 3 + 1
                if r["quarter"] != exp:
                    issues.append(("ERROR", f"ID {rid}: quarter={r['quarter']} but date_req={r['date_req']} → expected Q{exp}"))
            except (ValueError, IndexError): pass
        if r.get("date_close") and r.get("quarter_close"):
            try:
                m = int(r["date_close"][5:7])
                exp = (m - 1) // 3 + 1
                if r["quarter_close"] != exp:
                    issues.append(("ERROR", f"ID {rid}: quarter_close={r['quarter_close']} but date_close={r['date_close']} → expected Q{exp}"))
            except (ValueError, IndexError): pass
        if r.get("date_close") and r.get("year_close"):
            try:
                if r["year_close"] != int(r["date_close"][:4]):
                    issues.append(("ERROR", f"ID {rid}: year_close={r['year_close']} but date_close={r['date_close'][:4]}"))
            except (ValueError, IndexError): pass
        if r.get("is_open") and r.get("date_close"):
            issues.append(("WARN", f"ID {rid}: is_open=True but has date_close={r['date_close']} (status={r.get('status')})"))
        if not r.get("is_open") and not r.get("date_close") and r.get("status") not in ("Unknown", ""):
            issues.append(("WARN", f"ID {rid}: status={r.get('status')} (closed) but date_close missing"))
        req = r.get("sqft_req") or 0
        asn = r.get("sqft_asn") or 0
        if asn > 0 and req > 0 and asn > req * 4:
            issues.append(("WARN", f"ID {rid}: sqft_asn={asn:,} is >4× sqft_req={req:,}"))
        if r.get("status") and r["status"] not in known_statuses:
            issues.append(("ERROR", f"ID {rid}: unknown status '{r['status']}'"))
    return issues


def main():
    print("LET Interactive Dashboard Generator v2")
    rows = load_and_clean()
    reuse = [r for r in rows if r["status"] == "Reuse"]
    sqft_asn = sum(r["sqft_asn"] for r in reuse if r["sqft_asn"] > 0)
    sqft_req = sum(r["sqft_req"] for r in reuse if r["sqft_req"] > 0)
    print(f"  Loaded {len(rows)} records  |  Reuse: {len(reuse)}")
    print(f"  SqFt Assigned (official OKR): {sqft_asn:,}  => Cost: ${sqft_asn*COST_PER_SQFT:,.0f}")
    print(f"  SqFt Requested (higher):      {sqft_req:,}  => Cost: ${sqft_req*COST_PER_SQFT:,.0f}")

    issues = run_sanity_checks(rows)
    errors = [(l, m) for l, m in issues if l == "ERROR"]
    warns  = [(l, m) for l, m in issues if l == "WARN"]
    if issues:
        print(f"\n  DATA QUALITY: {len(errors)} errors, {len(warns)} warnings")
        for level, msg in issues[:30]:
            print(f"    [{level}] {msg}")
        if len(issues) > 30:
            print(f"    ... and {len(issues)-30} more")
    else:
        print("  DATA QUALITY: ✓ all checks passed")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    html = generate_html(rows, generated_at, issues)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    size_kb = os.path.getsize(OUTPUT_FILE) / 1024
    print(f"  Saved {size_kb:.0f} KB -> {OUTPUT_FILE}")

    # Write separate data file for GitHub Pages fetch()
    DATA_FILE = os.path.join(os.path.dirname(__file__), "let_data.json")
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, separators=(',', ':'))
    data_kb = os.path.getsize(DATA_FILE) / 1024
    print(f"  Saved {data_kb:.0f} KB -> {DATA_FILE}")
    print("  Filters: Month range, Q1-Q4 toggle, Status checkboxes, BU/Planner/Site multi-select, Touch")
    print("  SqFt toggle: Assigned (official ~$484M) vs Requested (~$611M)")

if __name__ == "__main__":
    main()
