# LET Vetting — Improvements Backlog

Living list. Add items as we spot them. Work from here next session.
Status: `Idea` → `Planned` → `In Progress` → `Done`

---

## ✅ Done (shipped)

| Commit | What |
|--------|------|
| `6c097da` | Hold as separate status + amber color |
| `a54794f` | Filter bypass for unlisted Site/Planner/BU values |
| `db20766` | CRESD-based site normalization (937 codes, letters-only fallback) |
| `a181d1f` | Trade Net Zero vs Growth tracking (filter, chart, table column); fix ITK=Japan |
| `2a363f6` | **Q-filter bug fix** (Req 2433-style: request in Q3-2025 closed Q2-2026 no longer shown under "2026 Q3"); **build-time sanity checks** (26 data warnings surfaced); **QA banner** in dashboard header |

---

## Power App / SharePoint Form

| # | Area | Improvement | Why | Effort | Priority | Status | Notes |
|---|------|-------------|-----|--------|----------|--------|-------|
| 1 | Form validation | Make **LETClassification** (Net Zero / Growth) a **required** field — block save if empty | Stop "Yet to be Defined" leakage; 4 of 10 classified were closed without ever being properly set | S | High | Idea | **New items only**. Block submit, no default — force user to pick |
| 2 | Site/location source | Pull **FWR / building / site** from **UDM** ([RE&WS UDM library](https://intel.sharepoint.com/sites/RealEstateandWorkplaceServices/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FRealEstateandWorkplaceServices%2FShared+Documents%2FUDM)) instead of free text | Eliminate typos + removes 5-tier CRESD normalization hack | M-L | High | Idea | UDM is SharePoint-based → can read via Graph (same auth). Need to find which list has site master |
| 3 | Planner | Convert **Planner** from free text → **People picker / dropdown** | Avoid typos. Today just Yariv, but was 4 people historically (2 AMER, 1 EMEA, 1 APAC) | S | Med | Idea | Global list. Low urgency while solo |
| 4 | BU / Group | Convert to dropdowns from HR/org system | Free-text filter noise | M | Low | Idea | |
| 5 | Status transitions | Enforce workflow (can't go Withdrawn → Open) | Data integrity | S | Low | Idea | |
| 6 | Audit log | Track who changed what & when | No provenance today | M | Low | Idea | SP versioning already on — just surface it |

---

## Dashboard

| # | Area | Improvement | Why | Effort | Priority | Status | Notes |
|---|------|-------------|-----|--------|----------|--------|-------|
| D1 | Trade chart | **Decide on 4 closed "Yet to be Defined"** (IDs 2470, 2475, 2520.1, 2527) | Closed without ever being set — inflate Undefined slice. **Option A:** exclude (treat as legacy). **Option B:** split slice into "Open – Yet to Decide" (6) and "Closed – Never Set" (4) | S | Med | **Pending decision** | Yariv to decide which option |
| D2 | QA warnings — fix data | Act on the 26 warnings now surfaced in build output | e.g. 6 closed items missing date_close, 1 Hold with both is_open=True and date_close, 14 sqft_asn >> sqft_req | M | High | Idea | Warnings listed at every build. Some may need SP list fixes; others need normalization tweaks |
| D3 | Unknown region (125 rows) | Map remaining sites: YH, HYD, PENANG, CAMPEON, YTC, RHM1, RM, SEO, etc. | Improves region accuracy | M | Med | Idea | Many will auto-resolve if #2 (UDM) lands |
| D4 | Auto-refresh verification | Confirm scheduled refresh runs reliably | Currently unclear if task is active | S | Med | Idea | Check `_auto_refresh_runner.ps1` task in Task Scheduler |
| D5 | Cost benchmark | Make `$692/sqft` editable from dashboard UI or `config.json` | Today requires rebuild to change | S | Low | Idea | |
| D6 | Per-row cost column | Add **Cost Avoided** to bottom table (per row, Reuse only) | Can't audit which row contributed what today (e.g. Req 2485 = $401K) | S | Med | Idea | Rule: `sqft_asn × 692` for 2022+, `sqft_req × 692` pre-2022 |
| D7 | Drill-down | Click chart slice → filter table to that subset | UX | M | Low | Idea | |

---

## Exploration / Research

| # | Topic | Goal | Notes |
|---|-------|------|-------|
| E1 | UDM library structure | Find which list in [RE&WS UDM](https://intel.sharepoint.com/sites/RealEstateandWorkplaceServices/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FRealEstateandWorkplaceServices%2FShared+Documents%2FUDM) holds the FWR/site master. Confirm read via current Graph token | Prerequisite for #2 + #3 |
| E2 | UDM update cadence | How often does it refresh? Can we detect diffs? | Affects cache strategy |

---

## Cost Avoidance — Quick Reference

- **Benchmark:** `$692/sqft` ([generate_dashboard.py L13](generate_dashboard.py#L13))
- **Rule:** Pre-2022 → `sqft_req × $692` | 2022+ → `sqft_asn × $692` (falls back to req if asn=0)
- **Only `status=Reuse` rows contribute**
- Example: Req 2485 — Reuse, 2026, sqft_asn=580 → **$401,360**
- Official OKR ~$484M cumulative uses assigned-only methodology

---

## Data Quality Issues Found (build `2a363f6`, 2026-07-02)

Run `py generate_dashboard.py` to see current warnings. As of last build: **0 errors, 26 warnings**.

Key categories:
- **6 closed items with missing date_close** (IDs: 2382, 2377, 2365, 2363, 2362, 2281, 2226, 2178, 2199, 2031, 1886, 1711, 2527) — fix in SP list
- **ID 2514: Hold with is_open=True AND date_close** — SP list entry needs cleanup
- **14 items with sqft_asn > 4× sqft_req** — some may be valid (partial request), worth reviewing

---

## Resolved Decisions

- LaMP/Tririga API → **skip, use UDM instead** (SharePoint, same Graph auth)
- Planner scope → **global**
- Required-field rollout → **new items only**
