# Property Management – Automated Renewal Workbook

**Version 2.7.0 (2026-07-07)**  

A VBA-powered Excel system for property management teams to automate monthly lease renewal tracking. It ingests reports directly from Yardi and RealPage, populates pre-structured month sheets organized by floor plan, and builds a rolling multi-year performance summary — eliminating manual data entry and providing consistent renewal analytics across your portfolio.

---

## Table of Contents

1. [What This System Does](#what-this-system-does)
2. [Prerequisites](#prerequisites)
3. [First-Time Setup](#first-time-setup)
   - [Step 1 – Enable Macro Trust Settings](#step-1--enable-macro-trust-settings)
   - [Step 2 – Import All 9 Modules](#step-2--import-all-9-modules)
   - [Step 3 – Wire Up the Change Event](#step-3--wire-up-the-change-event)
   - [Step 4 – Run SetupWorkbook](#step-4--run-setupworkbook)
   - [Step 5 – Configure the Property Setup Sheet](#step-5--configure-the-property-setup-sheet)
   - [Step 6 – Generate Month Sheets](#step-6--generate-month-sheets)
   - [Step 7 – Create the Overview Sheet](#step-7--create-the-overview-sheet)
4. [Monthly Workflow](#monthly-workflow)
   - [Source Reports Required](#source-reports-required)
   - [Running the Import](#running-the-import)
   - [Manual Fields to Fill After Import](#manual-fields-to-fill-after-import)
5. [MTM Tracker Workflow](#mtm-tracker-workflow)
6. [Column Reference (Month Sheets)](#column-reference-month-sheets)
7. [Overview Sheet](#overview-sheet)
8. [Dynamic Row Insertion](#dynamic-row-insertion)
9. [Health Check](#health-check)
10. [Module Reference](#module-reference)
11. [Troubleshooting](#troubleshooting)
12. [Version History](#version-history)

---

## What This System Does

Each month, renewal coordinators typically copy-paste data from multiple Yardi and RealPage reports into a tracking spreadsheet by hand. This workbook automates that process:

- **Reads up to 5 source reports** (Yardi Rent Roll, Yardi Unit Statistics, RealPage Renewal Offer Analysis, Unit Rents Grid, Move-in Box Score) and merges them into a single month sheet
- **Organizes units by floor plan group** (e.g., 1BD, 2BD, etc.) with running totals per section
- **Pre-populates** market rent, current rent, YieldStar recommended increase, new lease rent, current lease term, occupied avg rent, and recent move-in avg for every renewing unit
- **Tracks renewal decisions** month by month and surfaces them in a **multi-year summary Overview sheet** with 10 key metrics per month, YTD totals, and exclude toggles for in-progress months
- **Grows dynamically** — when you fill the second-to-last buffer row in a section, a new blank row is inserted automatically so you never run out of space

---

## Prerequisites

| Requirement | Detail |
|---|---|
| Excel version | Excel 2016 or later (Windows). Excel for Mac is not supported — `Application.FileDialog` is Windows-only. |
| Macros enabled | Macros must be enabled for the workbook. |
| VBA project access | Required for `SetupWorkbook` to verify all modules are present (one-time). |
| Source reports | Yardi Rent Roll is required every month. All other reports are optional but recommended. |

---

## First-Time Setup

Complete these steps once per workbook. After setup, only the [Monthly Workflow](#monthly-workflow) section applies.

### Step 1 – Enable Macro Trust Settings

Two settings must be enabled in Excel before importing the modules:

1. Open Excel → **File → Options → Trust Center → Trust Center Settings**
2. Go to **Macro Settings**:
   - Select **Enable all macros** (or at minimum "Enable VBA macros")
   - Check **Trust access to the VBA project object model**
3. Click OK and restart Excel if prompted.

> **Why:** The `SetupWorkbook` macro programmatically checks that all 9 modules (including `modMTM`) are present before adding buttons. This requires VBA project object model access. Without it, setup falls back to a warning but still runs.

---

### Step 2 – Import All 9 Modules

Open the VBA Editor (**Alt + F11**), then import every `.bas` file in this folder:

**File → Import File** → select each file below, in this order:

| # | File | Why this order matters |
|---|---|---|
| 1 | `modConfig.bas` | **Must be first.** Declares the shared `PropConfig` type and all public constants. All other modules depend on it. |
| 2 | `modSheetUtils.bas` | Low-level sheet and range helpers used by Import, Dynamic, Setup, and Overview. |
| 3 | `modReaders.bas` | All file-reading logic (Yardi, RealPage, Grid, Box Score). Depends on modConfig. |
| 4 | `modImport.bas` | Import orchestration and sheet-fill logic. Depends on modConfig, modReaders, modSheetUtils. |
| 5 | `modDynamic.bas` | Dynamic row insertion on sheet change. Depends on modSheetUtils. |
| 6 | `modSetup.bas` | Creates the Property Setup sheet and month sheets. Depends on modConfig, modSheetUtils. |
| 7 | `modOverview.bas` | Builds the multi-year summary. Depends on modConfig, modSheetUtils. |
| 8 | `modAdmin.bas` | One-time setup wizard and health check. Depends on all other modules. |
| 9 | `modMTM.bas` | MTM tracker sheet refresh and import tools (see [MTM Tracker Workflow](#mtm-tracker-workflow)). Depends on modConfig, modReaders, modSheetUtils, modAdmin. **Included in `SetupWorkbook`'s module-presence check** — import it like every other module. |

After importing, the Project Explorer should show all 9 modules under your workbook's **Modules** folder. If any are missing, re-import them.

---

### Step 3 – Wire Up the Change Event

The dynamic row-insertion feature (see [Dynamic Row Insertion](#dynamic-row-insertion)) is triggered by Excel's `Workbook_SheetChange` event. You must add a one-time code snippet to `ThisWorkbook`:

1. In the VBA Editor, double-click **ThisWorkbook** in the Project Explorer.
2. Paste the following:

```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    modDynamic.HandleSheetChange Sh, Target
End Sub
```

3. Save with **Ctrl + S**.

> **Note:** If this event is not wired, the workbook functions normally — the only missing behavior is automatic buffer row insertion when a section fills up.

---

### Step 4 – Run SetupWorkbook

1. Back in Excel, open the **Immediate Window** in the VBA Editor (**Ctrl + G**) and type:
   ```
   modAdmin.SetupWorkbook
   ```
   Press Enter. Alternatively, use **Macro → Run → modAdmin.SetupWorkbook**.

2. SetupWorkbook will:
   - Verify all 9 modules are imported (including `modMTM` — see [Step 2](#step-2--import-all-9-modules))
   - Find your Overview sheet (or use Sheet 1 if none exists)
   - Add 5 buttons to that sheet:
   - Add the 4 MTM tracker buttons to the `MTM` sheet if it already exists (if it doesn't yet, they're wired automatically the first time you run **Refresh MTM Tracker** — see [MTM Tracker Workflow](#mtm-tracker-workflow))

| Button | Macro | Purpose |
|---|---|---|
| Import Monthly Data | `modImport.ImportMonthlyData` | Run each month to pull in Yardi/RP data |
| Generate Month Sheets | `modSetup.GenerateMonthSheets` | Create blank month tabs for a year |
| Create Setup Sheet | `modSetup.CreateSetupSheet` | Build/rebuild the Property Setup configuration tab |
| Create Overview | `modOverview.CreateOverviewSheet` | Build/rebuild the multi-year summary tab |
| Health Check | `modAdmin.HealthCheck` | Diagnose issues with config, sheets, or named ranges |

---

### Step 5 – Configure the Property Setup Sheet

Click **Create Setup Sheet**. A new tab called `Property Setup` is added (or rebuilt if it already exists). Fill in every yellow-highlighted cell:

#### Section A – Property Info

| Field | Description | Example |
|---|---|---|
| Property Full Name | Appears in the title bar and Overview header | `Property Name` |
| Property Short Name | Used as a column header label inside month sheets | `Short Name` |
| Workbook Year | The primary year for this workbook | `2027` |
| Unit Number Pattern(s) | Comma-separated patterns where `N` = any digit, `A` = any letter. Used to identify unit rows in Yardi reports. | `NN-NNN` or `NNNA, NNNB` |
| MTM Cap % | Maximum month-to-month rent increase (decimal). Used in overview calculations. | `0.05` |
| MTM Cap Through Date | Date through which the MTM cap applies | `12/31/2027` |
| Buffer Rows | Blank rows to leave at the bottom of each floor-plan section for new units. Minimum 1, recommended 2. Changes take effect immediately without re-running SetupWorkbook. | `2` |
| # of Floor Plan Groups | Informational only — the list below is the actual source of truth. | `6` |

#### Section B – Floor Plan Groups

List every floor plan group name, one per row, under the **Floor Plan Groups** header. These become the section headers on every month sheet. Names must match exactly what you use in the Yardi Code Map below.

Example:
```
Studio
1 Bedroom
1 Bedroom Den
2 Bedroom
2 Bedroom Den
3 Bedroom
```

#### Section C – Yardi Code Map

Two columns: **Yardi Code** (left) and **Floor Plan Group** (right).

Every Yardi unit type code that appears in your Rent Roll must be listed here and mapped to one of the Floor Plan Groups above. Group names must match exactly (case-insensitive).

Example:
```
s1a      Studio
s1b      Studio
1a       1 Bedroom
1a-den   1 Bedroom Den
2a       2 Bedroom
2b       2 Bedroom
```

> **Important:** The Setup sheet is pre-filled with example floor plan names and Yardi codes from a sample property. **Replace all of this example data entirely** — every floor plan group name, every Yardi code, and every floor plan mapping — before running your first import. Using the example codes will cause all of your real units to be flagged as unmapped.

> If a unit's Yardi code is not in this map after you've set it up, the unit is skipped during import and its code appears in the "Unmapped Yardi Codes" warning at the end of the import. Add the missing code to the map and re-import.

#### Section D – Column Fallbacks

The Yardi Rent Roll reader always uses the column numbers configured here — it does not attempt header detection. The **Unit Rents Grid** and **Move-in Box Score** readers auto-detect columns by header name first and only fall back to these numbers when a header is not found. The **Renewal Offer Analysis** (`.csv`) relies entirely on header detection and has no configurable column fallbacks — if its headers are not recognized, that file will be silently skipped with no data imported.

Set these to match your actual report layouts. Only change them if you see import failures or blank columns after import.

| Field | What it points to |
|---|---|
| Rent Roll: Unit Col | Column in the Yardi Rent Roll that contains the unit number |
| Rent Roll: Unit Type Col | Column with the Yardi unit type code |
| Rent Roll: Resident Col | Column with the resident name |
| Rent Roll: Market Rent Col | Column with the market/asking rent |
| Rent Roll: Actual Rent Col | Column with the actual/current rent |
| Rent Roll: Lease Expiry Col | Column with the lease expiration date |
| Rents Grid: Unit Col | Column in the Unit Rents Grid for unit number |
| Rents Grid: Cur Eff Rent Col | Current effective rent column |
| Rents Grid: Best Offer Col | Best renewal offer rent column |
| Rents Grid: Best Term Col | Best offer lease term (months) column |
| Rents Grid: New Lease Col | New lease rent column |
| Box Score: Unit Col | Column in Move-in Box Score for unit number |
| Box Score: Unit Type Col | Unit type code column |
| Box Score: Rent Col | Effective rent column |
| Box Score: Move-In Col | Move-in date column |

After filling out the setup sheet, run **Health Check** (see [Health Check](#health-check)) to verify the configuration is valid.

---

### Step 6 – Generate Month Sheets

Click **Generate Month Sheets**. You will be prompted three times:

1. **Which month(s):** Enter a single month number (`6`), a range (`1-6`), a comma-separated list (`1,4,7`), or `ALL` to generate all 12 sheets at once.
2. **Year:** Enter the 4-digit year (e.g., `2027`).
3. **Empty rows per section (optional):** Press Enter to use your configured Buffer Rows value, or enter a number to override for this run only. Useful when back-filling historical months that need fewer rows.

The system creates the requested month sheets (named `Jan 27`, `Feb 27`, etc.) with:

- A title row with the property name and month/year
- Section headers (grey bars) for each floor plan group you defined
- Pre-built formulas for renewal calculations
- Buffer rows below each section for new units

> Run this once per year. If a sheet for a requested month already exists, you will be prompted whether to replace it (**all existing data on that sheet will be lost if you choose Yes**) or keep it.

> **Note:** Generate Month Sheets also refreshes the Overview sheet automatically after completing. You do not need to click Create Overview separately afterward.

---

### Step 7 – Create the Overview Sheet

Click **Create Overview**. A sheet called `Renewal Overview` is created (or rebuilt) with:

- One block per year detected from existing month sheets
- 12 data rows per year (one per month), pulling stats from named ranges on each month sheet
- A YTD Total row with renewal-weighted averages
- An Exclude column — type `x` to grey out an in-progress month and exclude it from YTD totals

The Overview rebuilds itself non-destructively: your existing `x` exclude marks are preserved when you click Create Overview again.

---

## Monthly Workflow

### Source Reports Required

Each month you will import up to 5 files. Only the Yardi Rent Roll is required; all others are optional but enable additional columns.

| # | Report | Source | Format | What it fills |
|---|---|---|---|---|
| 1 | **Yardi Rent Roll** | Yardi | `.xlsx` | **REQUIRED.** Unit list, resident names, floor plan codes, market rent (col K), actual/current rent (col E), lease expiry date (col P) |
| 2 | **Yardi Unit Statistics** | Yardi | `.xlsx` | Occupied average rent by floor plan (col L) and inplace lease average (col X) |
| 3 | **RealPage Renewal Offer Analysis** | RealPage | `.csv` | YieldStar recommended increase (col F) and current lease term (col T) — used as fallback when the Unit Rents Grid is not provided |
| 4 | **Unit Rents Grid** | RealPage | `.xlsx` | YieldStar recommended increase (col F), new lease rent (col N), current lease term (col T), best offer term (col U). **Preferred over the RP csv for cols F and T.** |
| 5 | **Move-in Box Score** | Yardi CRM | `.xls` | 3-month average effective rent for recent move-ins by floor plan (col M). Pulled from the Resident Activity Detail drill-down in the Yardi CRM Box Score Summary — see [PULLING_REPORTS.md](PULLING_REPORTS.md) for step-by-step instructions. **Files containing VBA macros are rejected automatically.** |

> **Tip:** Export all reports for the same month before starting the import. The import wizard walks you through file selection one at a time — click Cancel on any file you don't have to skip it.

---

### Running the Import

1. Click **Import Monthly Data** on the Overview sheet (or run `modImport.ImportMonthlyData` from the VBA editor).
2. Enter the **month number** (1–12) when prompted.
3. Enter the **year** (e.g., `2027`).
4. Read the instructions dialog, then click OK.
5. For each of the 5 steps, a file picker dialog opens:
   - Navigate to the report file and click Open.
   - Click **Cancel** to skip that report (the wizard continues to the next file).
6. After all files are selected, the import runs automatically.
7. A completion summary appears showing:
   - How many units were imported and to which sheet
   - Which optional files were skipped
   - Any **unmapped Yardi codes** (unit types not in your code map — these units were skipped)

> If unmapped codes appear, add them to the Yardi Code Map on the Property Setup sheet and re-run the import. Re-importing overwrites existing data for that month.

---

### Manual Fields to Fill After Import

Two columns are always left blank intentionally — they require your judgment:

| Column | Field | What to enter |
|---|---|---|
| **G** | *[Property Short Name]* Recommended Increase | Your property's recommended renewal increase for each unit, in dollars |
| **I** | Pet Fees | Verify and enter pet fee amounts per unit (defaults to 0 after import) |

Review and spot-check the other imported columns, particularly:
- **Col F** (YieldStar Inc) — confirms the Yieldstar recommended increase came through
- **Col T** (Current Term) — confirms the tenant's current lease term in months
- **Col M** (Recent Avg Eff. Rent) — floor-plan-level average from recent move-ins

---

## MTM Tracker Workflow

The MTM Tracker is a dedicated sheet (`MTM`) that tracks month-to-month units separately from the monthly renewal month sheets. It is built and refreshed with `modMTM.RefreshMTMSheet`, independent of the monthly `ImportMonthlyData` flow — though `ImportSelectedMTM` places tracked units onto month sheets, and `FillMTMRows` later enriches those rows during the next monthly import.

> **Sheet rename:** the tab is named `MTM`. This is automatic and self-healing — if your workbook still has the previous `MTM & STL` tab, the next time you run `RefreshMTMSheet`, `ImportSelectedMTM`, or `ClearSelectionMTM` it will be renamed in place (all data and formatting preserved). No manual step is required.

### Refreshing the Tracker

Click **Refresh MTM Tracker** (or run `modMTM.RefreshMTMSheet`). You'll be prompted for a single file:

| # | File | Required? | What it's used for |
|---|---|---|---|
| 1 | Yardi Rent Roll | **Yes** — Cancel aborts the refresh | Base source for all tracked units (unit, name, floor plan, current rent, market rent, lease expiry) |

`RefreshMTMSheet` creates the `MTM` sheet if it doesn't already exist, formats it, updates existing tracked units, removes units that no longer appear in the current scan, and adds any newly-detected units. The sheet is rebuilt as a single flat list, sorted by Next Increase date after every refresh.

### Tracker Column Layout

| Col | Field | Notes |
|---|---|---|
| A | Unit | |
| B | Name | |
| C | Floor Plan | Yardi unit type code |
| D | Lease Expiry | From Yardi Rent Roll |
| E | Current Rent | From Yardi Rent Roll |
| F | Last Increase | **Manual entry** — the date of the unit's last rent increase. Drives the calculated Next Increase date. |
| G | Market Rent | From Yardi Rent Roll |
| H | Next Increase | Calculated — Last Increase + 1 year + 1 day |
| I | Status | See [Status Column Values](#status-column-values) below |
| J | Notes | Manual. If Last Increase (F) is blank and the unit's lease expiry is more than 15 months past-due, a note is auto-written here on refresh ("Lease expired 15+ months ago — verify resident ledger for actual increase history") — but only if J is currently blank; an existing manual note is never overwritten. |
| K | Import | Form Control checkbox — check to include the row in the next `ImportSelectedMTM` run |

### Status Column Values

| Status | Highlight | Meaning |
|---|---|---|
| `Eligible` | Green | Next Increase is today or in the past (or blank). |
| `Not Yet Eligible` | Amber/orange | Next Increase is a future date. |

Status is **informational only** — it does not affect checkbox availability or import behavior. A CM can check and import any row regardless of its Status.

### Importing Selected Units (`ImportSelectedMTM`)

Check the Import column (K) for the units you want to move onto a month sheet, then click **Import Selected MTM** (or run `modMTM.ImportSelectedMTM`). For each checked row, the unit is placed onto the correct month sheet — determined by its Next Increase date — in the correct floor-plan section.

`ImportSelectedMTM` also carries over the Resident Name, Floor Plan, and Current Rent from the tracker row onto the month sheet at the same time.

### `FillMTMRows` (runs during the monthly import)

`FillMTMRows` is not part of `RefreshMTMSheet` — it runs automatically as part of the normal monthly `ImportMonthlyData` import (see [Running the Import](#running-the-import)). It fills in additional fields on any month-sheet row that came from an MTM import, using that month's fresh report data:

- The monthly report data is allowed to **override** the Name / Floor Plan / Current Rent that `ImportSelectedMTM` carried over, in case the tracker snapshot was stale by import time.
- The RealPage Renewal Offer Analysis report also feeds these rows for column F, using the same grid-preferred / RealPage-fallback precedence used for regular renewal rows.
- Column T (Current Term) shows a bold, yellow-highlighted `"MTM"` tag for these rows instead of a numeric term.
- Column A (Renewal Status) is deliberately left untouched/blank for MTM rows — it stays a purely manual field. (See the [Column Reference](#column-reference-month-sheets) notes on columns A and T.)

### Buttons

These 4 buttons are wired onto the `MTM` sheet automatically — no manual setup step is required. `SetupWorkbook` adds them if the `MTM` sheet already exists; if it doesn't yet, `RefreshMTMSheet` adds them itself the first time it creates the sheet.

| Button | Macro | Purpose |
|---|---|---|
| Refresh MTM Tracker | `modMTM.RefreshMTMSheet` | Build/update the MTM tracker from the Yardi Rent Roll |
| Import Selected MTM | `modMTM.ImportSelectedMTM` | Place checked tracker rows onto the correct month sheet |
| Clear Selection MTM | `modMTM.ClearSelectionMTM` | Uncheck every Import checkbox on the tracker |
| Sort MTM Tracker | `modMTM.SortMTMSheet` | Manually re-sort by Next Increase (col H) on demand — also recalculates Next Increase/Status from any manual Last Increase edits before sorting — without running a full refresh |

---

## Column Reference (Month Sheets)

| Col | Letter | Field | Source |
|---|---|---|---|
| 1 | A | **Renewal Status** | **Manual entry** — dropdown: `Renewed`, `MTM`, `NTV` (Notice to Vacate), `Pending`. Drives row color (green = Renewed, pink = NTV, blue = MTM) and all Overview metrics (renewal count, signed revenue, capture ratio). This is the primary field you fill each month. *Note: the automated [MTM Tracker Workflow](#mtm-tracker-workflow) never writes to this column — it stays purely manual. The automated MTM signal is the highlighted tag in column T instead, to avoid confusing it with the manual `MTM` dropdown value here.* |
| 2 | B | Apt # (Unit Number) | Yardi Rent Roll |
| 3 | C | Resident Name | Yardi Rent Roll |
| 4 | D | Floor Plan (Yardi Code) | Yardi Rent Roll |
| 5 | E | Current Rent (Actual Rent) | Yardi Rent Roll |
| 6 | F | YieldStar Recommended Increase | Unit Rents Grid (preferred) or RP Renewal Offer Analysis |
| 7 | G | Property Recommended Increase | **Manual entry** |
| 9 | I | Pet Fees | **Manual entry** (defaults to $0) |
| 11 | K | Market Rent | Yardi Rent Roll (market/asking rent) |
| 12 | L | Occupied Avg Rent | Yardi Unit Statistics (weighted avg by floor plan) |
| 13 | M | Recent Move-In Avg Rent | Move-In Box Score (3-month avg by floor plan) |
| 14 | N | New Lease Rent | Unit Rents Grid |
| 16 | P | Lease Expiry Date | Yardi Rent Roll |
| 20 | T | Current Lease Term (months) | Unit Rents Grid (preferred) or RP Renewal Offer Analysis. *Rows placed by the [MTM Tracker Workflow](#mtm-tracker-workflow) (`ImportSelectedMTM`/`FillMTMRows`) show a bold, yellow-highlighted `"MTM"` tag here instead of a numeric term.* |
| 21 | U | Recommended Lease Term / Notes | Unit Rents Grid (best offer term) |
| 24 | X | Inplace Lease Avg | Yardi Unit Statistics (currently receives the same weighted average as col L) |

---

## Overview Sheet

The Overview sheet aggregates renewal performance across all months and years. It is rebuilt from live named ranges on each month sheet — no data is stored in the Overview itself.

### Metrics Tracked Per Month

| Metric | Description |
|---|---|
| # of Renewals | Total renewing units for the month |
| # of Increases | Units that received a rent increase |
| Average $ Increase | Dollar increase per renewal (YTD = renewal-weighted avg) |
| Average % Increase | Percent increase per renewal (YTD = renewal-weighted avg) |
| Potential Revenue | Sum of recommended increases across all renewing units |
| Signed Revenue | Sum of increases for units marked as renewed |
| # Signed | Count of units marked as renewed |
| Capture Ratio % | Signed / Total renewals (YTD = renewal-weighted avg) |
| MTM $ Increase | Premium captured from units that went month-to-month |
| Total $ Captured | Signed Revenue + MTM $ Increase |

### YTD Totals

The YTD Total row at the bottom of each year block uses **renewal-weighted averages** for Average $, Average %, and Capture Ratio — months with more renewals count proportionally, not as a simple average of the monthly figures.

### Exclude Column

Type `x` in the Exclude column for any month row to:
- Grey out the entire row with strikethrough text
- Remove that month from all YTD totals

Clear the `x` to include the month again. Exclude marks survive when you click Create Overview to rebuild.

---

## Dynamic Row Insertion

When you type a unit number into the second-to-last buffer row of any floor plan section, the system automatically inserts a new blank row above the section boundary — keeping your configured number of buffer rows available at all times.

This requires the `Workbook_SheetChange` event to be wired in `ThisWorkbook` (see [Step 3](#step-3--wire-up-the-change-event)).

**How it works:**
- Only fires on month sheets (sheets whose name matches a month pattern)
- Only watches column B (Apt#)
- When you fill the second-to-last buffer row, one new row is inserted with:
  - Formats and formulas copied from the row above
  - Data columns cleared (unit, name, rent, etc.)
  - Floor-plan averages inherited from the row above (cols L, M, X)
  - Pet Fees (col I) defaulted to $0

---

## Health Check

Click **Health Check** (or run `modAdmin.HealthCheck`) at any time to get a diagnostic report:

| Check | What it looks for |
|---|---|
| Module version | Confirms the version number from modConfig |
| Property Setup | Confirms the Setup sheet exists and all required fields are valid |
| Overview sheet | Confirms an Overview sheet exists |
| Month sheets | Counts all month sheets found; flags any that are missing their stats named ranges |
| Defined names | Lists any named ranges that resolve to `#REF` (broken references) |

If broken named ranges are found, regenerate the affected month sheet by deleting it and running Generate Month Sheets again for that year.

---

## Module Reference

| Module | Role | Key Procedures |
|---|---|---|
| `modConfig` | Core config type and all loading/lookup helpers. **Import first.** | `LoadConfig`, `GetGroupForCode`, `MatchesAnyPattern`, `GroupIndex` |
| `modSheetUtils` | Worksheet utilities: sheet naming, section detection, merge safety, row insert helper | `InsertRowCopyFromSource`, `IsSectionBar`, `ParseMonthSheet`, `MonthSheetName`, `SheetExists` |
| `modReaders` | All external file reading — pure data extraction, no sheet writes | `ReadYardi`, `ReadYardiMTM`, `ReadUnitStats`, `ReadRP`, `ReadUnitRentsGrid`, `ReadMovein`, `PickFile` |
| `modImport` | Import button handler and orchestration; calls Readers then writes to month sheet | `ImportMonthlyData`, `DoImport`, `FillSheet`, `ResolveMonthSheet` |
| `modDynamic` | Live buffer row insertion via `Workbook_SheetChange` | `HandleSheetChange` |
| `modSetup` | Property Setup sheet creation and month sheet generation | `CreateSetupSheet`, `GenerateMonthSheets` |
| `modOverview` | Builds and refreshes the multi-year renewal summary | `CreateOverviewSheet`, `RefreshOverview`, `FindOverviewName` |
| `modAdmin` | One-time setup wizard and health check; also exposes the shared `AddButton` helper used by both its own setup buttons and `modMTM.EnsureMTMButtons` | `SetupWorkbook`, `HealthCheck`, `AddButton` |
| `modMTM` | MTM tracker sheet refresh and checkbox-based import to month sheets (see [MTM Tracker Workflow](#mtm-tracker-workflow)) | `RefreshMTMSheet`, `ImportSelectedMTM`, `ClearSelectionMTM`, `SortMTMSheet`, `EnsureMTMButtons` |

---

## Troubleshooting

### "No 'Property Setup' sheet found"
Run **Create Setup Sheet** first, fill in all required fields, then re-run the failing action.

### "Re-run CreateSetupSheet to repair it"
A named range on the Property Setup sheet is missing or broken. Click **Create Setup Sheet** — if you choose "Yes" to replace it, re-enter your configuration. If you choose "No", check the VBA Names manager (Formulas → Name Manager) for broken `PS.*` names.

### Units were skipped — "Unmapped Yardi Codes"
The import found unit type codes in the Rent Roll that are not in your Yardi Code Map. Add the missing codes (and their floor plan group) to the Property Setup sheet, then click **Import Monthly Data** again for the same month.

### Columns F, N, or T are blank after import
- **Col F / T blank**: Neither the Unit Rents Grid nor the RP Renewal Offer Analysis was selected during import. Re-run the import and select at least one of these files.
- **Col N blank**: The Unit Rents Grid was not selected, or no matching unit number was found in the grid. Verify the unit number format matches between the Rent Roll and the Grid.

### "VBA Access Needed" during SetupWorkbook
Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings** and enable "Trust access to the VBA project object model". Re-run `SetupWorkbook`.

### Import crashes with "merged cell" error
This was a known bug fixed in v2.0.0. Ensure you are running version 2.1.0 or later (check the Health Check output). If the error recurs on a specific sheet, the sheet may have manually created multi-row merged cells outside the standard layout. Unmerge them manually and re-run the import.

### Overview shows blank cells for a month
The month sheet exists but its stats named ranges were not created (e.g., the sheet was manually added rather than generated). Delete the sheet and regenerate it using **Generate Month Sheets**, then re-import that month's data.

### Dynamic row insertion is not working
Verify that `ThisWorkbook` contains the `Workbook_SheetChange` event calling `modDynamic.HandleSheetChange` (see [Step 3](#step-3--wire-up-the-change-event)). Also confirm macros are enabled — events do not fire when macros are disabled.

### "The Move-in Box Score file contains VBA macros and was not loaded"
The `.xls` file you selected contains embedded VBA code, which is not present in a legitimate RealPage export. Re-export the Move-in Box Score directly from RealPage and try again. Do not use a file that has been modified or re-saved in Excel.

### "Syntax error" when importing modMTM.bas into the VBA Editor
If **File → Import File** rejects `modMTM.bas` with a bare "Syntax error" and no line number, the leading (unproven but easy to test) theory is a stale or conflicting `modMTM` component already sitting in the workbook's VBA project — for example, a partially-imported module left over from a previous attempt, or a copy under a slightly different internal state than the Project Explorer shows. Before re-importing:

1. Open the VBA Editor (**Alt + F11**).
2. In the Project Explorer, right-click any existing `modMTM` under **Modules** and choose **Remove modMTM** (if prompted to export first, you can skip it — you're re-importing the file from disk anyway).
3. Save the workbook, close and reopen Excel (a fresh VBA project state rules out any stale in-memory component).
4. Re-run **File → Import File** and select `modMTM.bas` again.

If the error still recurs after a clean removal and a fresh Excel session, that points back to a source-level defect rather than a stale-module conflict — in that case, check the file for anything the [structural hygiene checklist](#module-reference) below is meant to catch (non-ASCII characters, a BOM, mixed line endings, a missing trailing newline, or a blank line immediately before a declaration) using a hex-capable editor, since Notepad/Wordpad-style editors can hide these.

---

## Version History

| Version | Date | Changes |
|---|---|---|
| 2.7.0 | 2026-07-07 | Removed the "Import Target Month" dropdown feature (`SetupTargetMonthControl`, `TargetMonthEnd`, `EndOfMonth`, the N1/P1:P24 cells) after it caused a persistent, unresolved Import "Syntax Error" across 7 sessions — Status is now always computed against today+90 days, inlined at both call sites. Added alternating row banding (`ApplyRowBanding`) and fixed the MTM tracker header font color. Tested clean end-to-end in real Excel (import, compile, refresh, sort, import). |
| 2.6.0 | 2026-07-06 | Clean structural rewrite of all 9 modules (ASCII-only source, no BOM, consistent line endings, one trailing newline, standardized `Next<Noun>:` loop-continue labels, balanced-block self-check on every file). `modAdmin.SetupWorkbook` now includes `modMTM` in its module-presence check and wires the 4 MTM tracker buttons directly onto the `MTM` sheet if it already exists; a new `modMTM.EnsureMTMButtons` helper wires them automatically the first time `RefreshMTMSheet` creates the sheet, so no manual button setup is ever required. Fixed a real bug in `modDynamic`: the module-level Buffer Rows cache was only invalidated by an explicit call from `SetupWorkbook`, contradicting the documented "takes effect immediately" behavior — the cache is removed entirely and `GetBufferRows` now reads `PS.BufferRows` fresh every time. `modImport.FillSheet`'s inlined 4-step row-insert sequence now calls the shared `modSheetUtils.InsertRowCopyFromSource` helper instead of duplicating it. `modReaders.ReadYardiMTM` now returns a `Dictionary` of `MTMUnitRec` records (Unit, Name, FloorPlanCode, MarketRent, ActualRent, ExpiryVal, StaleOut) instead of a positional `Array(...)`, and `modImport.FillMTMRows` / `modMTM.WriteMTMDataRow` read named fields off the record instead of guessing array indices. Replaced a runtime `ChrW(8212)` em-dash in `modMTM`'s auto-note with a plain `" - "` string. Fixed a stray mis-encoded em-dash in a `modDynamic` comment that a prior encoding-corruption cleanup pass had missed. |
| 2.5.0 | 2026-07-03 | MTM Tracker reverted to a single flat, single-sorted list (no more Active MTM / Short Term sections or divider row); removed the RealPage and Resident Lease Expirations report prompts and short-term-lease detection entirely — back to Yardi Rent Roll only; Status (col I) is now a purely informational `Eligible`/`Not Yet Eligible` label with no effect on checkbox availability or import behavior; Next Increase now computed as Last Increase + 1 year + 1 day; added a 15-month-staleness auto-note in col J; `SortMTMSheet` now also recalculates Next Increase/Status from manual edits before sorting; tracker tab renamed back `MTM & STL` → `MTM` with automatic self-healing migration for existing workbooks. |
| 2.4.0 | 2026-07-03 | MTM Tracker sheet rebuilt as two physically separated, independently-sorted sections (Active MTM, then a divider bar, then Short Term) instead of one flat sorted list; removed the `⚠ Review - may have renewed` status entirely — units that drop out of the current scan are now silently removed instead of flagged; removed `SelectAllMTM` and hardened `ImportSelectedMTM`/`ClearSelectionMTM` against section-bar rows; tracker tab renamed `MTM` → `MTM & STL` with automatic self-healing migration for existing workbooks. |
| 2.3.0 | 2026-07-03 | MTM tracker carries Name/Floor Plan/Current Rent on import; RealPage report enriches MTM rows; MTM tag highlighted instead of column A; short-term-lease detection via Resident Lease Expirations report (RealPage fallback). |
| 2.2.0 | 2026-07-03 | Added modMTM: MTM tracker with checkbox import, FillMTMRows auto-fill. |
| 2.1.1 | 2026-06-22 | Bug fixes and hardening: fixed workbook resource leak when import errors mid-run; replaced `Integer` with `Long` for month/year variables to prevent overflow; added missing column-0 guard in `ReadUnitRentsGrid`; fixed import prompt column labels (L/X were mislabelled K/W); added `HasVBProject` security check on Move-in Box Score; removed stale buffer-row cache so Setup sheet changes take effect immediately; replaced magic-number literals with named constants; removed full-sheet font assignments that caused slow month sheet generation. |
| 2.1.0 | 2026-06-21 | Refactored from 3 monolithic modules into 8 focused modules. `modConfig`, `modSheetUtils`, `modReaders`, `modImport`, `modDynamic`, `modSetup`, `modOverview`, `modAdmin` separated for maintainability. |
| 2.0.0 | — | Fixed merged-cell crash during row insertion. 4-step unmerge/insert/format/formula sequence consolidated in `InsertRowCopyFromSource`. |
| 1.2.1 | — | Previous monolithic release: `modRenewalImporter`, `modRenewalDynamic`, `modPropertySetup`. |
