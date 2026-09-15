# SLA Dashboard — Format 1

A customer-facing React/Vite dashboard that accepts Excel (`.xlsx`, `.xls`) or CSV files, calculates the derived SLA fields at runtime, and renders the delivery-performance dashboard.

## Run locally

Requirements:
- Node.js 18+ (20+ recommended)

```bash
npm install
npm run dev
```

Then open the localhost URL printed by Vite (normally `http://localhost:5173`).

## Upload conventions

The application does **not** store calculated results in the source code. It stores only the calculation rules.

It detects:
- `Outbound` or `FG` → Finished Goods / outbound quantity source
- `LR` → LR source
- `REPL` is ignored for Format 1 because the supplied Summary is based on FG/Outbound + LR.

You can upload:
- one Excel workbook containing Outbound/FG + LR sheets, or
- separate CSV/XLSX files.

Column matching is tolerant of spaces, case, underscores and common aliases.

## Runtime calculations

### 1. TAT
For each record:

**FTL**
`TAT = ROUNDUP(Distance / 300, 0.1)`

**PTL**
`TAT = ROUNDUP(Distance / 100 + 1, 0.1)`

### 2. Lead-time fields
Using Excel-style date differences in calendar days:
- OBD → Delivery = Actual Delivery Date − OBD Date
- Invoice → Delivery = Actual Delivery Date − Invoice Date
- Invoice → Reporting = Actual Reporting Date at CP End − Invoice Date
- OBD → Invoice = Invoice Date − OBD Date

Bucket rule:
- <= 0 → Same Day
- <= 1 → Next Day
- <= 3 → 2~3 Days
- <= 6 → 4~7 Days
- > 6 → 7+ Days

### 3. Timeline / Outbound TAT
The final Timeline classification is created only after TAT and Invoice → Reporting are calculated:

`IF(InvoiceToReportingDays <= TAT, "Within TAT", "Out TAT")`

Dependency chain:

`Dispatch Mode + Distance → TAT`

`Actual Reporting Date at CP End − Invoice Date → Invoice to Reporting Lead Time`

`Invoice to Reporting Lead Time + TAT → Timeline / Outbound TAT`

### 4. SLA
Format 1 uses the clarified business rule that **1–100 km is Local Delivery**:
- 1–100 km → Local Delivery
- 101–300 km → 101–300 km
- 301–600 km → 301–600 km
- 601–1000 km → 601–1000 km
- >1000 km → >1000 km

### 5. State
State is derived from the 6-digit Indian PIN code. Format 1 explicitly supports the four states represented by the supplied Kolkata report:
- West Bengal: 70xxxx–74xxxx
- Odisha: 75xxxx–77xxxx
- Bihar: 80xxxx and 84xxxx–85xxxx
- Jharkhand: 81xxxx–83xxxx

More states can be added to `src/lib/calculations.js` without changing the dashboard UI.

### 6. Dashboard measures
**FG / Outbound**
- Invoice Qty = sum of Invoice Qty
- TAT analysis is quantity-weighted by Invoice Qty

**LR**
- LR = count of **unique LR No.**, never row count
- TAT/SLA/State LR analysis uses unique LR numbers

### 7. Source context
The dashboard attempts to identify:
- location/customer/report name
- month
- year

from the workbook/file name and sheet name. It also displays the exact calculated source sheet (`Outbound`, `FG`, or `LR`) in the dashboard header.

## Important implementation note

The supplied workbook has a few legacy naming/formula inconsistencies. Format 1 intentionally follows the workbook's actual formula for TAT status while using your clarified SLA rule (Local = 1–100 km).

No workbook totals are hardcoded into the app.


## Validation against Kolkata May 2026 source/solution

The raw `Kolkata May 2026.xlsx` was checked against the provided solution workbook.

A key finding: the solution workbook's `Outbount TAT` column is not consistently formula-driven. 63 Outbound rows contain hardcoded `Within TAT` values even though their visible formula logic would classify them differently. The application intentionally follows the stated calculation method rather than copying those hardcoded values.

The solution workbook also confirms that `TAT` values are integer ceilings despite the formula being written with `ROUNDUP(...,0.1)`. For example, 1.3 becomes stored as 2. The app reproduces the resulting stored TAT values.

The dashboard's top SLA chart now displays numeric data labels directly above each bar.

TAT-wise performance is now shown with TAT categories as rows (TAT 1 through TAT 8) for both Invoice Qty and Unique LR.


## Kolkata reconciliation notes
The source solution uses Local Delivery for distances up to 70 km (not 100 km), with 71–299 km classified as `< 300 Km`. PIN classification also contains the source-specific 811–813 Bihar / 814–816 Jharkhand split and two explicit PIN exceptions (743437 and 721659).

The supplied solution workbook contains manually overridden Outbound TAT statuses in 66 rows. Those overrides change the Summary from the raw formula result (4,651 Out / 14,872 Within) to the displayed Summary (2,997 Out / 16,526 Within). They cannot be reconstructed from the raw file alone because no corresponding rule/field exists in the raw source. The dashboard therefore keeps the formula-driven result rather than hardcoding row-specific exceptions.


## TAT-wise performance layout (final)
The TAT-wise section is a single table. Rows are **Out TAT**, **Within TAT**, and **Grand Total**. Columns are **Sum of Invoice Qty** (from Outbound/FG) and **No. of LR No.s** (unique LR count from LR). There are no TAT 1/TAT 2 rows in this section.


## TAT-wise layout
The TAT-wise section intentionally mirrors the supplied MIS: **two separate tables side by side** — Qty Wise and LR Wise. Each table has Row Labels (Out TAT, Within TAT, Grand Total), its measure column, Delivery TAT %, and the Within TAT footer percentage.
