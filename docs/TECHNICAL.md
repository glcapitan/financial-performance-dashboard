# Technical Documentation — Financial Performance Dashboard

> Full build, architecture, and maintenance reference. For the project overview, see the [main README](../README.md).

> A one-click FP&A dashboard built in Excel with **Power Query**, **Power Pivot**, and **DAX**. Change a single date cell, hit *Refresh All*, and every KPI, variance, and chart updates — Revenue, Gross Profit, and EBITDA against budget, prior period, and prior year.

![Excel](https://img.shields.io/badge/Built%20with-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/ETL-Power%20Query-377C2B)
![Power Pivot](https://img.shields.io/badge/Data%20Model-Power%20Pivot-1F3864)
![DAX](https://img.shields.io/badge/Measures-DAX-2E75B6)
![Platform](https://img.shields.io/badge/Platform-Windows%20%2F%20Microsoft%20365-blue)

A monthly management-reporting pack. Four departmental income statements (and their budget twins) are unpivoted in Power Query, loaded into the Power Pivot data model, related through lookup tables, and summarised with DAX measures. The Dashboard then reads those measures through pivot tables and renders KPI tiles, donut gauges, and charts. The only thing the user changes is **one input cell**.

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [How It Was Built (Methodology)](#how-it-was-built-methodology)
- [Architecture & Data Model](#architecture--data-model)
- [How the Engine Works](#how-the-engine-works)
- [The Single Input & Reporting Windows](#the-single-input--reporting-windows)
- [Workbook Structure](#workbook-structure)
- [Under the Hood: Measures & Conditional Charts](#under-the-hood-measures--conditional-charts)
- [Example Output (Sample Data)](#example-output-sample-data)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Known Issues / Model Integrity](#known-issues--model-integrity)
- [Roadmap](#roadmap)
- [Acknowledgements](#acknowledgements)
- [Disclaimer](#disclaimer)

---

## Overview

The Financial Performance Dashboard answers three questions every month, for Revenue, Gross Profit, and EBITDA:

1. **How did we do versus the budget?** — Budget vs Actual (BvA)
2. **How did we do versus last period?** — Actual vs Prior Period (AvPP)
3. **How did we do versus last year?** — Actual vs Prior Year (AvPY)

It then decomposes operating spend two ways — **by department** (Customer Support, G&A, R&D, Sales & Marketing) and **by cost type** (Payroll, Advertising, Professional Fees, Other G&A) — so any variance can be traced to a responsible owner and a spend category.

The headline design goal is *one-click reporting*: department owners fill in their own actual and budget tabs, those tabs are consolidated through Power Query, and a single "Current Month" cell drives which reporting window the entire dashboard displays.

---

## Key Features

- **One input, full refresh** — change the `Current Month` cell, choose a period on the slicer, and *Refresh All*; every tile, table, and chart recalculates.
- **Power Query ETL** — each department tab is unpivoted from a human-readable grid into a clean columnar table and appended into a single fact table for actuals and one for budget.
- **Power Pivot data model** — fact tables related to lookup tables (date, account mapping, departments) for scalable, relationship-driven reporting.
- **DAX measures** — period logic (current month, trailing 3/6/12, YTD, and prior-period / prior-year equivalents) handled in measures rather than fragile cell formulas.
- **Three benchmark comparisons** with `$` and `%` variances and automatic "Hit" / "Miss" labels.
- **Conditional donut gauges** — favourable variances render green, unfavourable render red, with a grey remainder.
- **Slicer-driven** — re-slice by Department, Section, and Period without touching a formula.

---

## How It Was Built (Methodology)

The workbook follows a standard Power Query → Power Pivot → DAX pipeline. The high-level build sequence:

1. **Design first.** A colour scheme, fonts, and section layout are laid out on the `Dashboard` sheet before any data is wired in (KPI boxes, chart placeholders, variance tables).
2. **Name the source ranges.** Every actual and budget tab is turned into a named range (e.g. `GnA`, `GnA_Budget`) scoped to its own sheet.
3. **Extract & transform in Power Query.** Each named range is loaded, the first row promoted to headers, the blank column dropped, and the date columns **unpivoted** into a tidy `Account | Date | Amount` shape. A `Department` column is added so rows can be filtered/aggregated later.
4. **Append.** The four actual queries are appended into **Departmental P&L**; the four budget queries into **Departmental P&L Budget**.
5. **Load to the data model.** The two appended queries are added to **Power Pivot** (connection-only; the source queries are not loaded to sheets).
6. **Build lookup tables & relationships.** A date table is generated, and mapping/department tables are loaded, then related to the fact tables (see below).
7. **Write DAX measures.** Period and variance measures are authored in the data model.
8. **Pivot + present.** Pivot tables on the `Data` sheet expose the measures; the `Dashboard` pulls values from those pivot cells into shapes, KPI text boxes, and pivot charts.

---

## Architecture & Data Model

Three clean tiers:

```
INPUTS                      ENGINE                          OUTPUT
8 P&L tabs        →   Power Query (unpivot, append)   →   Pivot tables
(4 depts ×        →   Power Pivot data model          →   on the Data sheet
 Actual/Budget)        + DAX measures                  →   Dashboard
                       + lookup tables                      (tiles, charts)
```

### Data model tables

| Table | Kind | Purpose |
|-------|------|---------|
| `Departmental P&L` | Fact | All actuals, unpivoted and appended across the four departments. |
| `Departmental P&L Budget` | Fact | All budget figures, same shape. |
| `Date` (calendar) | Lookup | Auto-generated date table; one row per day across the data range. |
| `Mappings` | Lookup | Account → **Section** → **Summary Grouping** → **Sign** (`+1` income / `−1` expense). |
| `Departments` | Lookup | The four departments. |
| `Table View` (periods) | Helper | Feeds the period slicer (no relationship — wired via measures). |
| `Predefined Dates` | Helper | Pre-computed start/end dates for every window, referenced by measures. |

### Relationships

- `Date[Date]` → both fact tables on **Date**
- `Mappings[Account]` → both fact tables on **Account**
- `Departments[Department]` → both fact tables on **Department**

Because the lookup tables sit in the middle, selecting any date, section, summary grouping, or department automatically filters both fact tables consistently.

---

## How the Engine Works

There are really **two layers** worth understanding:

### 1. The sample-data layer (on the P&L tabs)

Each line item on a department tab is projected forward from a seed value by a single monthly growth rate:

```
line(month) = line(prior month) × (1 + monthly growth rate)
```

- **Column C** = growth rate (the assumption)
- **Column E** = seed value (first month)
- **Columns F →** = `=PriorColumn*(1+$C)`

This is how the included sample P&L is generated. In a live deployment you would replace these projected values with actuals exported from your accounting system — the structure downstream does not change.

### 2. The reporting layer (Power Pivot + DAX)

Power Query reshapes the tabs; the data model relates them; DAX measures do the period maths and variance logic. The Dashboard never references the department tabs directly — it reads measures via pivot tables, which keeps the presentation layer insulated from the raw data.

### Metric definitions

| Metric | Definition |
|--------|------------|
| **Revenue** | Total billings across all departments (sign `+1`). |
| **Gross Profit** | Revenue less Cost of Goods Sold (hosting / merchant / web-domain fees). |
| **EBITDA** | Gross Profit less total operating expenses (labelled *Net Operating Income* on the tabs). |
| **Net Income** | EBITDA plus Net Other Income (dominated by interest and depreciation). |

---

## The Single Input & Reporting Windows

The **only** input the user needs to touch is one blue cell — `Current Month` (a workbook-scoped named range, e.g. `7/31/2024`). A `Predefined Dates` table derives every other date from it with `EOMONTH` / `DATE` formulas, so changing that one cell re-points the whole model.

Available reporting windows (chosen on the **Period** slicer), each with matched **Prior Period** and **Prior Year** equivalents:

- Actual full years (2023 / 2024 / 2025)
- Current Month
- Trailing 3 Months
- Trailing 6 Months
- Trailing 12 Months
- Year-to-Date (YTD)

The figures published below use the **Trailing-12-Month window ending July 2024**.

---

## Workbook Structure

| Sheet | Role | Type |
|-------|------|------|
| `Dashboard` | Presentation layer: slicer-driven KPI tiles, donut gauges, and variance tables. | Output |
| `Data` | Hosts the pivot tables that expose the DAX measures; feeds the Dashboard. | Calculation |
| `List` | Reference layer: account mapping, department/metric/version lists, period table, predefined dates, and the `Current Month` input. | Config |
| `G&A`, `R&D`, `CS`, `S&M` | Actual monthly income statements for the four departments. | Input |
| `G&A_Budget`, `R&D_Budget`, `CS_Budget`, `S&M_Budget` | Budget-version income statements. | Input |

> The Power Pivot **data model** and **Power Query queries** are not worksheets — they live inside the workbook and are edited via *Data → Manage Data Model* and *Data → Queries & Connections*.

Eight P&L sheets in total (4 departments × Actual/Budget), each running **36 monthly columns from January 2023 through December 2025**.

---

## Under the Hood: Measures & Conditional Charts

### Period selection via measures

A `Selected View` measure reads the period slicer with `VALUES(...)`; an `Actual` measure then uses a `SWITCH` to return the matching period measure (e.g. *Trailing 12 Months → Trailing 12 Mo Values*). This is what lets the slicer drive the numbers without any relationship on the period table.

### Variance sign convention

Two variance flavours keep "good" and "bad" intuitive:

- **Income accounts:** `Actual − Budget` → positive is favourable.
- **Expense accounts:** `Budget − Actual` → positive is an under-spend (favourable), negative is an over-spend (unfavourable).

### Donut gauges (green / red)

Each gauge is driven by a small helper block with four series — `A(+) / B(+) / A(−) / B(−)` — where `A` is the value and `B` is its inverse (`1 − value`), clamped to `[0, 1]` with `MIN`/`MAX`. A favourable variance fills the `A(+)` arc **green**; an unfavourable one fills `A(−)` **red**; the `B` remainder is **grey**. The hole is set to ~90% so it reads as a gauge.

---

## Example Output (Sample Data)

> These figures come from the **sample dataset** shipped with the template (generated by the growth-driver tabs). They illustrate what the dashboard surfaces; they will change the moment you load your own actuals.

| Metric | Actual | Budget | Δ vs Budget | vs Prior Year |
|--------|-------:|-------:|------------:|--------------:|
| Revenue | 801,118 | 806,240 | (5,122) | +12.1% |
| Gross Profit | 659,664 | 782,846 | (123,182) | +18.1% |
| Gross Margin | 82.3% | 97.1% | −14.8 pts | — |
| Operating Expenses | 714,150 | 942,212 | 228,062 ✅ | — |
| EBITDA | (54,486) | (159,366) | 104,879 ✅ | +51.1% |
| Net Income | (130,991) | — | — | — |

**What the sample tells the story-reader:**

1. **Revenue is on plan; the margin is not.** Revenue lands within ~1% of budget, but gross margin is 82.3% vs a budgeted 97.1% — ~15 points of compression that erases $123k of planned gross profit. Since revenue did not miss, the entire shortfall is a cost-of-goods problem.
2. **Cost discipline cushions the loss.** OpEx comes in $228k (24%) under budget, turning a budgeted EBITDA loss of $(159k) into $(54k).
3. **Still unprofitable.** After $(76k) of net other expense — mostly $74k of interest — net income is $(131k).

<details>
<summary><strong>OpEx variance — by cost type (sample)</strong></summary>

| Cost type | Budget | Actual | Under/(Over) | % |
|-----------|-------:|-------:|-------------:|--:|
| Advertising | 245,241 | 130,261 | 114,980 | 47% |
| Other G&A | 292,884 | 212,909 | 79,975 | 27% |
| Payroll | 327,740 | 264,464 | 63,276 | 19% |
| Professional Fees | 76,346 | 106,516 | (30,170) | (40%) |
| **Total OpEx** | **942,212** | **714,150** | **228,062** | **24%** |

Professional Fees is the only category over budget — a useful demo of how the dashboard flags a single hot line against a backdrop of under-spend.

</details>

---

## Getting Started

### Requirements

- **Microsoft Excel on Windows**, or **Microsoft 365** — Power Pivot is required.
- ⚠️ **Power Pivot is not available on Excel for Mac.** The file will open, but the data model, measures, and one-click refresh will not function. Use Windows/365.
- If Power Pivot is not visible: *File → Options → Add-ins → COM Add-ins → enable "Microsoft Power Pivot for Excel".*

### Open it

```bash
git clone https://github.com/<your-org>/financial-performance-dashboard.git
cd financial-performance-dashboard
# then open Financial_Performance_Dashboard.xlsx in Excel for Windows / Microsoft 365
```

---

## Usage

| Task | How |
|------|-----|
| **Re-point the reporting period** | Change the single `Current Month` input cell on the `List` sheet, then *Data → Refresh All*. |
| **Switch the window/benchmark** | Click the **Period** slicer (Current Month, Trailing 3/6/12, YTD, full years). |
| **Filter by department** | Use the Department slicer / timeline on the relevant view. |
| **Update the data** | Replace the projected values on the department tabs with your actuals/budget, then *Refresh All* so Power Query re-imports and the model recalculates. |
| **Add an account** | Add the line to the department tab **and** register it in the `Mappings` table (Section, Summary Grouping, Sign) so roll-ups pick it up. |
| **Add a department** | Add the tab + named range, build its Power Query query, append it to the fact query, and add it to the `Departments` lookup. |

> Always *Refresh All* before sharing so slicer-driven values and the "Hit / Miss" labels are current.

---

## Known Issues / Model Integrity

The model is well-structured and the major subtotals tie out (`Revenue − COGS = Gross Profit`; `Gross Profit − OpEx = EBITDA`). A few items to clean up before relying on the pack externally:

- [ ] **Inconsistent variance-% basis.** The Dashboard expresses opex variance as a `%` of **budget**, while the `Data`/measure layer expresses the same variance as a `%` of **actual**. Standardise on one denominator (variance ÷ budget is conventional).
- [ ] **Sales & Marketing variance % shows 0** in the by-department block despite a real ~$100k dollar variance — a broken/blank cell to repair.
- [ ] **Budget COGS assumption looks unrealistic.** A budgeted ~97% gross margin (≈3% cost of delivery) is the root of the headline "miss" in the sample data; re-baseline before judging performance.
- [ ] **Compounding drivers need a sanity check.** Some sample lines (merchant fees, interest) grow at a fixed 10%/month, which compounds aggressively over 36 months. These are sample assumptions, not real terms.
- [ ] **Mac users blocked.** Document the Windows/365 requirement prominently for collaborators (see Requirements).

**Scope limitations**

- Income-statement only — no balance sheet or cash-flow statement, so liquidity / runway cannot be read directly.
- The shipped figures are model-generated sample data, not a live accounting feed.
- No currency or rounding convention is declared in the file.

---

## Roadmap

- [ ] Wire the fact tables to a real GL export (e.g. QuickBooks Online) so the model refreshes from source.
- [ ] Standardise variance-% conventions across the Dashboard and measures.
- [ ] Move all driver assumptions onto a dedicated, documented Assumptions tab.
- [ ] Add a break-even bridge projecting when EBITDA and net income cross zero.
- [ ] Add a balance sheet and cash-flow statement for a full three-statement view.

---

## Acknowledgements

Built by following a step-by-step YouTube tutorial on creating a one-click FP&A dashboard in Excel with Power Query and Power Pivot. _Add the video link and creator credit here._ The template-store / plugin referenced in that tutorial is Model Wiz (modelwiz.com).

---

## License

Released under the [MIT License](LICENSE) — free to use, modify, and distribute with attribution. Update the copyright holder in the `LICENSE` file before publishing.

---

## Disclaimer

This workbook is a planning and analysis tool. The included figures are driver-generated sample data — not audited financials — and should not be used as a substitute for statements prepared under an applicable accounting framework. Validate all assumptions and load your own data before relying on outputs for decision-making.
