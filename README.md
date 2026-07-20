# AHT Capacity Dashboard (Google Apps Script)

A live capacity-planning dashboard for the disputes/claims team. It pulls data
straight from three Google Sheets and calculates a capacity plan **per employee
(reporter)** and **for the team**, on the assumption of **7 productive hours per
employee per working day**.

Recommended allocations are editable on the dashboard, and saving writes the
overrides back into the **Rota** spreadsheet — so the rota stays the single
source of truth.

---

## What it does

| Source | Sheet / tab | Used for |
|---|---|---|
| **AHT analysis** | `AHT Validation` tab, **column E** (`new AHT`) | Minutes per case, per scheme (averaged across entities); column A (`Reconciliation`) + column B (`Scheme`) also supply the entity+scheme → plain-scheme lookup used to translate Rota's rows |
| **Rota** | `Rota` tab | Who is working, which day, which scheme → available capacity + real scheme rostering |
| **Backlog** | `Scheme_View` tab | Open cases per scheme (`Grand Total` column) → demand |

**Data quirks this script accounts for:**
- `AHT Validation` has one row per **entity+scheme** (reconciliation key, e.g. `US-VISA`), not one row per scheme — AHT is averaged across all entities sharing a scheme.
- `Rota` is scheme-first: each row is one reconciliation key, and the reviewer covering it on a given weekday sits *inside* that day's cell (Mon–Fri columns), not in a per-reviewer row. Its layout is hard-coded in `CONFIG.rota.matrix` rather than header-detected, because it's too irregular for generic detection — see `probeRota()`.
- `Scheme_View` is a Google Sheets pivot table with no entity breakdown (its own note reads `**refine by entity`), so demand is computed at the plain-scheme level; Rota's entity-level rows are rolled up to match via the AHT-derived lookup above.
- Each reviewer is typically rostered for **many** schemes at once (all `DAILY`), not one — recommendations only ever assign a reviewer work within schemes they're actually rostered for.

**Calculations**

- **Capacity (per employee)** = working days in range × 7 hrs (× 60 = minutes).
- **Demand (per scheme)** = backlog cases (`Grand Total`) × new AHT (minutes), scheme name matched via the Reconciliation → Scheme lookup.
- **Team capacity** = sum of every rostered employee's capacity.
- **Utilisation** = total demand ÷ team capacity.
- **Recommended allocation** — two selectable strategies, both scoped to each reporter's real scheme rostering (never assigns a scheme a reporter doesn't actually cover):
  - *Proportional* (default): each reporter's own capacity is split across just the schemes they're rostered for, weighted by how backlogged each one is.
  - *Greedy*: clear the biggest backlog first, filling each scheme from the reporters rostered for it (biggest remaining capacity first) before moving to the next — produces focused, fewer-scheme assignments per person.
  Any manual override saved from the dashboard takes precedence under either.

**Charts**: demand-vs-capacity by scheme, capacity per employee, and a
**planned-allocation-by-employee** chart (recommended hours per person, stacked
and coloured by scheme).

**Filters** (one row above the charts): date from/to (+ Today / 7d / 30d / MTD /
All presets), reporter (employee), scheme (multi-select), and allocation
strategy.

**Theme** — light, Checkout-branded (indigo accent `--brand`, easily swapped to
your exact pantone). Chart series use a colour-blind-validated categorical
palette; every bar is directly labelled and a full table view is always present.

---

## Deploy

You need edit access to all three source sheets (the web app runs **as the
deploying user**, so their Google account must be able to read them).

### Option A — Apps Script editor (quickest)

1. Go to <https://script.google.com> → **New project**.
2. Create three files and paste the repo contents:
   - `Code.gs` → `Code.gs`
   - `Index.html` → **File ▸ New ▸ HTML** named `Index`
   - Project settings ▸ tick **"Show appsscript.json"**, then paste `appsscript.json`.
3. **Deploy ▸ New deployment ▸ Web app.** Execute as *me*, access as you prefer
   (Domain recommended). Authorise the requested Sheets scopes.
4. Open the web-app URL.

### Option B — clasp (version-controlled)

```bash
npm i -g @google/clasp
clasp login
clasp create --type webapp --title "AHT Capacity Dashboard" --rootDir .
# copy the generated scriptId into .clasp.json (see .clasp.json.example)
clasp push
clasp deploy
```

---

## Adjusting to the real column layouts

**AHT and Backlog** are simple flat tables — the script **auto-detects their
columns by header name** (except AHT column E, as specified). If something
looks wrong:

1. In the Apps Script editor, select **`diagnostics`** from the function
   dropdown and click **Run**.
2. Open **View ▸ Logs**. For AHT/Backlog you'll see the detected headers, which
   column was matched to each field, and a sample row.
3. If a field wasn't matched, add the real header text to the relevant
   `columns` candidate list in the `CONFIG` block at the top of `Code.gs`.

**Rota** is a bespoke scheme-first matrix (see "Data quirks" above) — its
column positions are hard-coded in `CONFIG.rota.matrix` rather than
header-detected, since its layout is too irregular for generic detection.
`diagnostics()` sanity-checks it by running the real reader and reporting the
entry count, reviewer list, and a sample. If the sheet's structure changes:

1. Run **`probeRota()`** and check **View ▸ Logs** for the raw row layout.
2. Adjust `CONFIG.rota.matrix` (`dateHeaderRow`, `dataStartRow`, `priorityCol`,
   `frequencyCol`, `reconciliationKeyCol`, `dayCols`) to match.

Two more general-purpose editor helpers exist if a tab or sub-tab layout ever
needs re-investigating: **`listTabNames()`** (lists every real tab name in all
three spreadsheets) and **`probeRawRows_`**'s other wrappers (`probeBacklog()`,
`probeSchemeView()`, `probeBacklogAllocation()`) for dumping raw rows of a
specific tab.

---

## Write-back

Editing an "Allocated cases" value and clicking **Save to Rota** upserts rows
into a **`Dashboard Allocations`** tab inside the Rota spreadsheet
(created automatically), keyed on (date, reporter, scheme), stamped with the
editor's email and timestamp. The raw `Rota` tab is never overwritten; these
overrides are re-read on the next refresh and shown as recommendations.

---

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Server side: readers, capacity model, write-back, `diagnostics()` |
| `Index.html` | Dashboard UI (filters, KPIs, charts, editable table) |
| `appsscript.json` | Manifest (web-app config + OAuth scopes) |
| `preview/mock-dashboard.html` | Standalone visual preview with sample data (no Google account needed) |
