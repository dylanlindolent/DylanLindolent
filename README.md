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
| **AHT analysis** | `AHT Validation` tab, **column E** (`new AHT`) | Minutes per case, per scheme (and per agent if present) |
| **Rota** | `Rota` tab | Who is working, which day → available capacity |
| **Backlog** | `backlog team planning` tab | Open cases per scheme → demand |

**Calculations**

- **Capacity (per employee)** = working days in range × 7 hrs (× 60 = minutes).
  If the rota has an explicit hours column it is used instead of the 7-hr default.
- **Demand (per scheme)** = backlog cases × new AHT (minutes).
- **Team capacity** = sum of every rostered employee's capacity.
- **Utilisation** = total demand ÷ team capacity.
- **Recommended allocation** = demand-weighted split of team capacity across
  schemes, then split across employees in proportion to each person's capacity.
  Any manual override saved from the dashboard takes precedence.

**Filters** (one row above the charts): date from/to (+ Today / 7d / 30d / MTD /
All presets), reporter (employee), and scheme (multi-select).

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

The three sheets are maintained by different people and their columns may not
match the defaults. The script **auto-detects columns by header name** — it does
not rely on fixed positions (except AHT column E, as specified).

If something looks wrong:

1. In the Apps Script editor, select **`diagnostics`** from the function
   dropdown and click **Run**.
2. Open **View ▸ Logs**. You'll see, for each sheet, the detected headers, which
   column was matched to each field, and a sample row.
3. If a field wasn't matched, add the real header text to the relevant
   `columns` candidate list in the `CONFIG` block at the top of `Code.gs`.

`CONFIG` also holds: the sheet IDs and tab names, the AHT column letter/unit,
the list of rota statuses that count as "not working", and the 7-hr productivity
assumption — all in one place.

**Rota layouts** — both are supported and auto-detected:
- *Long*: one row per employee/day with a `status` column.
- *Matrix*: employee down the side, dates across the top, shift/status in cells.

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
