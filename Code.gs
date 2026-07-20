/**
 * AHT Capacity Planning Dashboard  —  server side (Google Apps Script)
 * ---------------------------------------------------------------------
 * Pulls live data from three Google Sheets and computes a capacity plan
 * per employee (reporter) and for the team:
 *
 *   1. AHT analysis -> "AHT Validation" tab. Column A = Reconciliation key
 *      (entity+scheme, e.g. "US-VISA"), column B = Scheme, column E = "new
 *      AHT" (minutes/case). Used both for AHT and as the lookup that
 *      translates Rota's reconciliation keys into plain scheme names.
 *   2. Rota -> "Rota" tab. Scheme-first matrix: each row is one
 *      reconciliation key, and the reviewer covering it on a given weekday
 *      sits inside that day's cell (column layout is hard-coded in
 *      CONFIG.rota.matrix — see readRota_ — because this sheet's layout is
 *      too irregular for generic header-based detection).
 *   3. Backlog -> "Scheme_View" tab (a Google Sheets pivot). Real header is
 *      row 2; "Grand Total" is the open-case count per scheme. This pivot is
 *      scheme-only (no entity breakdown), so demand is computed at the plain
 *      scheme level and Rota's entity-level rows are rolled up to match.
 *
 * Core assumption: 7 productive hours per employee per working day.
 *
 * AHT and Backlog columns are still detected by header name (candidate lists
 * in CONFIG) since those two tabs are simple flat tables. If a layout
 * differs, run `diagnostics()` from the editor (View > Logs) to see the
 * detected headers and adjust the candidate lists below. Rota's layout is
 * hard-coded (see above) — if that sheet's structure changes, re-run
 * `probeRota()` and adjust CONFIG.rota.matrix.
 *
 * Manual allocation changes made on the dashboard are written back into the
 * ROTA spreadsheet (into an "Dashboard Allocations" tab, created on demand),
 * so the rota remains the single source of truth for allocation.
 */

/* ============================ CONFIGURATION ============================ */

const CONFIG = {
  productivityHoursPerDay: 7,          // <- the "7hr productivity" assumption
  minutesPerHour: 60,

  aht: {
    sheetId: '15IKpN7QrvmewqfuNuhtsqWks-7ugvOVVNhv4j_qpPDI',
    tab: 'AHT Validation',
    headerRow: 1,
    // Column E is the "new AHT" per the brief. Kept explicit as requested.
    ahtColumnLetter: 'E',
    ahtUnit: 'minutes',                // 'minutes' | 'seconds' | 'hours'
    columns: {
      reconciliation: ['reconciliation', 'reconciliation key', 'recon key', 'recon'],
      scheme:   ['scheme', 'schemes', 'claim type', 'type', 'reason', 'reason code', 'category', 'queue', 'product'],
      reporter: ['reporter', 'employee', 'agent', 'analyst', 'name', 'owner'] // optional per-agent AHT
    }
  },

  rota: {
    sheetId: '1-K8fLgvU7h52ZYdEwEfZR5q3gYSyURjsCVJRlwB_K-g',
    tab: 'Rota',
    allocationsTab: 'Dashboard Allocations', // written back into the rota file
    // The Rota tab is scheme-first, not reviewer-first: each row is one
    // reconciliation key (entity+scheme, e.g. "US-VISA"), and the reviewer
    // covering it on a given weekday sits INSIDE that day's cell. Row 1 is
    // just "MON".."FRI" labels; row 2 holds the real per-day dates for the
    // first week block. A second Mon-Fri block exists further right but is
    // a stale/unused leftover (confirmed with the sheet owner) and is
    // intentionally not read. All positions are 0-based column indices
    // (column A = 0), confirmed against the real sheet via probeRota() —
    // re-run that and adjust here if the sheet's layout changes.
    matrix: {
      dateHeaderRow: 2,          // row with the real Mon-Fri dates
      dataStartRow: 3,           // first scheme row
      priorityCol: 1,            // column B
      frequencyCol: 2,           // column C
      reconciliationKeyCol: 3,   // column D
      dayCols: [4, 5, 6, 7, 8]   // columns E-I = Mon-Fri, current week only
    }
  },

  backlog: {
    sheetId: '1-ywS9-yFh0rJYwX421uADNZbr3UjgAWVcxkB1kfsQJE',
    tab: 'Scheme_View',
    // Real header is row 2 — row 1 just has "COUNTA of Reporter"/"Status"
    // labels left over from the pivot table this tab is built from.
    headerRow: 2,
    columns: {
      scheme: ['scheme', 'queue', 'claim type', 'type', 'category', 'product', 'reason'],
      // "Grand Total" (sum of the status-breakdown columns) is the open
      // backlog volume per scheme; matched here via the 'total' candidate.
      volume: ['backlog', 'volume', 'open', 'outstanding', 'wip', 'cases', 'count', 'tickets', 'items', 'total'],
      date:   ['date', 'week', 'as of', 'week commencing', 'w/c']
    }
  }
};

/* ============================ WEB APP ENTRY ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AHT Capacity Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================ PUBLIC API =============================== */
/* These are called from the client via google.script.run.               */

/**
 * Returns the option lists used to populate the dashboard filters.
 */
function getFilterOptions() {
  try {
    const aht = readAht_();
    const backlog = readBacklog_();
    const rota = readRota_();

    // Rota's own "scheme" values are raw entity-level reconciliation keys
    // (e.g. "US-VISA"), not plain scheme names, so they're excluded here —
    // AHT and Backlog already cover the full plain-scheme vocabulary.
    const schemes = uniqueSorted_(
      aht.map(function (r) { return r.scheme; })
        .concat(backlog.map(function (r) { return r.scheme; }))
    );
    const reporters = uniqueSorted_(rota.map(function (r) { return r.reporter; }));

    const dates = rota
      .map(function (r) { return r.date; })
      .filter(Boolean)
      .map(function (d) { return d.getTime(); });
    const range = dates.length
      ? { min: fmtDate_(new Date(Math.min.apply(null, dates))),
          max: fmtDate_(new Date(Math.max.apply(null, dates))) }
      : { min: '', max: '' };

    return { ok: true, schemes: schemes, reporters: reporters, dateRange: range };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

/**
 * Main data call. `filters` = { dateFrom, dateTo, schemes:[], reporter:'' }.
 * Returns everything the dashboard renders.
 */
function getDashboardData(filters) {
  filters = filters || {};
  try {
    const model = buildModel_(filters);
    return { ok: true, data: model, generatedAt: fmtDateTime_(new Date()) };
  } catch (err) {
    return { ok: false, error: String(err && err.stack || err) };
  }
}

/**
 * Persists one or more manual allocation overrides back to the ROTA sheet.
 * Each allocation = { date:'yyyy-MM-dd', reporter, scheme, cases, note }.
 */
function saveAllocations(allocations) {
  try {
    if (!allocations || !allocations.length) return { ok: true, saved: 0 };
    const ss = SpreadsheetApp.openById(CONFIG.rota.sheetId);
    let sheet = ss.getSheetByName(CONFIG.rota.allocationsTab);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.rota.allocationsTab);
      sheet.appendRow(['Date', 'Reporter', 'Scheme', 'Allocated cases', 'Note', 'Updated by', 'Updated at']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const user = Session.getActiveUser().getEmail() || 'unknown';
    const now = new Date();

    // Read existing rows so we upsert on (date, reporter, scheme).
    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 1, last - 1, 7).getValues() : [];
    const keyOf = function (d, rep, sch) {
      return [String(d), String(rep).toLowerCase(), String(sch).toLowerCase()].join('||');
    };
    const index = {};
    existing.forEach(function (row, i) {
      index[keyOf(fmtDate_(toDate_(row[0])), row[1], row[2])] = i + 2; // sheet row
    });

    let saved = 0;
    allocations.forEach(function (a) {
      const rowValues = [a.date, a.reporter, a.scheme, Number(a.cases) || 0, a.note || '', user, now];
      const k = keyOf(a.date, a.reporter, a.scheme);
      if (index[k]) {
        sheet.getRange(index[k], 1, 1, 7).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }
      saved++;
    });
    return { ok: true, saved: saved };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

/**
 * Editor helper: run this FIRST if diagnostics()/the dashboard errors with
 * "Tab not found". Lists every actual tab name in each of the three source
 * spreadsheets, so you can see exactly what to put in CONFIG.*.tab.
 */
function listTabNames() {
  const out = [];
  [['AHT', CONFIG.aht.sheetId], ['ROTA', CONFIG.rota.sheetId], ['BACKLOG', CONFIG.backlog.sheetId]]
    .forEach(function (pair) {
      const label = pair[0], id = pair[1];
      try {
        const ss = SpreadsheetApp.openById(id);
        const names = ss.getSheets().map(function (s) { return s.getName(); });
        out.push(label + '  file: "' + ss.getName() + '"');
        out.push('   tabs: ' + JSON.stringify(names));
      } catch (e) {
        out.push(label + '  ERROR opening ' + id + ': ' + e);
      }
    });
  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Editor helper: run from the Apps Script editor and open View > Logs to see
 * the headers each source sheet actually exposes, plus the columns detected.
 * If a tab isn't found, this also lists the real tab names available in that
 * spreadsheet (same info as listTabNames(), inline) so you can fix CONFIG in
 * one pass.
 */
function diagnostics() {
  const out = [];

  // AHT and Backlog are simple flat tables — use generic header detection.
  [['AHT', CONFIG.aht], ['BACKLOG', CONFIG.backlog]].forEach(function (pair) {
    const label = pair[0], cfg = pair[1];
    try {
      const t = readTable_(cfg.sheetId, cfg.tab, cfg.headerRow);
      out.push(label + ' [' + cfg.tab + ']  headers: ' + JSON.stringify(t.headers));
      out.push('   detected: ' + JSON.stringify(detectColumns_(t.headers, cfg.columns)));
      out.push('   rows: ' + t.rows.length + '  sample: ' + JSON.stringify(t.rows[0] || {}));
    } catch (e) {
      out.push(label + '  ERROR: ' + e);
      try {
        const ss = SpreadsheetApp.openById(cfg.sheetId);
        const names = ss.getSheets().map(function (s) { return s.getName(); });
        out.push('   real tabs in "' + ss.getName() + '": ' + JSON.stringify(names));
      } catch (e2) {
        out.push('   could not open spreadsheet ' + cfg.sheetId + ': ' + e2);
      }
    }
  });

  // Rota's layout is hard-coded (CONFIG.rota.matrix), not header-detected —
  // sanity-check it by running the real reader and reporting what it found.
  try {
    const rota = readRota_();
    const days = uniqueSorted_(rota.map(function (r) { return r.date ? fmtDate_(r.date) : ''; }));
    out.push('ROTA [' + CONFIG.rota.tab + ']  entries: ' + rota.length +
      '  reviewers: ' + JSON.stringify(uniqueSorted_(rota.map(function (r) { return r.reporter; }))) +
      '  days: ' + JSON.stringify(days));
    out.push('   sample: ' + JSON.stringify(rota[0] || {}));
  } catch (e) {
    out.push('ROTA  ERROR: ' + e);
  }

  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Editor helpers: dump the first N raw rows of the Rota / Backlog tabs
 * exactly as stored (no header assumption, no column detection). Use these
 * when diagnostics() shows headers/detected columns that don't look right —
 * e.g. a multi-row header or a pivot-table layout — so the real row-by-row
 * structure can be seen and CONFIG/readers adjusted to match it.
 */
function probeRawRows_(sheetId, tabName, n) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return 'Tab "' + tabName + '" not found in ' + sheetId;
  const rows = Math.min(n, sheet.getLastRow());
  const cols = sheet.getLastColumn();
  if (rows < 1 || cols < 1) return '(sheet is empty)';
  const values = sheet.getRange(1, 1, rows, cols).getValues();
  const msg = values.map(function (row, i) { return 'row ' + (i + 1) + ': ' + JSON.stringify(row); }).join('\n');
  Logger.log(msg);
  return msg;
}

function probeRota() {
  return probeRawRows_(CONFIG.rota.sheetId, CONFIG.rota.tab, 40);
}

function probeBacklog() {
  return probeRawRows_(CONFIG.backlog.sheetId, CONFIG.backlog.tab, 6);
}

/**
 * probeBacklog() showed "Backlog_Team Planning" is a pivot table with no
 * scheme column — these two probe candidate tabs in the same spreadsheet
 * that sound more likely to hold per-scheme backlog volume.
 */
function probeSchemeView() {
  return probeRawRows_(CONFIG.backlog.sheetId, 'Scheme_View', 40);
}

function probeBacklogAllocation() {
  return probeRawRows_(CONFIG.backlog.sheetId, 'Backlog_allocation', 40);
}

/* ============================ MODEL BUILD ============================= */

function buildModel_(filters) {
  const ahtRows = readAht_();
  const backlogRows = readBacklog_();
  const rotaRows = readRota_();
  const overrides = readAllocations_();

  // ---- AHT per scheme (minutes/case), averaged across every entity row for
  // that scheme (AHT Validation has one row per reconciliation key, i.e. per
  // entity+scheme). Also build the Reconciliation -> Scheme lookup used to
  // translate Rota's entity-level rows into the plain scheme names that
  // Scheme_View (backlog) and this AHT aggregation both use. --------------
  const ahtSums = {};             // scheme -> { sum, n }
  const ahtByAgentScheme = {};
  const schemeByReconciliation = {};
  ahtRows.forEach(function (r) {
    if (r.reconciliation && r.scheme) {
      schemeByReconciliation[normalize_(r.reconciliation)] = r.scheme;
    }
    if (!r.scheme || !(r.aht > 0)) return;
    const bucket = ahtSums[r.scheme] || (ahtSums[r.scheme] = { sum: 0, n: 0 });
    bucket.sum += r.aht;
    bucket.n += 1;
    if (r.reporter) ahtByAgentScheme[r.reporter + '||' + r.scheme] = r.aht;
  });
  const ahtByScheme = {};
  Object.keys(ahtSums).forEach(function (s) { ahtByScheme[s] = ahtSums[s].sum / ahtSums[s].n; });
  const avgAht = avg_(Object.keys(ahtByScheme).map(function (k) { return ahtByScheme[k]; })) || 0;

  // Rota rows carry a raw reconciliation key (e.g. "US-VISA") in .scheme;
  // translate to the plain scheme name Scheme_View/AHT use for matching.
  // Falls back to the raw key if it isn't found in AHT Validation (that
  // assignment just won't match any known demand scheme downstream).
  const schemeFor_ = function (reconciliationKeyOrScheme) {
    return schemeByReconciliation[normalize_(reconciliationKeyOrScheme)] || reconciliationKeyOrScheme;
  };

  // ---- Filters ----------------------------------------------------------
  const schemeFilter = (filters.schemes && filters.schemes.length)
    ? new Set(filters.schemes) : null;
  const reporterFilter = filters.reporter || '';
  // Compared as "yyyy-MM-dd" strings, not Date objects — parseUkDate_ builds
  // Rota's dates in the script's local timezone, while a naive
  // `new Date("2026-07-20")` filter boundary parses as UTC midnight. Those
  // two can land on different sides of midnight depending on the deployed
  // script's timezone/DST, silently dropping rows (e.g. the whole team
  // showing 0 capacity). Formatting both sides through fmtDate_ before
  // comparing avoids that entirely.
  const from = filters.dateFrom || '';
  const to = filters.dateTo || '';

  const inDate = function (d) {
    if (!d) return !from && !to;         // rows with no date only survive when no date filter
    const key = fmtDate_(d);
    if (from && key < from) return false;
    if (to && key > to) return false;
    return true;
  };

  // ---- Backlog (demand) per scheme -------------------------------------
  const backlogByScheme = {};
  backlogRows.forEach(function (r) {
    if (!r.scheme) return;
    if (schemeFilter && !schemeFilter.has(r.scheme)) return;
    if (r.date && !inDate(r.date)) return;
    backlogByScheme[r.scheme] = (backlogByScheme[r.scheme] || 0) + (Number(r.volume) || 0);
  });

  // ---- Rota -> working slots per reporter ------------------------------
  const dayMinutes = CONFIG.productivityHoursPerDay * CONFIG.minutesPerHour;
  const perReporter = {};        // reporter -> {days, minutes, dates:Set}
  const teamDayReporters = {};   // dayKey -> { reporterName: true } — who's rostered each calendar day
  let teamWorkingSlots = 0;

  rotaRows.forEach(function (r) {
    if (!r.reporter) return;
    if (reporterFilter && r.reporter !== reporterFilter) return;
    if (!inDate(r.date)) return;
    if (!r.working) return;

    const rep = perReporter[r.reporter] || (perReporter[r.reporter] = {
      reporter: r.reporter, days: 0, minutes: 0, dates: {}, rotaScheme: {}
    });
    const dayKey = r.date ? fmtDate_(r.date) : ('row' + r.__row);
    if (!rep.dates[dayKey]) {
      rep.dates[dayKey] = true;
      rep.days += 1;
      const mins = (r.hours != null && r.hours > 0) ? r.hours * CONFIG.minutesPerHour : dayMinutes;
      rep.minutes += mins;
      teamWorkingSlots += 1;
    }
    const plainScheme = r.scheme ? schemeFor_(r.scheme) : '';
    if (plainScheme) rep.rotaScheme[plainScheme] = (rep.rotaScheme[plainScheme] || 0) + 1;

    (teamDayReporters[dayKey] || (teamDayReporters[dayKey] = {}))[r.reporter] = true;
  });

  // "Available headcount for the day" — averaged across every calendar day
  // in the filtered window (a single-day filter makes this exactly that
  // day's headcount; a multi-day range gives the typical daily staffing
  // level, which is what actually limits day-to-day throughput).
  const dayKeysWorked = Object.keys(teamDayReporters);
  const avgDailyHeadcount = dayKeysWorked.length
    ? round1_(dayKeysWorked.reduce(function (s, k) { return s + Object.keys(teamDayReporters[k]).length; }, 0) / dayKeysWorked.length)
    : 0;

  const reporters = Object.keys(perReporter).map(function (k) {
    const rep = perReporter[k];
    return {
      reporter: rep.reporter,
      days: rep.days,
      capacityHours: round1_(rep.minutes / CONFIG.minutesPerHour),
      capacityMinutes: rep.minutes,
      currentScheme: topKey_(rep.rotaScheme),
      // every plain scheme this reporter is actually rostered for — used to
      // restrict recommendations to schemes they're really assigned to.
      schemes: Object.keys(rep.rotaScheme)
    };
  }).sort(function (a, b) { return b.capacityMinutes - a.capacityMinutes; });

  const teamCapacityMinutes = reporters.reduce(function (s, r) { return s + r.capacityMinutes; }, 0);

  // ---- Demand per scheme (minutes) & scheme table ----------------------
  const schemeNames = uniqueSorted_(
    Object.keys(backlogByScheme).concat(Object.keys(ahtByScheme))
  ).filter(function (s) { return !schemeFilter || schemeFilter.has(s); });

  const schemes = schemeNames.map(function (name) {
    const cases = backlogByScheme[name] || 0;
    const aht = ahtByScheme[name] || avgAht;                 // fall back to team avg AHT
    const demandMinutes = cases * aht;
    return {
      scheme: name,
      backlogCases: cases,
      ahtMinutes: round1_(aht),
      ahtEstimated: !(ahtByScheme[name] > 0),
      demandHours: round1_(demandMinutes / CONFIG.minutesPerHour),
      demandMinutes: demandMinutes
    };
  }).sort(function (a, b) { return b.demandMinutes - a.demandMinutes; });

  const totalDemandMinutes = schemes.reduce(function (s, x) { return s + x.demandMinutes; }, 0);
  const totalBacklogCases = schemes.reduce(function (s, x) { return s + x.backlogCases; }, 0);

  // ---- Recommended allocation ------------------------------------------
  // See recommendAllocation_ — each reporter's own capacity is split only
  // across the schemes they're actually rostered for (proportional), or
  // schemes are filled from their rostered reporters biggest-backlog-first
  // (greedy). Never assigns work outside a reporter's real scheme coverage.
  const recommendations = recommendAllocation_(schemes, reporters, teamCapacityMinutes, overrides, filters);

  // ---- KPIs -------------------------------------------------------------
  // Utilisation = how much of the team's 7hr/day capacity the AHT-based
  // demand (backlog cases x new AHT) consumes.
  const utilisation = teamCapacityMinutes > 0
    ? round1_((totalDemandMinutes / teamCapacityMinutes) * 100) : 0;
  const surplusMinutes = teamCapacityMinutes - totalDemandMinutes;
  const daysToClear = (avgDailyHeadcount > 0)
    ? round1_(totalDemandMinutes / dayMinutes / avgDailyHeadcount) : 0;

  return {
    assumptions: {
      productivityHoursPerDay: CONFIG.productivityHoursPerDay,
      dayMinutes: dayMinutes
    },
    kpis: {
      teamCapacityHours: round1_(teamCapacityMinutes / CONFIG.minutesPerHour),
      teamSize: reporters.length, // distinct people rostered anywhere in the filtered window
      demandHours: round1_(totalDemandMinutes / CONFIG.minutesPerHour),
      utilisation: utilisation,
      backlogCases: totalBacklogCases,
      headcount: avgDailyHeadcount, // available headcount for the day (avg/day across the window)
      workingDays: teamWorkingSlots,
      surplusHours: round1_(surplusMinutes / CONFIG.minutesPerHour),
      daysToClear: daysToClear
    },
    schemes: schemes,
    reporters: reporters,
    recommendations: recommendations,
    overridesCount: overrides.length
  };
}

/**
 * Recommended allocation. Two strategies (chosen via filters.strategy):
 *   'proportional' (default) — each reporter's own capacity is split across
 *                              only the schemes they're actually rostered
 *                              for (rep.schemes), weighted by demand.
 *   'greedy'                 — clear the biggest backlog first, filling each
 *                              scheme from the reporters rostered for it
 *                              (biggest remaining capacity first) before
 *                              moving to the next scheme.
 * Both respect Rota's real scheme rostering — a reporter is never assigned
 * work for a scheme they aren't actually covering.
 */
function recommendAllocation_(schemes, reporters, teamCapacityMinutes, overrides, filters) {
  const strategy = (filters && filters.strategy) || 'proportional';
  if (!reporters.length) return [];
  return strategy === 'greedy'
    ? greedyAllocation_(schemes, reporters, overrides)
    : proportionalAllocation_(schemes, reporters, overrides);
}

function proportionalAllocation_(schemes, reporters, overrides) {
  const recs = [];
  if (!reporters.length) return recs;

  const schemeByName = {};
  schemes.forEach(function (sc) { schemeByName[sc.scheme] = sc; });

  const ovMap = {};
  overrides.forEach(function (o) { ovMap[o.reporter + '||' + o.scheme] = o.cases; });

  reporters.forEach(function (rep) {
    // only split this reporter's own capacity across schemes they actually
    // cover, in proportion to how backlogged each of those schemes is.
    const eligible = (rep.schemes || []).map(function (s) { return schemeByName[s]; }).filter(Boolean);
    if (!eligible.length) return;

    const portfolioDemand = eligible.reduce(function (s, sc) { return s + sc.demandMinutes; }, 0);

    eligible.forEach(function (sc) {
      const share = portfolioDemand > 0 ? (sc.demandMinutes / portfolioDemand) : (1 / eligible.length);
      const minutes = rep.capacityMinutes * share;
      const cases = sc.ahtMinutes > 0 ? minutes / sc.ahtMinutes : 0;
      const overrideKey = rep.reporter + '||' + sc.scheme;
      const overridden = Object.prototype.hasOwnProperty.call(ovMap, overrideKey);
      recs.push({
        reporter: rep.reporter,
        scheme: sc.scheme,
        recommendedCases: Math.round(cases),
        recommendedHours: round1_(minutes / CONFIG.minutesPerHour),
        allocatedCases: overridden ? ovMap[overrideKey] : Math.round(cases),
        overridden: overridden,
        capacityHours: rep.capacityHours
      });
    });
  });

  appendUnplacedOverrides_(recs, ovMap, reporters);
  recs.sort(function (a, b) { return b.recommendedHours - a.recommendedHours; });
  return recs;
}

/**
 * Greedy allocation — walk schemes from largest backlog to smallest and fill
 * each one from the reporters actually rostered for it (biggest remaining
 * capacity first) before moving to the next scheme. Produces focused
 * assignments (fewer schemes per person) rather than spreading everyone
 * thinly across every scheme they cover.
 */
function greedyAllocation_(schemes, reporters, overrides) {
  const recs = [];
  const ovMap = {};
  overrides.forEach(function (o) { ovMap[o.reporter + '||' + o.scheme] = o.cases; });

  const remaining = {};
  reporters.forEach(function (r) { remaining[r.reporter] = r.capacityMinutes; });

  const emitted = {}; // reporter||scheme -> true

  schemes.forEach(function (sc) {
    let need = sc.demandMinutes;
    if (need <= 0) return;
    // only reporters actually rostered for this scheme, most remaining capacity first
    const pool = reporters
      .filter(function (r) { return (r.schemes || []).indexOf(sc.scheme) !== -1; })
      .sort(function (a, b) { return remaining[b.reporter] - remaining[a.reporter]; });

    for (let i = 0; i < pool.length && need > 0.01; i++) {
      const rep = pool[i];
      const avail = remaining[rep.reporter];
      if (avail <= 0) continue;
      const use = Math.min(avail, need);
      remaining[rep.reporter] -= use;
      need -= use;
      const cases = sc.ahtMinutes > 0 ? use / sc.ahtMinutes : 0;
      const key = rep.reporter + '||' + sc.scheme;
      const overridden = Object.prototype.hasOwnProperty.call(ovMap, key);
      recs.push({
        reporter: rep.reporter,
        scheme: sc.scheme,
        recommendedCases: Math.round(cases),
        recommendedHours: round1_(use / CONFIG.minutesPerHour),
        allocatedCases: overridden ? ovMap[key] : Math.round(cases),
        overridden: overridden,
        capacityHours: rep.capacityHours
      });
      emitted[key] = true;
    }
  });

  appendUnplacedOverrides_(recs, ovMap, reporters, emitted);
  recs.sort(function (a, b) { return b.recommendedHours - a.recommendedHours; });
  return recs;
}

/** Surfaces manual overrides for (reporter, scheme) pairs the strategy didn't otherwise place. */
function appendUnplacedOverrides_(recs, ovMap, reporters, emitted) {
  const seen = emitted || {};
  if (!emitted) recs.forEach(function (r) { seen[r.reporter + '||' + r.scheme] = true; });
  Object.keys(ovMap).forEach(function (key) {
    if (seen[key]) return;
    const parts = key.split('||');
    const rep = reporters.filter(function (r) { return r.reporter === parts[0]; })[0];
    recs.push({
      reporter: parts[0], scheme: parts[1],
      recommendedCases: 0, recommendedHours: 0,
      allocatedCases: ovMap[key], overridden: true,
      capacityHours: rep ? rep.capacityHours : 0
    });
  });
}

/* ============================ READERS ================================= */

function readAht_() {
  const cfg = CONFIG.aht;
  const t = readTable_(cfg.sheetId, cfg.tab, cfg.headerRow);
  const cols = detectColumns_(t.headers, cfg.columns);
  const ahtIdx = columnLetterToIndex_(cfg.ahtColumnLetter); // 0-based
  const factor = cfg.ahtUnit === 'seconds' ? (1 / 60)
    : cfg.ahtUnit === 'hours' ? 60 : 1;                     // -> minutes

  return t.rows.map(function (row) {
    const raw = row.__cells[ahtIdx];
    return {
      reconciliation: cleanStr_(cols.reconciliation != null ? row.__cells[cols.reconciliation] : ''),
      scheme: cleanStr_(cols.scheme != null ? row.__cells[cols.scheme] : ''),
      reporter: cleanStr_(cols.reporter != null ? row.__cells[cols.reporter] : ''),
      aht: parseNum_(raw) * factor
    };
  }).filter(function (r) { return r.scheme || r.aht > 0; });
}

function readBacklog_() {
  const cfg = CONFIG.backlog;
  const t = readTable_(cfg.sheetId, cfg.tab, cfg.headerRow);
  const cols = detectColumns_(t.headers, cfg.columns);
  return t.rows.map(function (row) {
    return {
      scheme: cleanStr_(cols.scheme != null ? row.__cells[cols.scheme] : ''),
      volume: cols.volume != null ? parseNum_(row.__cells[cols.volume]) : 0,
      date: cols.date != null ? toDate_(row.__cells[cols.date]) : null
    };
  }).filter(function (r) {
    if (!r.scheme) return false;
    // Scheme_View is a pivot table: skip its own "Grand Total" row and any
    // "**note" rows (e.g. "**refine by entity") rather than real schemes.
    if (r.scheme.toLowerCase().indexOf('grand total') !== -1) return false;
    if (r.scheme.indexOf('**') === 0) return false;
    return true;
  });
}

/**
 * Rota reader — bespoke to the real "Rota" tab layout (see CONFIG.rota.matrix
 * and the module doc comment): each row is one reconciliation key (entity +
 * scheme), and the reviewer covering it on a given weekday sits inside that
 * day's cell. Emits one entry per (reconciliation key, weekday) where a
 * reviewer is actually assigned — .scheme holds the RAW reconciliation key
 * (e.g. "US-VISA"); callers translate it to a plain scheme name via the
 * AHT-derived Reconciliation -> Scheme lookup before matching against
 * backlog/demand.
 */
function readRota_() {
  const cfg = CONFIG.rota;
  const m = cfg.matrix;
  const ss = SpreadsheetApp.openById(cfg.sheetId);
  const sheet = ss.getSheetByName(cfg.tab);
  if (!sheet) throw new Error('Tab "' + cfg.tab + '" not found in ' + cfg.sheetId);
  const values = sheet.getDataRange().getValues();
  if (values.length < m.dataStartRow) return [];

  const dateRow = values[m.dateHeaderRow - 1];
  const dayDates = m.dayCols.map(function (col) { return parseUkDate_(dateRow[col]); });

  const out = [];
  for (let i = m.dataStartRow - 1; i < values.length; i++) {
    const row = values[i];
    const reconciliationKey = cleanStr_(row[m.reconciliationKeyCol]);
    if (!reconciliationKey) continue;
    const frequency = cleanStr_(row[m.frequencyCol]);
    m.dayCols.forEach(function (col, idx) {
      const reviewer = cleanStr_(row[col]);
      if (!reviewer) return;
      out.push({
        __row: i + 1,
        reporter: reviewer,
        date: dayDates[idx],
        status: frequency,
        working: true,
        scheme: reconciliationKey,
        hours: null
      });
    });
  }
  return out;
}

function readAllocations_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.rota.sheetId);
    const sheet = ss.getSheetByName(CONFIG.rota.allocationsTab);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return values.map(function (r) {
      return { date: fmtDate_(toDate_(r[0])), reporter: cleanStr_(r[1]), scheme: cleanStr_(r[2]), cases: parseNum_(r[3]) };
    }).filter(function (r) { return r.reporter && r.scheme; });
  } catch (e) {
    return [];
  }
}

/* ============================ SHEET UTILITIES ========================= */

function readTable_(sheetId, tabName, headerRow) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Tab "' + tabName + '" not found in ' + sheetId);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < headerRow) return { headers: [], rows: [] };

  const headers = values[headerRow - 1].map(function (h) { return cleanStr_(h); });
  const rows = [];
  for (let i = headerRow; i < values.length; i++) {
    const cells = values[i];
    if (cells.every(function (c) { return c === '' || c == null; })) continue;
    rows.push({ __row: i + 1, __cells: cells });
  }
  return { headers: headers, rows: rows };
}

function detectColumns_(headers, candidatesByField) {
  const norm = headers.map(function (h) { return normalize_(h); });
  const result = {};
  Object.keys(candidatesByField).forEach(function (field) {
    const candidates = candidatesByField[field];
    let found = null;
    // 1) exact normalised match
    for (let c = 0; c < candidates.length && found == null; c++) {
      const cand = normalize_(candidates[c]);
      for (let i = 0; i < norm.length; i++) {
        if (norm[i] === cand) { found = i; break; }
      }
    }
    // 2) contains match
    if (found == null) {
      for (let c = 0; c < candidates.length && found == null; c++) {
        const cand = normalize_(candidates[c]);
        for (let i = 0; i < norm.length; i++) {
          if (norm[i] && (norm[i].indexOf(cand) !== -1 || cand.indexOf(norm[i]) !== -1)) { found = i; break; }
        }
      }
    }
    if (found != null) result[field] = found;
  });
  return result;
}

function columnLetterToIndex_(letter) {
  let idx = 0;
  const s = String(letter).toUpperCase();
  for (let i = 0; i < s.length; i++) {
    idx = idx * 26 + (s.charCodeAt(i) - 64);
  }
  return idx - 1; // 0-based
}

/* ============================ SMALL HELPERS =========================== */

function normalize_(v) {
  return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function cleanStr_(v) { return String(v == null ? '' : v).trim(); }

function parseNum_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v == null || v === '') return null;
  if (typeof v === 'number') { // spreadsheet serial fallback
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Parses UK-style "DD/MM/YYYY" text dates, which is how the Rota tab's first
 * week-block header row stores its dates (as plain text, not real Date
 * cells) — the plain `new Date(v)` constructor used by toDate_ would
 * misread "20/07/2026" as MM/DD/YYYY (or fail outright). Falls back to
 * toDate_ for anything that isn't that exact text shape.
 */
function parseUkDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      return isNaN(d.getTime()) ? null : d;
    }
  }
  return toDate_(v);
}

function fmtDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Europe/London', 'yyyy-MM-dd');
}
function fmtDateTime_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Europe/London', 'yyyy-MM-dd HH:mm');
}

function uniqueSorted_(arr) {
  const seen = {};
  const out = [];
  arr.forEach(function (v) {
    const s = cleanStr_(v);
    if (s && !seen[s.toLowerCase()]) { seen[s.toLowerCase()] = true; out.push(s); }
  });
  return out.sort(function (a, b) { return a.localeCompare(b); });
}

function avg_(arr) {
  const nums = (arr || []).filter(function (n) { return typeof n === 'number' && !isNaN(n); });
  if (!nums.length) return 0;
  return nums.reduce(function (s, n) { return s + n; }, 0) / nums.length;
}

function topKey_(obj) {
  let best = '', bestN = -1;
  Object.keys(obj || {}).forEach(function (k) { if (obj[k] > bestN) { bestN = obj[k]; best = k; } });
  return best;
}

function round1_(n) { return Math.round((Number(n) || 0) * 10) / 10; }
