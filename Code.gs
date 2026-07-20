/**
 * AHT Capacity Planning Dashboard  —  server side (Google Apps Script)
 * ---------------------------------------------------------------------
 * Pulls live data from three Google Sheets and computes a capacity plan
 * per employee (reporter) and for the team:
 *
 *   1. AHT analysis   -> "AHT Validation" tab, column E ("new AHT", minutes/case)
 *   2. Rota           -> "Rota" tab (who is working, which day, which scheme)
 *   3. Backlog        -> "Backlog_Team Planning" tab (open cases per scheme)
 *
 * Core assumption: 7 productive hours per employee per working day.
 *
 * The three source sheets are owned/maintained by other people and their
 * exact column layouts can change, so this script DOES NOT hard-code column
 * positions (except the explicitly requested AHT column E). Instead it
 * detects columns by header name using the candidate lists in CONFIG. If a
 * layout differs, run `diagnostics()` from the editor (View > Logs) to see
 * the detected headers and adjust the candidate lists below.
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
      scheme:   ['scheme', 'schemes', 'claim type', 'type', 'reason', 'reason code', 'category', 'queue', 'product'],
      reporter: ['reporter', 'employee', 'agent', 'analyst', 'name', 'owner'] // optional per-agent AHT
    }
  },

  rota: {
    sheetId: '1-K8fLgvU7h52ZYdEwEfZR5q3gYSyURjsCVJRlwB_K-g',
    tab: 'Rota',
    headerRow: 1,
    allocationsTab: 'Dashboard Allocations', // written back into the rota file
    columns: {
      reporter: ['reporter', 'employee', 'agent', 'name', 'analyst', 'team member', 'person'],
      date:     ['date', 'day', 'shift date', 'rota date', 'work date'],
      status:   ['status', 'shift', 'activity', 'availability', 'attendance', 'state'],
      scheme:   ['scheme', 'allocation', 'assigned scheme', 'queue', 'assigned'],
      hours:    ['hours', 'productive hours', 'capacity hours', 'available hours']
    },
    // values (case-insensitive, matched as substrings) that mean "not working"
    nonWorkingStatuses: [
      'off', 'day off', 'rest', 'annual leave', 'al', 'holiday', 'hol', 'leave',
      'sick', 'absent', 'absence', 'bank holiday', 'bh', 'lieu', 'toil', 'n/a', 'na'
    ],
    // if a status is neither non-working nor blank it is treated as working.
    // Blank cells in a matrix layout are treated as NOT working.
    trainingStatuses: ['training', 'train', 'meeting', 'admin', '1:1', 'coaching']
  },

  backlog: {
    sheetId: '1-ywS9-yFh0rJYwX421uADNZbr3UjgAWVcxkB1kfsQJE',
    tab: 'Backlog_Team Planning',
    headerRow: 1,
    columns: {
      scheme: ['scheme', 'queue', 'claim type', 'type', 'category', 'product', 'reason'],
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

    const schemes = uniqueSorted_(
      aht.map(function (r) { return r.scheme; })
        .concat(backlog.map(function (r) { return r.scheme; }))
        .concat(rota.map(function (r) { return r.scheme; }))
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
 * Editor helper: run from the Apps Script editor and open View > Logs to see
 * the headers each source sheet actually exposes, plus the columns detected.
 */
function diagnostics() {
  const out = [];
  [['AHT', CONFIG.aht], ['ROTA', CONFIG.rota], ['BACKLOG', CONFIG.backlog]].forEach(function (pair) {
    const label = pair[0], cfg = pair[1];
    try {
      const t = readTable_(cfg.sheetId, cfg.tab, cfg.headerRow);
      out.push(label + ' [' + cfg.tab + ']  headers: ' + JSON.stringify(t.headers));
      out.push('   detected: ' + JSON.stringify(detectColumns_(t.headers, cfg.columns)));
      out.push('   rows: ' + t.rows.length + '  sample: ' + JSON.stringify(t.rows[0] || {}));
    } catch (e) {
      out.push(label + '  ERROR: ' + e);
    }
  });
  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/* ============================ MODEL BUILD ============================= */

function buildModel_(filters) {
  const ahtRows = readAht_();
  const backlogRows = readBacklog_();
  const rotaRows = readRota_();
  const overrides = readAllocations_();

  // ---- AHT per scheme (minutes/case). Prefer scheme-level; keep per-agent. -
  const ahtByScheme = {};
  const ahtByAgentScheme = {};
  ahtRows.forEach(function (r) {
    if (!r.scheme || !(r.aht > 0)) return;
    // last non-empty wins for scheme-level; average would also be valid.
    ahtByScheme[r.scheme] = r.aht;
    if (r.reporter) ahtByAgentScheme[r.reporter + '||' + r.scheme] = r.aht;
  });
  const avgAht = avg_(Object.keys(ahtByScheme).map(function (k) { return ahtByScheme[k]; })) || 0;

  // ---- Filters ----------------------------------------------------------
  const schemeFilter = (filters.schemes && filters.schemes.length)
    ? new Set(filters.schemes) : null;
  const reporterFilter = filters.reporter || '';
  const from = filters.dateFrom ? toDate_(filters.dateFrom) : null;
  const to = filters.dateTo ? toDate_(filters.dateTo) : null;

  const inDate = function (d) {
    if (!d) return !from && !to;         // rows with no date only survive when no date filter
    if (from && d < from) return false;
    if (to && d > to) return false;
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
    if (r.scheme) rep.rotaScheme[r.scheme] = (rep.rotaScheme[r.scheme] || 0) + 1;
  });

  const reporters = Object.keys(perReporter).map(function (k) {
    const rep = perReporter[k];
    return {
      reporter: rep.reporter,
      days: rep.days,
      capacityHours: round1_(rep.minutes / CONFIG.minutesPerHour),
      capacityMinutes: rep.minutes,
      currentScheme: topKey_(rep.rotaScheme)
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
  // Distribute available team capacity across schemes in proportion to their
  // demand (largest backlog first), then split each scheme's assigned minutes
  // across rostered reporters proportionally to each reporter's capacity.
  const recommendations = recommendAllocation_(schemes, reporters, teamCapacityMinutes, overrides, filters);

  // ---- KPIs -------------------------------------------------------------
  const utilisation = teamCapacityMinutes > 0
    ? round1_((totalDemandMinutes / teamCapacityMinutes) * 100) : 0;
  const surplusMinutes = teamCapacityMinutes - totalDemandMinutes;
  const daysToClear = (teamCapacityMinutes > 0)
    ? round1_(totalDemandMinutes / dayMinutes / Math.max(reporters.length, 1)) : 0;

  return {
    assumptions: {
      productivityHoursPerDay: CONFIG.productivityHoursPerDay,
      dayMinutes: dayMinutes
    },
    kpis: {
      teamCapacityHours: round1_(teamCapacityMinutes / CONFIG.minutesPerHour),
      demandHours: round1_(totalDemandMinutes / CONFIG.minutesPerHour),
      utilisation: utilisation,
      backlogCases: totalBacklogCases,
      headcount: reporters.length,
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
 *   'proportional' (default) — demand-weighted split across all employees.
 *   'greedy'                 — clear the biggest backlog first, filling one
 *                              scheme from the capacity pool before the next.
 */
function recommendAllocation_(schemes, reporters, teamCapacityMinutes, overrides, filters) {
  const strategy = (filters && filters.strategy) || 'proportional';
  if (!reporters.length) return [];
  return strategy === 'greedy'
    ? greedyAllocation_(schemes, reporters, overrides)
    : proportionalAllocation_(schemes, reporters, teamCapacityMinutes, overrides);
}

function proportionalAllocation_(schemes, reporters, teamCapacityMinutes, overrides) {
  const recs = [];
  if (!reporters.length) return recs;

  const totalDemand = schemes.reduce(function (s, x) { return s + x.demandMinutes; }, 0);
  const totalRepMinutes = reporters.reduce(function (s, r) { return s + r.capacityMinutes; }, 0) || 1;

  // How much capacity (minutes) each scheme should get.
  const schemeShare = {};
  schemes.forEach(function (sc) {
    schemeShare[sc.scheme] = totalDemand > 0
      ? Math.min(sc.demandMinutes, teamCapacityMinutes * (sc.demandMinutes / totalDemand))
      : 0;
  });

  // Build override lookup: reporter||scheme -> cases (per current filter window)
  const ovMap = {};
  overrides.forEach(function (o) {
    ovMap[o.reporter + '||' + o.scheme] = o.cases;
  });

  reporters.forEach(function (rep) {
    const repFraction = rep.capacityMinutes / totalRepMinutes;
    // give this reporter their share of each scheme's assigned capacity
    schemes.forEach(function (sc) {
      const minutes = schemeShare[sc.scheme] * repFraction;
      if (minutes <= 0 && !ovMap[rep.reporter + '||' + sc.scheme]) return;
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

  // Keep only the meaningful lines (reporter's top schemes) to avoid noise:
  // sort by recommended hours desc.
  recs.sort(function (a, b) { return b.recommendedHours - a.recommendedHours; });
  return recs;
}

/**
 * Greedy allocation — walk schemes from largest backlog to smallest and fill
 * each one from the team's remaining capacity (biggest-capacity employees
 * first) before moving on. Produces focused assignments (fewer schemes per
 * person) rather than spreading everyone thinly across every scheme.
 */
function greedyAllocation_(schemes, reporters, overrides) {
  const recs = [];
  const ovMap = {};
  overrides.forEach(function (o) { ovMap[o.reporter + '||' + o.scheme] = o.cases; });

  // remaining capacity (minutes) per reporter; biggest capacity used first
  const pool = reporters.slice().sort(function (a, b) { return b.capacityMinutes - a.capacityMinutes; });
  const remaining = {};
  pool.forEach(function (r) { remaining[r.reporter] = r.capacityMinutes; });

  const emitted = {}; // reporter||scheme -> true

  schemes.forEach(function (sc) {
    let need = sc.demandMinutes;
    if (need <= 0) return;
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

  // surface any manual overrides that greedy didn't already place
  Object.keys(ovMap).forEach(function (key) {
    if (emitted[key]) return;
    const parts = key.split('||');
    const rep = reporters.filter(function (r) { return r.reporter === parts[0]; })[0];
    recs.push({
      reporter: parts[0], scheme: parts[1],
      recommendedCases: 0, recommendedHours: 0,
      allocatedCases: ovMap[key], overridden: true,
      capacityHours: rep ? rep.capacityHours : 0
    });
  });

  recs.sort(function (a, b) { return b.recommendedHours - a.recommendedHours; });
  return recs;
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
  }).filter(function (r) { return r.scheme; });
}

/**
 * Rota reader. Supports two layouts, auto-detected:
 *   (a) LONG   : one row per (reporter, date) with a status column.
 *   (b) MATRIX : reporter in one column, date headers across the top, the
 *                cell value being the shift/status.
 * Emits a normalised list of { reporter, date, status, working, scheme, hours }.
 */
function readRota_() {
  const cfg = CONFIG.rota;
  const t = readTable_(cfg.sheetId, cfg.tab, cfg.headerRow);
  const cols = detectColumns_(t.headers, cfg.columns);

  // Date columns = header cells that parse as dates (MATRIX layout signal).
  const dateCols = [];
  t.headers.forEach(function (h, i) {
    const d = toDate_(h);
    if (d && !isNaN(d.getTime())) dateCols.push({ index: i, date: d });
  });

  const out = [];
  const isLong = cols.date != null && dateCols.length < 2;

  if (isLong) {
    t.rows.forEach(function (row) {
      const reporter = cleanStr_(cols.reporter != null ? row.__cells[cols.reporter] : '');
      if (!reporter) return;
      const status = cleanStr_(cols.status != null ? row.__cells[cols.status] : '');
      const date = cols.date != null ? toDate_(row.__cells[cols.date]) : null;
      out.push({
        __row: row.__row,
        reporter: reporter,
        date: date,
        status: status,
        working: isWorking_(status, true),
        scheme: cleanStr_(cols.scheme != null ? row.__cells[cols.scheme] : ''),
        hours: cols.hours != null ? parseNum_(row.__cells[cols.hours]) : null
      });
    });
  } else {
    // MATRIX: reporter column + one column per date
    t.rows.forEach(function (row) {
      const reporter = cleanStr_(cols.reporter != null ? row.__cells[cols.reporter] : '');
      if (!reporter) return;
      dateCols.forEach(function (dc) {
        const val = cleanStr_(row.__cells[dc.index]);
        out.push({
          __row: row.__row,
          reporter: reporter,
          date: dc.date,
          status: val,
          working: isWorking_(val, false), // blank in matrix = not working
          scheme: '',
          hours: null
        });
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

function isWorking_(status, blankIsWorkingInLong) {
  const s = normalize_(status);
  if (!s) return !!blankIsWorkingInLong ? false : false; // blank never counts as working
  const non = CONFIG.rota.nonWorkingStatuses;
  for (let i = 0; i < non.length; i++) {
    if (s.indexOf(normalize_(non[i])) !== -1) return false;
  }
  return true; // any other non-blank status (In, WFH, Office, a scheme name, training, etc.) = working
}

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
