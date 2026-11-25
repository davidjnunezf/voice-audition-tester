/**
 * Portfolio Reviewer — Studio Tiny Tuko
 * ------------------------------------------------------------
 * Sidebar + voice remote for reviewing the artist application sheet.
 * Each rating fills in Reviewed? and the Rating dropdown, then opens the
 * next applicant's portfolio in the portfolio tab (or popup window).
 *
 *   next / "not feeling it" → Rating "Not feeling it", Reviewed? = Yes, open next portfolio
 *   good / "not bad"        → "Not bad!"
 *   peak                    → "PEAK"
 *   one / two / three       → the same three, in that order (keys 1–3 too)
 *   N/A                     → Reviewed? = No, Rating "Not feeling it" (no portfolio / broken link)
 *   skip                    → next portfolio, nothing written
 *   link                    → their next link, when they gave more than one
 *   open                    → reopen this portfolio
 *   add to layout           → also copy them into the Layout tab (any category tab) when rated
 *   remove from layout      → take that back before rating
 * A rating is also written into every other tab the applicant is already in (Notes there are left alone).
 *   go to 722 / go to Alyssa → jump to that applicant (say it again for the next Alyssa)
 *   progress                → says how many are done today and how long the rest will take
 *
 * Two people can review at once: each sidebar holds the applicant it's showing (and the next one)
 * for 10 minutes, and the other skips past them. The next applicant shows immediately; the
 * rating is saved in the background. "Build shortlist" makes a Shortlist tab of PEAK / Not bad per tab.
 *   notes …                 → appends "- …" to Notes
 *   back / undo / top / bottom / descending / ascending
 *
 * The Rating options are read from the column's dropdown, so the exact
 * option text (spaces and punctuation included) is written.
 * All functions are prefixed with pr_ so they won't clash with other scripts.
 */

const PR_CONFIG = {
  // Header names to look for (case-insensitive). Exact matches win, then "starts with".
  headers: {
    number:    ['Applicant #', 'Applicant'],
    name:      ['Full Name', 'Name'],
    reviewed:  ['Reviewed?', 'Reviewed'],
    rating:    ['Rating'],
    notes:     ['Notes'],
    portfolio: ['Portfolio / Website', 'Portfolio'],
    roles:     ['Which role(s)', 'Role'],
    instagram: ['Instagram Handle', 'Instagram'],
    about:     ['Tell us a little about yourself', 'Tell us about yourself'],
    country:   ['Country & Timezone', 'Country'],
  },

  // What each command picks in the Rating dropdown. Matching ignores case,
  // spaces and punctuation. The order here is the 1 / 2 / 3 order.
  ratings: [
    { cmd: 'next', option: 'Not feeling it', key: 'N' },
    { cmd: 'good', option: 'Not bad',        key: 'G' },
    { cmd: 'peak', option: 'PEAK',           key: 'P' },
  ],
  aliases: { excellent: 'peak' },

  // No portfolio / broken link.
  na: { reviewed: 'No', option: 'Not feeling it' },

  reviewedValue: 'Yes',

  // If the open tab has nothing left to review, use this tab instead.
  mainSheet: 'Form Responses 1',

  // How long a sidebar holds the applicant it's showing, so a second reviewer skips it.
  holdMinutes: 10,

  // The tab "Build shortlist" writes.
  shortlistSheet: 'Shortlist',

  // Where voice.html and portfolio-note.html are hosted (GitHub Pages).
  voiceUrl: 'https://davidjnunezf.github.io/voice-audition-tester/voice.html',
  noteUrl:  'https://davidjnunezf.github.io/voice-audition-tester/portfolio-note.html',
};

/* ---------------- Menu & sidebar ---------------- */

// If your project already has an onOpen(), delete this one and add
// pr_addMenu(); inside yours instead.
function onOpen() {
  pr_addMenu();
}

function pr_addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🎨 Portfolio Reviewer')
    .addItem('Open portfolio reviewer', 'pr_openSidebar')
    .addItem('Build shortlist', 'pr_buildShortlistMenu')
    .addToUi();
}

function pr_openSidebar() {
  const t = HtmlService.createTemplateFromFile('PortfolioSidebar');
  t.clientConfig = JSON.stringify({ voiceUrl: PR_CONFIG.voiceUrl, noteUrl: PR_CONFIG.noteUrl });
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Portfolio Reviewer'));
}

/* ---------------- Called from the sidebar ----------------
 * opts = { sid, showing }: sid identifies this sidebar (for holds); showing is the applicant the
 * sidebar already moved to on its own (instant next), which is kept if it's still free.
 * Every call returns { current, peek, remaining, others, stats }: peek is the applicant after
 * current, so the sidebar can show it the moment you rate.
 */

/** Opens on the selected row if it hasn't been reviewed yet, otherwise the next open one. */
function pr_init(opts) {
  opts = opts || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let cols = null;
  let scan = null;
  try {
    cols = pr_columns_(sheet);
    scan = pr_scan_(sheet, cols);
  } catch (e) {
    cols = null;
  }
  // A finished category tab (or a tab without the right columns): fall back to the main tab.
  if (!cols || !scan.remaining) {
    const main = ss.getSheetByName(PR_CONFIG.mainSheet);
    if (main && main.getSheetId() !== sheet.getSheetId()) {
      const mainCols = pr_columns_(main);
      const mainScan = pr_scan_(main, mainCols);
      if (!cols || mainScan.remaining) {
        sheet = main; cols = mainCols; scan = mainScan;
        sheet.activate();
      }
    }
  }
  if (!cols) cols = pr_columns_(sheet); // throws the helpful "couldn't find" message

  const dir = pr_getPref_('dir') === 'up' ? 'up' : 'down';
  const range = sheet.getActiveRange();
  const active = range && range.getSheet().getSheetId() === sheet.getSheetId() ? range.getRow() : 0;
  const at = active >= 2 && active <= scan.lastRow ? active : 0;
  const p = pr_place_(sheet, cols, opts.sid, dir, { showing: at, from: at, edge: true });
  const out = pr_land_(sheet, cols, p);
  out.sheetName = sheet.getName();
  out.title = ss.getName();
  out.tabs = pr_tabs_(ss, sheet);
  out.dir = dir;
  out.view = pr_getPref_('view') === 'popup' ? 'popup' : 'tab';
  out.choices = pr_choices_(pr_options_(sheet, cols, p.row || 2));
  return out;
}

/**
 * Rates the row, copies it into any tabs in addTo (e.g. ['Layout']), keeps its rating in sync in
 * the tabs it's already in, then moves on.
 */
function pr_rate(sheetName, row, cmd, dir, value, addTo, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  const spec = pr_spec_(cmd, value, pr_options_(sheet, cols, row));

  const reviewedCell = sheet.getRange(row, cols.reviewed);
  const ratingCell = sheet.getRange(row, cols.rating);
  const label = spec.reviewed === PR_CONFIG.na.reviewed ? 'N/A' : String(spec.rating).trim();
  const undo = {
    kind: 'rate', row: row, reviewed: reviewedCell.getValue(), rating: ratingCell.getValue(), copies: [],
    label: label, day: pr_today_(), wasOpen: String(reviewedCell.getDisplayValue()).trim() === '',
  };

  reviewedCell.setValue(spec.reviewed);
  ratingCell.setValue(spec.rating);
  const added = pr_copyAll_(sheet, cols, row, addTo);
  const synced = pr_syncTabs_(sheet, cols, row, added.map(function (a) { return a.tab; }));
  undo.copies = added.concat(synced);
  if (undo.wasOpen) pr_withLock_(function () { pr_bump_(sheet, label, 1); });

  const p = pr_place_(sheet, cols, opts.sid, dir, { from: row, showing: opts.showing });
  const out = pr_land_(sheet, cols, p);
  out.rated = { row: row, name: pr_nameAt_(sheet, cols, row), rating: String(spec.rating).trim(), reviewed: spec.reviewed, cmd: cmd, label: label };
  out.added = added.map(function (a) { return { tab: a.tab, updated: !!a.prev }; });
  out.synced = synced.map(function (a) { return a.tab; });
  out.undo = undo;
  return out;
}

/** Copies the row into the given tabs right away (used when the applicant is already rated, or on skip). */
function pr_addTo(sheetName, row, tabs) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  const added = pr_copyAll_(sheet, cols, row, tabs);
  return {
    added: added.map(function (a) { return { tab: a.tab, updated: !!a.prev }; }),
    undo: { kind: 'add', row: row, copies: added },
  };
}

/** Moves on without writing anything. */
function pr_skip(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  return pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { from: row, showing: opts.showing }));
}

/**
 * Reverts a rating or an "add to" (using the undo object pr_rate / pr_addTo returned):
 * copies it made are removed (or put back as they were), and a rating goes back to that row.
 */
function pr_undo(sheetName, undo, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, undo.row);
  const removed = pr_removeCopies_(undo.copies || []);
  if (undo.kind !== 'add') {
    sheet.getRange(undo.row, cols.reviewed).setValue(undo.reviewed === null ? '' : undo.reviewed);
    sheet.getRange(undo.row, cols.rating).setValue(undo.rating === null ? '' : undo.rating);
    if (undo.wasOpen && undo.label) pr_withLock_(function () { pr_bump_(sheet, undo.label, -1, undo.day); });
  }
  const out = pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { row: undo.row }));
  out.removed = removed;
  return out;
}

/** Appends "- text" to the Notes cell. */
function pr_addNote(sheetName, row, text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean) throw new Error('Empty note');
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  const cell = sheet.getRange(row, cols.notes);
  const current = String(cell.getDisplayValue() || '').replace(/\s+$/, '');
  const line = '- ' + clean;
  cell.setNumberFormat('@').setValue(current ? current + '\n' + line : line);
  return { row: row, notes: cell.getDisplayValue() };
}

/** Jump to the first open applicant from the top ('top') or bottom ('bottom'). */
function pr_jump(sheetName, where, opts) {
  opts = opts || {};
  const dir = where === 'bottom' ? 'up' : 'down';
  pr_setPref('dir', dir);
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const out = pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { edge: true }));
  out.dir = dir;
  return out;
}

/** Use whatever row is selected in the sheet right now (even if already reviewed). */
function pr_startHere(sheetName, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const range = sheet.getActiveRange();
  const row = range && range.getSheet().getSheetId() === sheet.getSheetId() ? range.getRow() : 0;
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Select a row with an applicant on "' + sheetName + '" first.');
  return pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { row: row }));
}

/** Opens a specific row without changing anything (Back, a new direction, or the keep-alive). */
function pr_rowAt(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  return pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { row: row }));
}

/** Keeps holding this row (and the next) without moving the sheet's selection. */
function pr_hold(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const p = pr_place_(sheet, cols, opts.sid, dir, { row: row });
  return {
    remaining: p.scan.remaining, others: p.others, stats: pr_stats_(sheet),
    current: { row: row }, peek: p.peek ? pr_rowInfo_(sheet, cols, p.peek) : null,
  };
}

/** "go to 722" / "go to Alyssa": the applicant number, a row ("row 40"), or part of a name. */
function pr_goto(sheetName, query, fromRow, dir, opts) {
  opts = opts || {};
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const hit = pr_find_(sheet, [cols.name], cols.number, query, fromRow, dir);
  const out = pr_land_(sheet, cols, pr_place_(sheet, cols, opts.sid, dir, { row: hit.row }));
  out.found = hit;
  return out;
}

/** Today's numbers for the progress panel (and the "progress" voice command). */
function pr_progress(sheetName, opts) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const claims = pr_claims_(sheet);
  return { remaining: pr_scan_(sheet, cols).remaining, stats: pr_stats_(sheet), others: pr_others_(claims, (opts || {}).sid) };
}

/** Saves a per-person setting for this spreadsheet ('dir' or 'view'). */
function pr_setPref(key, value) {
  if (['dir', 'view'].indexOf(key) === -1) throw new Error('Unknown setting ' + key);
  PropertiesService.getUserProperties().setProperty(pr_prefKey_(key), String(value));
  return value;
}

/* ---------------- Category tabs ("add to …") ---------------- */

/** Other tabs with the same applicant columns (Layout, Character Design, Story…), in sheet order. */
function pr_tabs_(ss, source) {
  const out = [];
  ss.getSheets().forEach(function (sh) {
    if (sh.getSheetId() === source.getSheetId() || sh.getName() === PR_CONFIG.mainSheet) return;
    if (sh.isSheetHidden && sh.isSheetHidden()) return;
    try {
      const c = pr_columns_(sh);
      if (c.name) out.push(sh.getName());
    } catch (e) {
      // Not an applicant tab.
    }
  });
  return out;
}

/**
 * Copies the row into each tab. Returns [{ tab, row, number, name, prev }].
 * If they're already in a tab, only Reviewed? and Rating are refreshed there, and prev holds the old ones.
 */
function pr_copyAll_(sheet, cols, row, tabs) {
  const list = (tabs || []).filter(function (t, i, a) { return t && a.indexOf(t) === i; });
  if (!list.length) return [];
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    return list.map(function (t) { return pr_copyTo_(sheet, cols, row, t); });
  } finally {
    lock.releaseLock();
  }
}

function pr_copyTo_(sheet, cols, row, tabName) {
  const ss = sheet.getParent();
  const target = ss.getSheetByName(tabName);
  if (!target) throw new Error('There is no "' + tabName + '" tab.');
  if (target.getSheetId() === sheet.getSheetId()) throw new Error('They are already on "' + tabName + '".');
  const tcols = pr_columns_(target);
  const number = cols.number ? String(sheet.getRange(row, cols.number).getDisplayValue()).trim() : '';
  const name = pr_nameAt_(sheet, cols, row);

  // Already in that tab? Then just refresh their rating there instead of adding a second row.
  const existing = pr_findApplicant_(target, tcols, number, name);
  if (existing) {
    const prev = {};
    ['reviewed', 'rating'].forEach(function (k) {
      const cell = target.getRange(existing, tcols[k]);
      prev[k] = String(cell.getDisplayValue());
      cell.setValue(sheet.getRange(row, cols[k]).getValue());
    });
    return { tab: tabName, row: existing, number: number, name: name, prev: prev };
  }

  const at = Math.max(1, pr_lastDataRow_(target, tcols)) + 1;
  if (at > target.getMaxRows()) target.insertRowsAfter(target.getMaxRows(), at - target.getMaxRows());
  if (pr_sameHeaders_(sheet, target, cols.lastCol)) {
    // Same columns: copy the whole row, dropdowns and formatting included.
    sheet.getRange(row, 1, 1, cols.lastCol).copyTo(target.getRange(at, 1, 1, cols.lastCol));
  } else {
    // Different layout: match columns by header.
    const head = sheet.getRange(1, 1, 1, cols.lastCol).getValues()[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim().toLowerCase(); });
    const vals = sheet.getRange(row, 1, 1, cols.lastCol).getValues()[0];
    const theads = target.getRange(1, 1, 1, tcols.lastCol).getValues()[0];
    const outRow = theads.map(function (h, i) {
      const k = head.indexOf(String(h).replace(/\s+/g, ' ').trim().toLowerCase());
      return k === -1 ? '' : vals[k];
    });
    target.getRange(at, 1, 1, outRow.length).setValues([outRow]);
  }
  return { tab: tabName, row: at, number: number, name: name, prev: null };
}

/**
 * Writes the row's Reviewed? and Rating into every other tab where this applicant already is
 * (skipping tabs in skip). Returns [{ tab, row, number, name, prev }] for undo.
 */
function pr_syncTabs_(sheet, cols, row, skip) {
  const number = cols.number ? String(sheet.getRange(row, cols.number).getDisplayValue()).trim() : '';
  const name = pr_nameAt_(sheet, cols, row);
  const want = pr_norm_(name);
  const reviewed = sheet.getRange(row, cols.reviewed).getValue();
  const rating = sheet.getRange(row, cols.rating).getValue();
  const out = [];
  if (!want) return out;
  sheet.getParent().getSheets().forEach(function (sh) {
    const tab = sh.getName();
    if (sh.getSheetId() === sheet.getSheetId() || (skip || []).indexOf(tab) !== -1) return;
    const data = sh.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    const t = pr_findCols_(data[0]);
    if (!t.name || !t.reviewed || !t.rating) return;
    for (let i = 1; i < data.length; i++) {
      if (pr_norm_(data[i][t.name - 1]) !== want) continue;
      if (number && t.number && String(data[i][t.number - 1]).trim() !== number) continue;
      const prev = { reviewed: data[i][t.reviewed - 1], rating: data[i][t.rating - 1] };
      if (prev.reviewed === String(reviewed) && prev.rating === String(rating)) continue;
      sh.getRange(i + 1, t.reviewed).setValue(reviewed);
      sh.getRange(i + 1, t.rating).setValue(rating);
      out.push({ tab: tab, row: i + 1, number: number, name: name, prev: prev });
    }
  });
  return out;
}

/** Last row with an applicant name (the sheet may have empty formatted rows below). */
function pr_lastDataRow_(sheet, cols) {
  const last = sheet.getLastRow();
  if (last < 2 || !cols.name) return last;
  const names = sheet.getRange(2, cols.name, last - 1, 1).getDisplayValues();
  for (let i = names.length - 1; i >= 0; i--) if (String(names[i][0]).trim()) return i + 2;
  return 1;
}

function pr_findApplicant_(sheet, cols, number, name) {
  const last = sheet.getLastRow();
  if (last < 2 || !cols.name) return 0;
  const names = sheet.getRange(2, cols.name, last - 1, 1).getDisplayValues();
  const nums = cols.number ? sheet.getRange(2, cols.number, last - 1, 1).getDisplayValues() : null;
  const want = pr_norm_(name);
  for (let i = 0; i < names.length; i++) {
    if (pr_norm_(names[i][0]) !== want) continue;
    if (nums && number && String(nums[i][0]).trim() !== number) continue;
    return i + 2;
  }
  return 0;
}

function pr_sameHeaders_(a, b, n) {
  if (b.getLastColumn() < n) return false;
  const ha = a.getRange(1, 1, 1, n).getDisplayValues()[0];
  const hb = b.getRange(1, 1, 1, n).getDisplayValues()[0];
  for (let i = 0; i < n; i++) if (pr_norm_(ha[i]) !== pr_norm_(hb[i])) return false;
  return true;
}

/** Undoes copies: rows we added are deleted, rows we refreshed get their old values back. */
function pr_removeCopies_(copies) {
  if (!copies.length) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  const done = [];
  try {
    copies.slice().reverse().forEach(function (c) {
      const target = ss.getSheetByName(c.tab);
      if (!target || c.row > target.getLastRow()) return;
      const tcols = pr_columns_(target);
      // Only touch the row if it's still that applicant (someone may have edited the tab since).
      const vals = target.getRange(c.row, 1, 1, tcols.lastCol).getDisplayValues()[0];
      if (pr_norm_(vals[tcols.name - 1]) !== pr_norm_(c.name)) return;
      if (c.number && tcols.number && String(vals[tcols.number - 1]).trim() !== c.number) return;
      if (c.prev) {
        Object.keys(c.prev).forEach(function (k) { if (tcols[k]) target.getRange(c.row, tcols[k]).setValue(c.prev[k]); });
      } else {
        target.deleteRow(c.row);
      }
      done.push(c.tab);
    });
  } finally {
    lock.releaseLock();
  }
  return done;
}

/* ---------------- Holding rows (two reviewers) and placing the sidebar ---------------- */

function pr_withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function pr_claimsKey_(sheet) {
  return 'pr_claims_' + sheet.getSheetId();
}

/** { row: { s: sidebar id, t: time } } for holds that haven't expired. */
function pr_claims_(sheet) {
  let c = {};
  try { c = JSON.parse(CacheService.getDocumentCache().get(pr_claimsKey_(sheet)) || '{}') || {}; } catch (e) { c = {}; }
  const now = Date.now();
  const ttl = PR_CONFIG.holdMinutes * 60000;
  Object.keys(c).forEach(function (r) { if (!c[r] || now - c[r].t > ttl) delete c[r]; });
  return c;
}

/** How many other sidebars are holding something on this tab right now. */
function pr_others_(claims, sid) {
  const seen = {};
  Object.keys(claims).forEach(function (r) { if (claims[r].s !== sid) seen[claims[r].s] = true; });
  return Object.keys(seen).length;
}

/**
 * Decides which row the sidebar shows, and holds it (plus the one after) for this sidebar.
 * how: { row } forces a row (Back, go to…); otherwise { showing } is kept if it's still open and
 * not held by someone else, else the next open row after { from }, else (with { edge }) the first
 * open row from the top or bottom. Rows other people hold are skipped.
 */
function pr_place_(sheet, cols, sid, dir, how) {
  return pr_withLock_(function () {
    const claims = pr_claims_(sheet);
    const blocked = {};
    Object.keys(claims).forEach(function (r) { if (claims[r].s !== sid) blocked[r] = true; });
    const scan = pr_scan_(sheet, cols);
    const free = function (r) { return r >= 2 && r <= scan.lastRow && scan.isOpen[r - 2] && !blocked[r]; };
    let row = 0;
    let changed = false;
    if (how.row) {
      row = how.row;
    } else {
      if (how.showing && free(how.showing)) {
        row = how.showing;
      } else {
        if (how.from) row = pr_nextFrom_(scan, how.from, dir, blocked);
        if (!row && how.edge) row = pr_edge_(scan, dir, blocked);
        changed = !!how.showing && row !== how.showing;
      }
    }
    const peek = row ? pr_nextFrom_(scan, row, dir, blocked) : 0;
    if (sid) {
      Object.keys(claims).forEach(function (r) { if (claims[r].s === sid) delete claims[r]; });
      const now = Date.now();
      if (row) claims[row] = { s: sid, t: now };
      if (peek) claims[peek] = { s: sid, t: now };
      CacheService.getDocumentCache().put(pr_claimsKey_(sheet), JSON.stringify(claims), 21600);
    }
    return { scan: scan, row: row, peek: peek, changed: changed, others: pr_others_(claims, sid) };
  });
}

/** What every call sends back to the sidebar. */
function pr_land_(sheet, cols, p) {
  if (p.row) pr_select_(sheet, cols, p.row);
  return {
    remaining: p.scan.remaining,
    current: p.row ? pr_rowInfo_(sheet, cols, p.row) : null,
    peek: p.peek ? pr_rowInfo_(sheet, cols, p.peek) : null,
    changed: p.changed,
    others: p.others,
    stats: pr_stats_(sheet),
  };
}

/* ---------------- Today's numbers ---------------- */

function pr_today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Toronto', 'yyyy-MM-dd');
}

function pr_statsKey_(sheet, day) {
  return 'pr_stats_' + sheet.getSheetId() + '_' + (day || pr_today_());
}

/** { total, by: { 'PEAK': 3, … } } for today (everyone reviewing this tab). */
function pr_stats_(sheet, day) {
  try {
    const s = JSON.parse(PropertiesService.getDocumentProperties().getProperty(pr_statsKey_(sheet, day)) || 'null');
    if (s && typeof s.total === 'number') return s;
  } catch (e) {
    // fall through
  }
  return { total: 0, by: {} };
}

/** Adds (or with -1 removes) one rating to a day's numbers. Call inside pr_withLock_. */
function pr_bump_(sheet, label, delta, day) {
  const s = pr_stats_(sheet, day);
  s.total = Math.max(0, s.total + delta);
  s.by[label] = Math.max(0, (s.by[label] || 0) + delta);
  if (!s.by[label]) delete s.by[label];
  PropertiesService.getDocumentProperties().setProperty(pr_statsKey_(sheet, day), JSON.stringify(s));
  return s;
}

/* ---------------- "go to …" ---------------- */

/** Lowercase, no accents. */
function pr_fold_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

/**
 * Finds a row for "go to …": "row 40" is that row; a number is the applicant number (or the row if
 * there's no number column); anything else matches names (all words, any order). With several
 * name matches it takes the next one after fromRow, so saying it again moves through them.
 */
function pr_find_(sheet, nameCols, numberCol, query, fromRow, dir) {
  const lastRow = sheet.getLastRow();
  const raw = String(query || '').trim().replace(/[.!?]+$/, '');
  const q = raw.replace(/^(?:the\s+)?(?:applicant|audition|number|no\.?|#)\s*(?:number\s*|#\s*)?/i, '');
  const asRow = q.match(/^row\s*(\d[\d\s,]*)$/i);
  if (asRow) {
    const r = Number(asRow[1].replace(/\D/g, ''));
    if (r < 2 || r > lastRow) throw new Error('There’s no row ' + r + '.');
    return { row: r, how: 'row', count: 1, index: 1 };
  }
  const digits = q.replace(/[\s,]/g, '');
  if (/^\d+$/.test(digits)) {
    if (numberCol && lastRow >= 2) {
      const nums = sheet.getRange(2, numberCol, lastRow - 1, 1).getDisplayValues();
      for (let i = 0; i < nums.length; i++) {
        if (String(nums[i][0]).replace(/\D/g, '') === digits) return { row: i + 2, how: 'number', count: 1, index: 1 };
      }
    }
    const r = Number(digits);
    if (r >= 2 && r <= lastRow) return { row: r, how: 'row', count: 1, index: 1 };
    throw new Error('Couldn’t find number ' + digits + '.');
  }
  const words = pr_fold_(q).split(/[^a-z0-9']+/).filter(Boolean);
  if (!words.length || lastRow < 2) throw new Error('Say who to go to, like “go to 722” or “go to Alyssa”.');
  const cols = nameCols.filter(Boolean);
  const vals = cols.map(function (c) { return sheet.getRange(2, c, lastRow - 1, 1).getDisplayValues(); });
  const matches = [];
  for (let i = 0; i < lastRow - 1; i++) {
    const hay = pr_fold_(vals.map(function (v) { return v[i][0]; }).join(' '));
    if (words.every(function (w) { return hay.indexOf(w) !== -1; })) matches.push(i + 2);
  }
  if (!matches.length) throw new Error('Couldn’t find anyone called “' + raw + '”.');
  const order = dir === 'up' ? matches.slice().reverse() : matches;
  let row = order.find(function (r) { return dir === 'up' ? r < fromRow : r > fromRow; });
  if (!row) row = order[0];
  return { row: row, how: 'name', count: matches.length, index: order.indexOf(row) + 1 };
}

/* ---------------- Shortlist ---------------- */

function pr_buildShortlistMenu() {
  const res = pr_buildShortlist();
  SpreadsheetApp.getActiveSpreadsheet().toast(res.summary, 'Shortlist');
}

/**
 * Writes the Shortlist tab: for each category tab, its PEAK then Not bad applicants with their
 * portfolio and notes, then rated PEAK / Not bad applicants who aren't in any tab yet.
 */
function pr_buildShortlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(PR_CONFIG.mainSheet) || ss.getActiveSheet();
  const mcols = pr_columns_(main);
  const options = pr_options_(main, mcols, 2);
  const want = function (cmd) {
    const r = PR_CONFIG.ratings.filter(function (x) { return x.cmd === cmd; })[0];
    return r ? pr_norm_(options.length ? (pr_resolve_(options, r.option) || r.option) : r.option) : '';
  };
  const peakKey = want('peak');
  const goodKey = want('good');
  const rank = function (v) { const n = pr_norm_(v); return n && n === peakKey ? 0 : n && n === goodKey ? 1 : -1; };
  const keyOf = function (num, name) { return String(num).trim() + '|' + pr_norm_(name); };

  const pick = function (data, t) {
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const rk = rank(r[t.rating - 1]);
      if (rk < 0 || !String(r[t.name - 1]).trim()) continue;
      rows.push({ rank: rk, row: r });
    }
    rows.sort(function (a, b) {
      return a.rank - b.rank || (Number(a.row[t.number - 1]) || 0) - (Number(b.row[t.number - 1]) || 0);
    });
    return rows;
  };

  const sections = [];
  const inTab = {};
  pr_tabs_(ss, main).forEach(function (tab) {
    const sh = ss.getSheetByName(tab);
    const data = sh.getDataRange().getDisplayValues();
    if (data.length < 1) return;
    const t = pr_findCols_(data[0]);
    if (!t.name || !t.rating) return;
    for (let i = 1; i < data.length; i++) inTab[keyOf(t.number ? data[i][t.number - 1] : '', data[i][t.name - 1])] = true;
    sections.push({ title: tab, t: t, rows: pick(data, t) });
  });
  const mdata = main.getDataRange().getDisplayValues();
  const mt = pr_findCols_(mdata[0]);
  const loose = pick(mdata, mt).filter(function (x) {
    return !inTab[keyOf(mt.number ? x.row[mt.number - 1] : '', x.row[mt.name - 1])];
  });
  sections.push({ title: 'Not in a category tab yet', t: mt, rows: loose });

  const HEAD = ['Rating', 'Applicant #', 'Name', 'Portfolio', 'Instagram', 'Roles', 'Notes'];
  const W = HEAD.length;
  const blank = function () { return HEAD.map(function () { return ''; }); };
  const q = function (s) { return String(s).replace(/"/g, '""'); };
  const link = function (url, label) { return url ? '=HYPERLINK("' + q(url) + '","' + q(label || url) + '")' : ''; };
  const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Toronto', 'MMM d, yyyy h:mm a');
  const out = [];
  const marks = { title: [], head: [], peak: [], good: [] };
  let nPeak = 0;
  let nGood = 0;
  out.push(['Shortlist: PEAK and Not bad, by tab (updated ' + day + ')'].concat(blank().slice(1)));
  out.push(blank());
  sections.forEach(function (sec) {
    marks.title.push(out.length + 1);
    out.push([sec.title + ' (' + sec.rows.length + ')'].concat(blank().slice(1)));
    if (!sec.rows.length) { out.push(['Nobody rated PEAK or Not bad here yet.'].concat(blank().slice(1))); out.push(blank()); return; }
    marks.head.push(out.length + 1);
    out.push(HEAD.slice());
    sec.rows.forEach(function (x) {
      const t = sec.t;
      const r = x.row;
      const get = function (c) { return c ? String(r[c - 1] || '').trim() : ''; };
      const portfolio = get(t.portfolio);
      const first = pr_links_(portfolio)[0] || pr_handleLink_(portfolio);
      const insta = pr_instagram_(get(t.instagram));
      (x.rank === 0 ? marks.peak : marks.good).push(out.length + 1);
      if (x.rank === 0) nPeak++; else nGood++;
      out.push([
        get(t.rating), get(t.number), get(t.name),
        first ? link(first.url, first.label) : portfolio.slice(0, 200),
        insta ? link(insta, get(t.instagram).replace(/^@?/, '@')) : '',
        get(t.roles), get(t.notes),
      ]);
    });
    out.push(blank());
  });

  let sh = ss.getSheetByName(PR_CONFIG.shortlistSheet);
  if (sh) sh.clear(); else sh = ss.insertSheet(PR_CONFIG.shortlistSheet);
  // Text columns as plain text (notes like "- nice colours" would otherwise turn into formulas).
  [1, 2, 3, 6, 7].forEach(function (c) { sh.getRange(1, c, out.length, 1).setNumberFormat('@'); });
  sh.getRange(1, 1, out.length, W).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  marks.title.forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground('#e8f0fe'); });
  marks.head.forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground('#f1f3f4'); });
  marks.peak.forEach(function (r) { sh.getRange(r, 1).setBackground('#b6d7a8'); });
  marks.good.forEach(function (r) { sh.getRange(r, 1).setBackground('#ffe599'); });
  sh.getRange(1, 7, out.length, 1).setWrap(true);
  [90, 90, 180, 200, 140, 200, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  sh.activate();
  const summary = nPeak + ' PEAK and ' + nGood + ' Not bad across ' + (sections.length - 1) + ' tabs (' + loose.length + ' not in a tab yet).';
  return { sheetName: sh.getName(), peak: nPeak, good: nGood, loose: loose.length, summary: summary };
}

/* ---------------- Helpers ---------------- */

function pr_prefKey_(key) {
  return 'pr_' + key + '_' + SpreadsheetApp.getActiveSpreadsheet().getId();
}

function pr_getPref_(key) {
  return PropertiesService.getUserProperties().getProperty(pr_prefKey_(key));
}

function pr_norm_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** The Rating column's dropdown options (exact text), or the values already used if there's no dropdown. */
function pr_options_(sheet, cols, row) {
  const list = [];
  const add = function (v) {
    const s = String(v == null ? '' : v);
    if (s.trim() && list.indexOf(s) === -1) list.push(s);
  };
  const r = Math.max(2, Math.min(row || 2, Math.max(2, sheet.getLastRow())));
  const dv = sheet.getRange(r, cols.rating).getDataValidation();
  if (dv) {
    const type = dv.getCriteriaType();
    const vals = dv.getCriteriaValues() || [];
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST && vals[0]) {
      vals[0].forEach(add);
    } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && vals[0] && vals[0].getDisplayValues) {
      vals[0].getDisplayValues().forEach(function (rowVals) { rowVals.forEach(add); });
    }
  }
  if (!list.length && sheet.getLastRow() >= 2) {
    sheet.getRange(2, cols.rating, sheet.getLastRow() - 1, 1).getValues().forEach(function (v) { add(v[0]); });
  }
  return list;
}

/** Finds the option that matches the wanted text (loosely), or null. */
function pr_resolve_(options, wanted) {
  const w = pr_norm_(wanted);
  if (!w) return null;
  for (let i = 0; i < options.length; i++) if (pr_norm_(options[i]) === w) return options[i];
  for (let i = 0; i < options.length; i++) {
    const o = pr_norm_(options[i]);
    if (w.length >= 3 && o.indexOf(w) === 0) return options[i]; // e.g. "Not bad" → "Not bad!"
  }
  return null;
}

/** Buttons for the sidebar, in 1 / 2 / 3 order: the configured ratings, then any other dropdown options. */
function pr_choices_(options) {
  const out = [];
  const used = {};
  PR_CONFIG.ratings.forEach(function (r) {
    const opt = options.length ? pr_resolve_(options, r.option) : r.option;
    if (opt === null || used[opt]) return;
    used[opt] = true;
    out.push({ cmd: r.cmd, option: opt, label: String(opt).trim(), key: r.key });
  });
  options.forEach(function (o) {
    if (used[o]) return;
    used[o] = true;
    out.push({ cmd: 'option', option: o, label: String(o).trim(), key: '' });
  });
  return out;
}

/** What a command writes: { reviewed, rating }. */
function pr_spec_(cmd, value, options) {
  const choices = pr_choices_(options);
  const name = PR_CONFIG.aliases[cmd] || cmd;
  const pick = function (wanted) {
    const opt = options.length ? pr_resolve_(options, wanted) : wanted;
    if (opt === null) {
      throw new Error('The Rating dropdown has no "' + wanted + '" option. Options: ' +
        options.map(function (o) { return String(o).trim(); }).join(', ') + '. Change PR_CONFIG.ratings to match.');
    }
    return opt;
  };

  if (name === 'na') return { reviewed: PR_CONFIG.na.reviewed, rating: pick(PR_CONFIG.na.option) };
  for (let i = 0; i < PR_CONFIG.ratings.length; i++) {
    if (PR_CONFIG.ratings[i].cmd === name) return { reviewed: PR_CONFIG.reviewedValue, rating: pick(PR_CONFIG.ratings[i].option) };
  }
  if (name === 'score') {
    const n = Number(value);
    if (n >= 1 && n <= choices.length && Math.round(n) === n) return { reviewed: PR_CONFIG.reviewedValue, rating: choices[n - 1].option };
    throw new Error('Say a number from 1 to ' + choices.length + ' (' +
      choices.map(function (c, i) { return (i + 1) + ' = ' + c.label; }).join(', ') + ').');
  }
  if (name === 'option') return { reviewed: PR_CONFIG.reviewedValue, rating: pick(value) };
  throw new Error('Unknown rating: ' + cmd);
}

function pr_sheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found. Reopen the reviewer.');
  return sheet;
}

function pr_checkRow_(sheet, row) {
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Row ' + row + ' is out of range.');
}

function pr_columns_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const cols = pr_findCols_(sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]);
  const missing = ['reviewed', 'rating', 'notes', 'portfolio'].filter(function (k) { return !cols[k]; });
  if (missing.length) {
    throw new Error('Couldn\'t find the ' + missing.join(', ') + ' column(s) on "' + sheet.getName() +
      '". Check the header names in PR_CONFIG.headers.');
  }
  return cols;
}

/** Column numbers (1-based, 0 = missing) for each PR_CONFIG.headers key, from a header row. */
function pr_findCols_(headerRow) {
  const headers = headerRow.map(function (h) {
    return String(h).replace(/\s+/g, ' ').trim().toLowerCase();
  });
  const find = function (candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = headers.indexOf(candidates[i].toLowerCase());
      if (idx !== -1) return idx + 1;
    }
    for (let i = 0; i < candidates.length; i++) {
      const c = candidates[i].toLowerCase();
      const idx = headers.findIndex(function (h) { return h.indexOf(c) === 0; });
      if (idx !== -1) return idx + 1;
    }
    return 0;
  };
  const cols = { lastCol: headerRow.length };
  Object.keys(PR_CONFIG.headers).forEach(function (key) { cols[key] = find(PR_CONFIG.headers[key]); });
  return cols;
}

/** One pass over the sheet: which rows still need reviewing. */
function pr_scan_(sheet, cols) {
  const lastRow = sheet.getLastRow();
  const n = Math.max(0, lastRow - 1);
  const isOpen = new Array(n);
  let remaining = 0;
  if (!n) return { isOpen: isOpen, remaining: 0, lastRow: lastRow };
  const reviewed = sheet.getRange(2, cols.reviewed, n, 1).getDisplayValues();
  const names = cols.name ? sheet.getRange(2, cols.name, n, 1).getDisplayValues() : null;
  for (let i = 0; i < n; i++) {
    const open = String(reviewed[i][0]).trim() === '';
    const hasPerson = !names || String(names[i][0]).trim() !== '';
    isOpen[i] = open && hasPerson;
    if (isOpen[i]) remaining++;
  }
  return { isOpen: isOpen, remaining: remaining, lastRow: lastRow };
}

/** Next open row after fromRow (exclusive) going 'down' or 'up', skipping blocked rows. 0 if none. */
function pr_nextFrom_(scan, fromRow, dir, blocked) {
  const step = dir === 'up' ? -1 : 1;
  for (let r = fromRow + step; r >= 2 && r <= scan.lastRow; r += step) {
    if (scan.isOpen[r - 2] && !(blocked && blocked[r])) return r;
  }
  return 0;
}

/** First open row from the top (dir 'down') or from the bottom (dir 'up'). */
function pr_edge_(scan, dir, blocked) {
  return dir === 'up' ? pr_nextFrom_(scan, scan.lastRow + 1, 'up', blocked) : pr_nextFrom_(scan, 1, 'down', blocked);
}

function pr_rowInfo_(sheet, cols, row) {
  const vals = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
  const get = function (c) { return c ? String(vals[c - 1] || '').trim() : ''; };
  const portfolio = get(cols.portfolio);
  const insta = pr_instagram_(get(cols.instagram));
  const links = pr_links_(portfolio);
  const fromHandle = links.length ? null : pr_handleLink_(portfolio);
  if (!links.length && fromHandle) links.push(fromHandle);
  if (!links.length && insta) links.push({ url: insta, label: 'Instagram', kind: 'instagram' });
  const about = get(cols.about);
  return {
    row: row,
    number: get(cols.number),
    name: get(cols.name) || '(no name)',
    reviewed: get(cols.reviewed),
    rating: get(cols.rating),
    notes: get(cols.notes),
    roles: get(cols.roles),
    country: get(cols.country),
    about: about.length > 900 ? about.slice(0, 900) + '…' : about,
    portfolio: portfolio,
    instagram: insta,
    links: links,
  };
}

function pr_nameAt_(sheet, cols, row) {
  return cols.name ? String(sheet.getRange(row, cols.name).getDisplayValue()).trim() : 'Row ' + row;
}

function pr_select_(sheet, cols, row) {
  sheet.getRange(row, cols.name || 1).activate();
}

const PR_TLDS_ = ('com net org co io me art design site page app dev xyz info us ca uk ph au nz de fr es it nl jp kr sg my ' +
  'gallery website studio online space store portfolio link ink ly gg tv blog work works pro games cc biz graphics ' +
  'digital ai to so one world life top club fun live social myportfolio carrd crd wixsite framer').split(' ');

/** Every web link in the portfolio cell, in order: [{ url, label, kind }]. */
function pr_links_(text) {
  const out = [];
  const seen = {};
  String(text || '').split(/[\s,;|<>"“”]+/).forEach(function (raw) {
    let tok = raw;
    const at = tok.search(/https?:\/\//i);
    if (at > 0) tok = tok.slice(at);
    tok = tok.replace(/^[(\[{'‘]+/, '').replace(/[)\]}.,;:!?'’]+$/, '');
    if (!tok) return;
    let url = null;
    if (/^https?:\/\/[^\s/]+\.[^\s/]+/i.test(tok)) {
      url = tok;
    } else if (!/^[^/]*@/.test(tok) && /^(?:www\.)?(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+([a-z]{2,24})(?:[/?#]\S*)?$/i.test(tok)) {
      const tld = tok.replace(/^[^/?#]*?\.([a-z]{2,24})(?:[/?#].*)?$/i, '$1').toLowerCase();
      if (PR_TLDS_.indexOf(tld) !== -1) url = 'https://' + tok;
    }
    if (!url) return;
    const key = url.toLowerCase().replace(/^https?:\/\/(www\.)?/, '').replace(/\/+$/, '');
    if (seen[key]) return;
    seen[key] = true;
    out.push({ url: url, label: pr_label_(url), kind: 'link' });
  });
  return out;
}

/** "@handle" in the portfolio cell (when there's no link) → their Instagram. */
function pr_handleLink_(text) {
  const m = String(text || '').match(/(?:^|[\s(:])@([A-Za-z0-9._]{2,30})\b/);
  if (!m) return null;
  return { url: 'https://www.instagram.com/' + m[1].replace(/\.$/, '') + '/', label: 'Instagram @' + m[1], kind: 'instagram' };
}

/** The Instagram column as a link, or '' if it's empty / not a handle. */
function pr_instagram_(value) {
  const v = String(value || '').trim();
  if (!v) return '';
  const link = pr_links_(v)[0];
  if (link && /instagram\.com/i.test(link.url)) return link.url;
  const m = v.match(/^@?([A-Za-z0-9._]{2,30})$/);
  if (!m || /^n\/?a$/i.test(v) || /^none$/i.test(v)) return '';
  return 'https://www.instagram.com/' + m[1] + '/';
}

function pr_label_(url) {
  const m = String(url).match(/^https?:\/\/(?:www\.)?([^/?#]+)/i);
  return m ? m[1].toLowerCase() : url;
}
