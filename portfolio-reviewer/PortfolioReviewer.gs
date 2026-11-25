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
    .addToUi();
}

function pr_openSidebar() {
  const t = HtmlService.createTemplateFromFile('PortfolioSidebar');
  t.clientConfig = JSON.stringify({ voiceUrl: PR_CONFIG.voiceUrl, noteUrl: PR_CONFIG.noteUrl });
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Portfolio Reviewer'));
}

/* ---------------- Called from the sidebar ---------------- */

/** Opens on the selected row if it hasn't been reviewed yet, otherwise the next open one. */
function pr_init() {
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
  let row = 0;
  if (active >= 2 && active <= scan.lastRow && scan.isOpen[active - 2]) {
    row = active;
  } else if (active >= 2 && active <= scan.lastRow) {
    row = pr_nextFrom_(scan, active, dir);
  }
  if (!row) row = pr_edge_(scan, dir);
  if (row) pr_select_(sheet, cols, row);

  return {
    sheetName: sheet.getName(),
    title: ss.getName(),
    tabs: pr_tabs_(ss, sheet),
    dir: dir,
    view: pr_getPref_('view') === 'popup' ? 'popup' : 'tab',
    remaining: scan.remaining,
    choices: pr_choices_(pr_options_(sheet, cols, row || 2)),
    current: row ? pr_rowInfo_(sheet, cols, row) : null,
  };
}

/**
 * Rates the row, copies it into any tabs in addTo (e.g. ['Layout']),
 * then returns the next open applicant in the current direction.
 */
function pr_rate(sheetName, row, cmd, dir, value, addTo) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  const spec = pr_spec_(cmd, value, pr_options_(sheet, cols, row));

  const reviewedCell = sheet.getRange(row, cols.reviewed);
  const ratingCell = sheet.getRange(row, cols.rating);
  const undo = { kind: 'rate', row: row, reviewed: reviewedCell.getValue(), rating: ratingCell.getValue(), copies: [] };

  reviewedCell.setValue(spec.reviewed);
  ratingCell.setValue(spec.rating);
  const added = pr_copyAll_(sheet, cols, row, addTo);
  const synced = pr_syncTabs_(sheet, cols, row, added.map(function (a) { return a.tab; }));
  undo.copies = added.concat(synced);

  const scan = pr_scan_(sheet, cols);
  const nextRow = pr_nextFrom_(scan, row, dir);
  if (nextRow) pr_select_(sheet, cols, nextRow);

  return {
    rated: { row: row, name: pr_nameAt_(sheet, cols, row), rating: String(spec.rating).trim(), reviewed: spec.reviewed, cmd: cmd },
    added: added.map(function (a) { return { tab: a.tab, updated: !!a.prev }; }),
    synced: synced.map(function (a) { return a.tab; }),
    undo: undo,
    remaining: scan.remaining,
    next: nextRow ? pr_rowInfo_(sheet, cols, nextRow) : null,
  };
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
function pr_skip(sheetName, row, dir) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const scan = pr_scan_(sheet, cols);
  const nextRow = pr_nextFrom_(scan, row, dir);
  if (nextRow) pr_select_(sheet, cols, nextRow);
  return { remaining: scan.remaining, next: nextRow ? pr_rowInfo_(sheet, cols, nextRow) : null };
}

/**
 * Reverts a rating or an "add to" (using the undo object pr_rate / pr_addTo returned):
 * copies it made are removed (or put back as they were), and a rating goes back to that row.
 */
function pr_undo(sheetName, undo) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, undo.row);
  const removed = pr_removeCopies_(undo.copies || []);
  if (undo.kind === 'add') {
    return { removed: removed, remaining: pr_scan_(sheet, cols).remaining, current: pr_rowInfo_(sheet, cols, undo.row) };
  }
  sheet.getRange(undo.row, cols.reviewed).setValue(undo.reviewed === null ? '' : undo.reviewed);
  sheet.getRange(undo.row, cols.rating).setValue(undo.rating === null ? '' : undo.rating);
  const scan = pr_scan_(sheet, cols);
  pr_select_(sheet, cols, undo.row);
  return { removed: removed, remaining: scan.remaining, current: pr_rowInfo_(sheet, cols, undo.row) };
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
function pr_jump(sheetName, where) {
  const dir = where === 'bottom' ? 'up' : 'down';
  pr_setPref('dir', dir);
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const scan = pr_scan_(sheet, cols);
  const row = pr_edge_(scan, dir);
  if (row) pr_select_(sheet, cols, row);
  return { dir: dir, remaining: scan.remaining, current: row ? pr_rowInfo_(sheet, cols, row) : null };
}

/** Use whatever row is selected in the sheet right now (even if already reviewed). */
function pr_startHere(sheetName) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  const range = sheet.getActiveRange();
  const row = range && range.getSheet().getSheetId() === sheet.getSheetId() ? range.getRow() : 0;
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Select a row with an applicant on "' + sheetName + '" first.');
  const scan = pr_scan_(sheet, cols);
  return { remaining: scan.remaining, current: pr_rowInfo_(sheet, cols, row) };
}

/** Opens a specific row without changing anything (used by Back). */
function pr_rowAt(sheetName, row) {
  const sheet = pr_sheet_(sheetName);
  const cols = pr_columns_(sheet);
  pr_checkRow_(sheet, row);
  const scan = pr_scan_(sheet, cols);
  pr_select_(sheet, cols, row);
  return { remaining: scan.remaining, current: pr_rowInfo_(sheet, cols, row) };
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

/** Next open row after fromRow (exclusive) going 'down' or 'up'. 0 if none. */
function pr_nextFrom_(scan, fromRow, dir) {
  const step = dir === 'up' ? -1 : 1;
  for (let r = fromRow + step; r >= 2 && r <= scan.lastRow; r += step) {
    if (scan.isOpen[r - 2]) return r;
  }
  return 0;
}

/** First open row from the top (dir 'down') or from the bottom (dir 'up'). */
function pr_edge_(scan, dir) {
  return dir === 'up' ? pr_nextFrom_(scan, scan.lastRow + 1, 'up') : pr_nextFrom_(scan, 1, 'down');
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
