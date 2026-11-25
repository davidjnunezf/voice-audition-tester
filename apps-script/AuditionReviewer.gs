/**
 * Audition Reviewer — Studio Tiny Tuko
 * ------------------------------------------------------------
 * Sidebar + voice remote for reviewing audition response sheets.
 *
 *   next       → rating 1, Listened? = Yes, open next audition
 *   good       → rating 3, row highlighted yellow, open next
 *   peak       → rating 5, row highlighted green, open next
 *   custom 3.5 → any score 0–5 ("custom four and a half", or just say the number)
 *                Colour from the score: below 3 none, 3+ yellow, 4+ green
 *   N/A        → Listened? = No, both ratings = N/A, row highlighted red, open next
 *   notes ...  → appends "- ..." to the Notes cell (stays on the same audition)
 *   descending → work top → bottom      ascending → work bottom → top
 *   back       → goes back to the previous audition (rating unchanged)
 *   undo       → reverts the last rating and goes back to that audition
 *   skip       → next audition without rating this one
 *   go to 722 / go to Maria → jump to that row (or audition #) / person (again for the next match)
 *   progress   → says how many are left, done today, and how long the rest will take
 *
 * Two people can review the same sheet at once: each sidebar holds the audition it's on (and the
 * next one) for 10 minutes and the other skips past them. The next audition shows immediately;
 * the rating is saved in the background.
 *
 * Columns are found by header name, so the same code works in every
 * role sheet (Tama, Kiko, Rinda, Duwende, Okrasio, Support Roles…).
 * All functions are prefixed with ar_ so they won't clash with your
 * existing Apps Script code.
 */

const AR_CONFIG = {
  // Header names to look for (case-insensitive). Exact matches win, then "starts with".
  headers: {
    name:     ['Full Name'],
    number:   ['Audition #', 'Applicant #', 'Submission #'],
    alias:    ['Preferred Name/Alias', 'Preferred Name'],
    listened: ['Listened?', 'Listened'],
    rating:   ['General Rating', 'Rating'],
    roleRating: ['Role Rating'],
    notes:    ['Notes'],
    upload:   ['Submit Your Audition'],
    link:     ['Submitting with a link'],
  },

  // Shortcut commands. A score gets its row colour from colorBands below.
  commands: {
    next:      { score: 1 },
    good:      { score: 3 },
    peak:      { score: 5 },
    excellent: { score: 5 }, // old name for peak
    na:        { listened: 'No', rating: 'N/A', roleRating: 'N/A', color: '#ea9999' }, // light red
  },

  // Row colour by score: the first band the score reaches. Below 3 = no highlight.
  colorBands: [
    { min: 4, color: '#b6d7a8' }, // light green  (4, 4.5, 5)
    { min: 3, color: '#ffe599' }, // light yellow (3, 3.5)
  ],

  listenedValue: 'Yes',

  // How long a sidebar holds the audition it's on, so a second reviewer skips it.
  holdMinutes: 10,

  // Where you host voice.html (GitHub Pages or any HTTPS site). See README.
  voiceUrl: 'https://davidjnunezf.github.io/voice-audition-tester/voice.html',

  // Audio files bigger than this play through Drive's previewer instead of inline.
  maxInlineMb: 20,
};

/* ---------------- Menu & sidebar ---------------- */

// If your project already has an onOpen(), delete this one and add
// ar_addMenu(); inside yours instead.
function onOpen() {
  ar_addMenu();
}

function ar_addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🎙️ Auditions')
    .addItem('Open reviewer', 'ar_openSidebar')
    .addToUi();
}

function ar_openSidebar() {
  const t = HtmlService.createTemplateFromFile('AuditionSidebar');
  t.clientConfig = JSON.stringify({
    voiceUrl: AR_CONFIG.voiceUrl,
    commands: AR_CONFIG.commands,
  });
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Audition Reviewer'));
}

/* ---------------- Called from the sidebar ---------------- */

/*
 * opts = { sid, showing }: sid identifies this sidebar (for holds); showing is the audition the
 * sidebar already moved to on its own (instant next), kept if it's still free.
 * Calls return { current, peek, remaining, others, stats }; peek is the audition after current.
 */

/** Opens on the selected row if it hasn't been listened to yet, otherwise the next open one. */
function ar_init(opts) {
  opts = opts || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cols = ar_columns_(sheet);
  const dir = ar_getDir_();
  const range = sheet.getActiveRange();
  const active = range ? range.getRow() : 0;
  const at = active >= 2 && active <= sheet.getLastRow() ? active : 0;
  const out = ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { showing: at, from: at, edge: true }));
  out.sheetName = sheet.getName();
  out.title = ss.getName();
  out.dir = dir;
  return out;
}

/** Rates the row, then moves on to the next open audition in the current direction. */
function ar_rate(sheetName, row, cmd, dir, value, opts) {
  opts = opts || {};
  const spec = ar_spec_(cmd, value);
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, row);

  const rowRange = sheet.getRange(row, 1, 1, cols.lastCol);
  const listenedCell = sheet.getRange(row, cols.listened);
  const ratingCell = sheet.getRange(row, cols.rating);
  const roleCell = cols.roleRating ? sheet.getRange(row, cols.roleRating) : null;
  const backgrounds = rowRange.getBackgrounds()[0];
  const label = spec.listened === AR_CONFIG.commands.na.listened && spec.rating === AR_CONFIG.commands.na.rating ? 'N/A' : String(spec.rating);

  const undo = {
    row: row,
    listened: listenedCell.getValue(),
    rating: ratingCell.getValue(),
    roleRating: roleCell ? roleCell.getValue() : null,
    backgrounds: backgrounds,
    label: label,
    day: ar_today_(),
    wasOpen: String(listenedCell.getDisplayValue()).trim() === '',
  };

  listenedCell.setValue(spec.listened);
  ratingCell.setValue(spec.rating);
  if (roleCell && spec.roleRating !== undefined) roleCell.setValue(spec.roleRating);

  if (spec.color) {
    rowRange.setBackground(spec.color);
  } else {
    // Low score: only clear a highlight this tool added earlier.
    const ours = ar_ourColors_();
    const hasOurs = backgrounds.some(function (b) { return ours.indexOf(String(b).toLowerCase()) !== -1; });
    if (hasOurs) rowRange.setBackground(null);
  }
  if (undo.wasOpen) ar_withLock_(function () { ar_bump_(sheet, label, 1); });

  const out = ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { from: row, showing: opts.showing }));
  out.rated = { row: row, name: ar_nameAt_(sheet, cols, row), rating: String(spec.rating), cmd: cmd, label: label };
  out.undo = undo;
  return out;
}

/** Moves on without rating. */
function ar_skip(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  return ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { from: row, showing: opts.showing }));
}

/** Reverts a rating (using the undo object ar_rate returned) and goes back to that row. */
function ar_undo(sheetName, undo, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, undo.row);

  sheet.getRange(undo.row, cols.listened).setValue(undo.listened === null ? '' : undo.listened);
  sheet.getRange(undo.row, cols.rating).setValue(undo.rating === null ? '' : undo.rating);
  if (cols.roleRating && undo.roleRating !== undefined) {
    sheet.getRange(undo.row, cols.roleRating).setValue(undo.roleRating === null ? '' : undo.roleRating);
  }

  const bgs = (undo.backgrounds || []).map(function (b) {
    return !b || String(b).toLowerCase() === '#ffffff' ? null : b;
  });
  if (bgs.length) sheet.getRange(undo.row, 1, 1, bgs.length).setBackgrounds([bgs]);
  if (undo.wasOpen && undo.label) ar_withLock_(function () { ar_bump_(sheet, undo.label, -1, undo.day); });

  return ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { row: undo.row }));
}

/** Appends "- text" to the Notes cell. */
function ar_addNote(sheetName, row, text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean) throw new Error('Empty note');
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, row);

  const cell = sheet.getRange(row, cols.notes);
  const current = String(cell.getDisplayValue() || '').replace(/\s+$/, '');
  const line = '- ' + clean;
  cell.setNumberFormat('@').setValue(current ? current + '\n' + line : line);
  return { row: row, notes: cell.getDisplayValue() };
}

/** Next open audition after row (doesn't change anything). */
function ar_peek(sheetName, row, dir) {
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const scan = ar_scan_(sheet, cols);
  const nextRow = ar_nextFrom_(scan, row, dir);
  return nextRow ? ar_rowInfo_(sheet, cols, nextRow) : null;
}

/** Keeps holding this row (and the next) without moving the sheet's selection. */
function ar_hold(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const p = ar_place_(sheet, cols, opts.sid, dir, { row: row });
  return {
    remaining: p.scan.remaining, others: p.others, stats: ar_stats_(sheet),
    current: { row: row }, peek: p.peek ? ar_rowInfo_(sheet, cols, p.peek) : null,
  };
}

/** Jump to the first open audition from the top ('top') or bottom ('bottom'). */
function ar_jump(sheetName, where, opts) {
  opts = opts || {};
  const dir = where === 'bottom' ? 'up' : 'down';
  ar_setDir(dir);
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const out = ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { edge: true }));
  out.dir = dir;
  return out;
}

/** Use whatever row is selected in the sheet right now (even if already reviewed). */
function ar_startHere(sheetName, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 0;
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Select a row with an audition first.');
  return ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { row: row }));
}

/** Opens a specific row without changing anything (used by Back). */
function ar_rowAt(sheetName, row, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, row);
  return ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { row: row }));
}

/** "go to 722" (the row, or the audition # if the sheet has one) / "go to Maria" (name or alias). */
function ar_goto(sheetName, query, fromRow, dir, opts) {
  opts = opts || {};
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const hit = ar_find_(sheet, [cols.name, cols.alias], cols.number, query, fromRow, dir);
  const out = ar_land_(sheet, cols, ar_place_(sheet, cols, opts.sid, dir, { row: hit.row }));
  out.found = hit;
  return out;
}

/** Today's numbers for the progress panel (and the "progress" voice command). */
function ar_progress(sheetName, opts) {
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  return { remaining: ar_scan_(sheet, cols).remaining, stats: ar_stats_(sheet), others: ar_others_(ar_claims_(sheet), (opts || {}).sid) };
}

function ar_setDir(dir) {
  PropertiesService.getUserProperties().setProperty(ar_dirKey_(), dir === 'up' ? 'up' : 'down');
  return dir === 'up' ? 'up' : 'down';
}

/** Loads a Drive audio/video file (or the first one inside a shared folder) for the sidebar player. */
function ar_getAudio(fileId) {
  try {
    const picked = ar_pickFile_(fileId);
    const file = picked.file;
    if (!file) {
      return { ok: false, mode: 'preview', id: fileId, folderCount: picked.folderCount, error: 'No audio or video files in this folder' };
    }
    let mime = file.getMimeType() || '';
    const size = file.getSize();
    if (mime === 'audio/x-m4a' || mime === 'audio/m4a') mime = 'audio/mp4';
    const base = { ok: true, id: fileId, fileId: file.getId(), mime: mime, name: file.getName(), folderCount: picked.folderCount || 0 };
    const playable = /^(audio|video)\//.test(mime);
    if (!playable || size > AR_CONFIG.maxInlineMb * 1024 * 1024) {
      base.mode = 'preview';
      return base;
    }
    base.mode = 'inline';
    base.b64 = Utilities.base64Encode(file.getBlob().getBytes());
    return base;
  } catch (e) {
    return { ok: false, mode: 'preview', id: fileId, error: String((e && e.message) || e) };
  }
}

/** A file ID gives that file; a folder ID gives the first audio/video file inside it. */
function ar_pickFile_(id) {
  try {
    const f = DriveApp.getFileById(id);
    if (f.getMimeType() !== MimeType.FOLDER) return { file: f };
  } catch (e) {
    // Not a file we can open directly; try it as a folder below.
  }
  const it = DriveApp.getFolderById(id).getFiles();
  let pick = null;
  let count = 0;
  while (it.hasNext()) {
    const f = it.next();
    count++;
    if (!pick && /^(audio|video)\//.test(f.getMimeType())) pick = f;
  }
  return { file: pick, folderCount: count };
}

/* ---------------- Holding rows (two reviewers) and placing the sidebar ---------------- */

function ar_withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function ar_claimsKey_(sheet) {
  return 'ar_claims_' + sheet.getSheetId();
}

/** { row: { s: sidebar id, t: time } } for holds that haven't expired. */
function ar_claims_(sheet) {
  let c = {};
  try { c = JSON.parse(CacheService.getDocumentCache().get(ar_claimsKey_(sheet)) || '{}') || {}; } catch (e) { c = {}; }
  const now = Date.now();
  const ttl = AR_CONFIG.holdMinutes * 60000;
  Object.keys(c).forEach(function (r) { if (!c[r] || now - c[r].t > ttl) delete c[r]; });
  return c;
}

function ar_others_(claims, sid) {
  const seen = {};
  Object.keys(claims).forEach(function (r) { if (claims[r].s !== sid) seen[claims[r].s] = true; });
  return Object.keys(seen).length;
}

/**
 * Decides which row the sidebar shows and holds it (plus the next) for this sidebar.
 * how: { row } forces a row; otherwise { showing } is kept if still open and not held by someone
 * else, else the next open row after { from }, else (with { edge }) the first open row from the end.
 */
function ar_place_(sheet, cols, sid, dir, how) {
  return ar_withLock_(function () {
    const claims = ar_claims_(sheet);
    const blocked = {};
    Object.keys(claims).forEach(function (r) { if (claims[r].s !== sid) blocked[r] = true; });
    const scan = ar_scan_(sheet, cols);
    const free = function (r) { return r >= 2 && r <= scan.lastRow && scan.isOpen[r - 2] && !blocked[r]; };
    let row = 0;
    let changed = false;
    if (how.row) {
      row = how.row;
    } else if (how.showing && free(how.showing)) {
      row = how.showing;
    } else {
      if (how.from) row = ar_nextFrom_(scan, how.from, dir, blocked);
      if (!row && how.edge) row = ar_edge_(scan, dir, blocked);
      changed = !!how.showing && row !== how.showing;
    }
    const peek = row ? ar_nextFrom_(scan, row, dir, blocked) : 0;
    if (sid) {
      Object.keys(claims).forEach(function (r) { if (claims[r].s === sid) delete claims[r]; });
      const now = Date.now();
      if (row) claims[row] = { s: sid, t: now };
      if (peek) claims[peek] = { s: sid, t: now };
      CacheService.getDocumentCache().put(ar_claimsKey_(sheet), JSON.stringify(claims), 21600);
    }
    return { scan: scan, row: row, peek: peek, changed: changed, others: ar_others_(claims, sid) };
  });
}

function ar_land_(sheet, cols, p) {
  if (p.row) ar_select_(sheet, cols, p.row);
  return {
    remaining: p.scan.remaining,
    current: p.row ? ar_rowInfo_(sheet, cols, p.row) : null,
    peek: p.peek ? ar_rowInfo_(sheet, cols, p.peek) : null,
    changed: p.changed,
    others: p.others,
    stats: ar_stats_(sheet),
  };
}

/* ---------------- Today's numbers ---------------- */

function ar_today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Toronto', 'yyyy-MM-dd');
}

function ar_statsKey_(sheet, day) {
  return 'ar_stats_' + sheet.getSheetId() + '_' + (day || ar_today_());
}

function ar_stats_(sheet, day) {
  try {
    const s = JSON.parse(PropertiesService.getDocumentProperties().getProperty(ar_statsKey_(sheet, day)) || 'null');
    if (s && typeof s.total === 'number') return s;
  } catch (e) {
    // fall through
  }
  return { total: 0, by: {} };
}

/** Adds (or with -1 removes) one rating to a day's numbers. Call inside ar_withLock_. */
function ar_bump_(sheet, label, delta, day) {
  const s = ar_stats_(sheet, day);
  s.total = Math.max(0, s.total + delta);
  s.by[label] = Math.max(0, (s.by[label] || 0) + delta);
  if (!s.by[label]) delete s.by[label];
  PropertiesService.getDocumentProperties().setProperty(ar_statsKey_(sheet, day), JSON.stringify(s));
  return s;
}

/* ---------------- "go to …" ---------------- */

function ar_fold_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

/**
 * "row 40" is that row; a number is the audition # if the sheet has that column, otherwise the row;
 * anything else matches names / aliases (all words, any order), next match after fromRow first.
 */
function ar_find_(sheet, nameCols, numberCol, query, fromRow, dir) {
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
  const words = ar_fold_(q).split(/[^a-z0-9']+/).filter(Boolean);
  if (!words.length || lastRow < 2) throw new Error('Say who to go to, like “go to 722” or “go to Maria”.');
  const cols = nameCols.filter(Boolean);
  const vals = cols.map(function (c) { return sheet.getRange(2, c, lastRow - 1, 1).getDisplayValues(); });
  const matches = [];
  for (let i = 0; i < lastRow - 1; i++) {
    const hay = ar_fold_(vals.map(function (v) { return v[i][0]; }).join(' '));
    if (words.every(function (w) { return hay.indexOf(w) !== -1; })) matches.push(i + 2);
  }
  if (!matches.length) throw new Error('Couldn’t find anyone called “' + raw + '”.');
  const order = dir === 'up' ? matches.slice().reverse() : matches;
  let row = order.find(function (r) { return dir === 'up' ? r < fromRow : r > fromRow; });
  if (!row) row = order[0];
  return { row: row, how: 'name', count: matches.length, index: order.indexOf(row) + 1 };
}

/* ---------------- Helpers ---------------- */

/** What a command writes: { listened, rating, roleRating?, color }. */
function ar_spec_(cmd, value) {
  const c = AR_CONFIG.commands[cmd];
  let score = null;
  if (cmd === 'score') {
    score = Number(value);
  } else if (c && typeof c.score === 'number') {
    score = c.score;
  } else if (c) {
    return { listened: c.listened || AR_CONFIG.listenedValue, rating: c.rating, roleRating: c.roleRating, color: c.color || null };
  }
  if (score === null || isNaN(score) || score < 0 || score > 5) {
    throw new Error('Unknown rating: ' + cmd + (value !== undefined && value !== null ? ' ' + value : ''));
  }
  score = Math.round(score * 10) / 10;
  return { listened: AR_CONFIG.listenedValue, rating: score, color: ar_colorFor_(score) };
}

function ar_colorFor_(score) {
  for (let i = 0; i < AR_CONFIG.colorBands.length; i++) {
    if (score >= AR_CONFIG.colorBands[i].min) return AR_CONFIG.colorBands[i].color;
  }
  return null;
}

/** Every highlight colour this tool uses (lowercase). */
function ar_ourColors_() {
  const out = AR_CONFIG.colorBands.map(function (b) { return b.color.toLowerCase(); });
  Object.keys(AR_CONFIG.commands).forEach(function (k) {
    if (AR_CONFIG.commands[k].color) out.push(AR_CONFIG.commands[k].color.toLowerCase());
  });
  return out;
}

function ar_sheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found. Reopen the reviewer.');
  return sheet;
}

function ar_checkRow_(sheet, row) {
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Row ' + row + ' is out of range.');
}

function ar_columns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
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

  const cols = { lastCol: lastCol };
  Object.keys(AR_CONFIG.headers).forEach(function (key) { cols[key] = find(AR_CONFIG.headers[key]); });

  const missing = ['listened', 'rating', 'notes'].filter(function (k) { return !cols[k]; });
  if (missing.length) {
    throw new Error('Couldn\'t find the ' + missing.join(', ') + ' column(s) on "' + sheet.getName() +
      '". Check the header names in AR_CONFIG.headers.');
  }
  if (!cols.upload && !cols.link) throw new Error('Couldn\'t find an audition upload or link column.');
  return cols;
}

/** One pass over the sheet: which rows still need listening. */
function ar_scan_(sheet, cols) {
  const lastRow = sheet.getLastRow();
  const n = Math.max(0, lastRow - 1);
  const isOpen = new Array(n);
  let remaining = 0;
  if (!n) return { isOpen: isOpen, remaining: 0, lastRow: lastRow };

  const listened = sheet.getRange(2, cols.listened, n, 1).getDisplayValues();
  const names = cols.name ? sheet.getRange(2, cols.name, n, 1).getDisplayValues() : null;
  for (let i = 0; i < n; i++) {
    const notListened = String(listened[i][0]).trim() === '';
    const hasPerson = !names || String(names[i][0]).trim() !== '';
    isOpen[i] = notListened && hasPerson;
    if (isOpen[i]) remaining++;
  }
  return { isOpen: isOpen, remaining: remaining, lastRow: lastRow };
}

/** Next open row after fromRow (exclusive) going 'down' or 'up', skipping blocked rows. 0 if none. */
function ar_nextFrom_(scan, fromRow, dir, blocked) {
  const step = dir === 'up' ? -1 : 1;
  for (let r = fromRow + step; r >= 2 && r <= scan.lastRow; r += step) {
    if (scan.isOpen[r - 2] && !(blocked && blocked[r])) return r;
  }
  return 0;
}

/** First open row from the top (dir 'down') or from the bottom (dir 'up'). */
function ar_edge_(scan, dir, blocked) {
  return dir === 'up' ? ar_nextFrom_(scan, scan.lastRow + 1, 'up', blocked) : ar_nextFrom_(scan, 1, 'down', blocked);
}

function ar_rowInfo_(sheet, cols, row) {
  const vals = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
  const get = function (c) { return c ? String(vals[c - 1] || '').trim() : ''; };
  return {
    row: row,
    name: get(cols.name) || '(no name)',
    alias: get(cols.alias),
    number: get(cols.number),
    listened: get(cols.listened),
    rating: get(cols.rating),
    notes: get(cols.notes),
    media: ar_media_(get(cols.upload), get(cols.link)),
  };
}

function ar_nameAt_(sheet, cols, row) {
  return cols.name ? String(sheet.getRange(row, cols.name).getDisplayValue()).trim() : 'Row ' + row;
}

function ar_select_(sheet, cols, row) {
  sheet.getRange(row, cols.name || 1).activate();
}

/** Turns the upload + link cells into a list of playable things. */
function ar_media_(uploadVal, linkVal) {
  const out = [];
  const seen = {};
  [uploadVal, linkVal].forEach(function (v) {
    String(v || '').split(/[\s,]+/).forEach(function (u) {
      if (!/^https?:\/\//i.test(u) || seen[u]) return;
      seen[u] = true;
      const id = ar_driveId_(u);
      out.push(id ? { kind: 'drive', id: id, url: u, folder: /\/folders\//i.test(u) } : { kind: 'link', url: u });
    });
  });
  return out;
}

function ar_driveId_(url) {
  if (!/(drive|docs)\.google\.com/i.test(url)) return null;
  const m = url.match(/[?&]id=([\w-]{20,})/) || url.match(/\/d\/([\w-]{20,})/) || url.match(/\/folders\/([\w-]{20,})/);
  return m ? m[1] : null;
}

function ar_dirKey_() {
  return 'ar_dir_' + SpreadsheetApp.getActiveSpreadsheet().getId();
}

function ar_getDir_() {
  return PropertiesService.getUserProperties().getProperty(ar_dirKey_()) === 'up' ? 'up' : 'down';
}
