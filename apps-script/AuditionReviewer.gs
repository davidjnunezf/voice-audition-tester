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

/** Opens on the selected row if it hasn't been listened to yet, otherwise the next open one. */
function ar_init() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cols = ar_columns_(sheet);
  const dir = ar_getDir_();
  const scan = ar_scan_(sheet, cols);

  const range = sheet.getActiveRange();
  const active = range ? range.getRow() : 0;
  let row = 0;
  if (active >= 2 && active <= scan.lastRow && scan.isOpen[active - 2]) {
    row = active;
  } else if (active >= 2 && active <= scan.lastRow) {
    row = ar_nextFrom_(scan, active, dir);
  }
  if (!row) row = ar_edge_(scan, dir);

  if (row) ar_select_(sheet, cols, row);
  return {
    sheetName: sheet.getName(),
    title: ss.getName(),
    dir: dir,
    remaining: scan.remaining,
    current: row ? ar_rowInfo_(sheet, cols, row) : null,
  };
}

/** Rates the row, then returns the next open audition in the current direction. */
function ar_rate(sheetName, row, cmd, dir, value) {
  const spec = ar_spec_(cmd, value);
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, row);

  const rowRange = sheet.getRange(row, 1, 1, cols.lastCol);
  const listenedCell = sheet.getRange(row, cols.listened);
  const ratingCell = sheet.getRange(row, cols.rating);
  const roleCell = cols.roleRating ? sheet.getRange(row, cols.roleRating) : null;
  const backgrounds = rowRange.getBackgrounds()[0];

  const undo = {
    row: row,
    listened: listenedCell.getValue(),
    rating: ratingCell.getValue(),
    roleRating: roleCell ? roleCell.getValue() : null,
    backgrounds: backgrounds,
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

  const scan = ar_scan_(sheet, cols);
  const nextRow = ar_nextFrom_(scan, row, dir);
  if (nextRow) ar_select_(sheet, cols, nextRow);

  return {
    rated: { row: row, name: ar_nameAt_(sheet, cols, row), rating: String(spec.rating), cmd: cmd },
    undo: undo,
    remaining: scan.remaining,
    next: nextRow ? ar_rowInfo_(sheet, cols, nextRow) : null,
  };
}

/** Reverts a rating (using the undo object ar_rate returned) and goes back to that row. */
function ar_undo(sheetName, undo) {
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

  const scan = ar_scan_(sheet, cols);
  ar_select_(sheet, cols, undo.row);
  return { remaining: scan.remaining, current: ar_rowInfo_(sheet, cols, undo.row) };
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

/** Next open audition after row (doesn't change anything). Used to preload audio. */
function ar_peek(sheetName, row, dir) {
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const scan = ar_scan_(sheet, cols);
  const nextRow = ar_nextFrom_(scan, row, dir);
  return nextRow ? ar_rowInfo_(sheet, cols, nextRow) : null;
}

/** Jump to the first open audition from the top ('top') or bottom ('bottom'). */
function ar_jump(sheetName, where) {
  const dir = where === 'bottom' ? 'up' : 'down';
  ar_setDir(dir);
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const scan = ar_scan_(sheet, cols);
  const row = ar_edge_(scan, dir);
  if (row) ar_select_(sheet, cols, row);
  return { dir: dir, remaining: scan.remaining, current: row ? ar_rowInfo_(sheet, cols, row) : null };
}

/** Use whatever row is selected in the sheet right now (even if already reviewed). */
function ar_startHere(sheetName) {
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 0;
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Select a row with an audition first.');
  const scan = ar_scan_(sheet, cols);
  return { remaining: scan.remaining, current: ar_rowInfo_(sheet, cols, row) };
}

/** Opens a specific row without changing anything (used by Back). */
function ar_rowAt(sheetName, row) {
  const sheet = ar_sheet_(sheetName);
  const cols = ar_columns_(sheet);
  ar_checkRow_(sheet, row);
  const scan = ar_scan_(sheet, cols);
  ar_select_(sheet, cols, row);
  return { remaining: scan.remaining, current: ar_rowInfo_(sheet, cols, row) };
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

/** Next open row after fromRow (exclusive) going 'down' or 'up'. 0 if none. */
function ar_nextFrom_(scan, fromRow, dir) {
  const step = dir === 'up' ? -1 : 1;
  for (let r = fromRow + step; r >= 2 && r <= scan.lastRow; r += step) {
    if (scan.isOpen[r - 2]) return r;
  }
  return 0;
}

/** First open row from the top (dir 'down') or from the bottom (dir 'up'). */
function ar_edge_(scan, dir) {
  return dir === 'up' ? ar_nextFrom_(scan, scan.lastRow + 1, 'up') : ar_nextFrom_(scan, 1, 'down');
}

function ar_rowInfo_(sheet, cols, row) {
  const vals = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
  const get = function (c) { return c ? String(vals[c - 1] || '').trim() : ''; };
  return {
    row: row,
    name: get(cols.name) || '(no name)',
    alias: get(cols.alias),
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
