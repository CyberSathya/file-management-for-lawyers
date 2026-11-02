// File Management System for Lawyers: a web app over a "one tab per rack" file-register spreadsheet.
// Searches every tab (Rack 01 ... Unsorted), and adds new files to a chosen rack.
//
// Requires the Google Sheets advanced service: Services (+) -> Google Sheets API -> Add.
// Run setupRackStatus() once from the editor to put the status dropdown in A1 of every rack.
// See README.md for setup.

// Optional: a prefix some file numbers carry that should be ignored when matching,
// e.g. 'ABC' makes ABC-14-011 the same file as 14-00011. Leave '' for none.
const IGNORED_PREFIX = '';

// The spreadsheet: the one this script is bound to (Extensions -> Apps Script from the
// sheet), or the Script Property SPREADSHEET_ID for a standalone script.
function sheetId_() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') ||
    SpreadsheetApp.getActiveSpreadsheet().getId();
}
function book_() { return SpreadsheetApp.openById(sheetId_()); }

const MAX_RESULTS = 200;
const DEFAULT_HEADERS = ['File No.', 'Place', 'Court', 'Case No.', 'Cause Title',
                         'Client Name', 'Case Status', 'Date'];

// Each rack's status is a dropdown in A1, replacing the unused "File No." header.
const STATUS_CELL = 'A1';
const OLD_STATUS_CELLS = ['Y1', 'Z1'];   // earlier location, cleared by setupRackStatus()
const STATUS_OPTIONS = ['Space Available', 'Almost Full', 'Full'];
const STATUS_COLOURS = { 'Space Available': '#b7e1cd', 'Almost Full': '#fce8b2', 'Full': '#f4c7c3' };
const TEMPLATE_RACK = 'Rack 01';         // copyStatusDropdown() copies this tab's A1 dropdown
const SEARCH_COLS = 24;   // A..X
const DATA_COLS = 8;      // A..H: File No. .. Date
const RECENT_KEY = 'recentRacks';

// A hidden tab recording who added / edited / took out / returned which file. It is
// not a rack, so every rack feature skips it.
const LOG_SHEET = 'Activity Log';
const LOG_HEADERS = ['When', 'By', 'Action', 'File No.', 'Rack', 'Taken by', 'Details', 'Note'];

function isSystemSheet_(name) { return name === LOG_SHEET; }
function rackSheets_(ss) {
  return ss.getSheets().filter(function (sh) { return !isSystemSheet_(sh.getName()); });
}

// File-number matching (same rule as the page): 08-00009 = 08-009 = 08-09, ignoring case,
// spaces and IGNORED_PREFIX.
function fileKey_(s) {
  let t = String(s || '').toUpperCase().replace(/\s+/g, '');
  if (IGNORED_PREFIX) t = t.replace(new RegExp('^' + IGNORED_PREFIX.toUpperCase() + '-?(?=\\d)'), '');
  const m = t.match(/^(.*?)(\d+)$/);
  return m ? m[1] + parseInt(m[2], 10) : '';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Law File Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Search ----------

// Called from the page. Returns { results: [...], truncated: bool, statuses: { rack: status } }.
function search(query) {
  query = String(query || '').trim();
  if (query.length < 2) return { results: [], truncated: false, statuses: {} };

  const ss = book_();

  // TextFinder searches all tabs at once, far faster than reading every sheet.
  const hits = ss.createTextFinder(query).matchCase(false).findAll();

  // Group hit row numbers by sheet, skipping header rows and duplicate rows.
  const bySheet = {};
  hits.forEach(function (r) {
    if (r.getRow() === 1 || r.getColumn() > SEARCH_COLS) return;
    const name = r.getSheet().getName();
    if (isSystemSheet_(name)) return;
    (bySheet[name] = bySheet[name] || new Set()).add(r.getRow());
  });

  const results = [];
  const statuses = {};
  const out = filesOut_(ss);
  let truncated = false;

  // Walk sheets in tab order so results come back rack by rack.
  for (const sheet of rackSheets_(ss)) {
    const rows = bySheet[sheet.getName()];
    if (!rows) continue;

    const lastCol = Math.min(Math.max(sheet.getLastColumn(), DEFAULT_HEADERS.length), SEARCH_COLS);
    const values = sheet.getDataRange().getDisplayValues();
    statuses[sheet.getName()] = normaliseStatus_((values[0] || [])[0]);   // A1 holds the rack status
    const headers = (values[0] || []).slice(0, lastCol).map(function (h, i) {
      // Column A header is junk on some tabs; always call it File No.
      if (i === 0) return 'File No.';
      return String(h).trim() || DEFAULT_HEADERS[i] || ('Col ' + (i + 1));
    });

    for (const rowNum of Array.from(rows).sort(function (a, b) { return a - b; })) {
      if (results.length >= MAX_RESULTS) { truncated = true; break; }
      const row = values[rowNum - 1] || [];
      results.push({
        sheet: sheet.getName(),
        row: rowNum,
        fileNo: row[0] || '',
        headers: headers,
        cells: headers.map(function (_, i) { return row[i] || ''; }),
        out: out[fileKey_(row[0])] || null   // { person, since, note } while taken out
      });
    }
    if (truncated) break;
  }

  return { results: results, truncated: truncated, statuses: statuses };
}

// ---------- "Did you mean" suggestions ----------

const VOCAB_CACHE_KEY = 'vocab';
const VOCAB_CACHE_SECS = 600;   // rebuild the word list at most every 10 minutes
const MAX_SUGGESTIONS = 8;

// Returns [{ word, query, count }] — words in the sheet spelled or sounding like a
// word of the query, with the full query to run if that word is swapped in.
function suggest(query) {
  query = String(query || '').trim();
  const words = query.toLowerCase().match(/[a-z]+/g) || [];
  if (!words.length) return [];

  const vocab = getVocab_();
  const out = [];
  const seen = {};

  words.forEach(function (w) {
    if (w.length < 3) return;
    const key = phoneticKey_(w);
    const maxDist = w.length <= 4 ? 1 : 2;
    const candidates = [];
    Object.keys(vocab).forEach(function (v) {
      if (v === w || Math.abs(v.length - w.length) > maxDist) return;
      const sounds = phoneticKey_(v) === key;
      const dist = editDistance_(w, v, maxDist);
      if (!sounds && dist > maxDist) return;
      candidates.push({ word: v, sounds: sounds, dist: dist, count: vocab[v] });
    });
    // Sound-alikes first, then closest spelling, then most common in the sheet.
    candidates.sort(function (a, b) {
      return (b.sounds - a.sounds) || (a.dist - b.dist) || (b.count - a.count);
    });
    candidates.slice(0, MAX_SUGGESTIONS).forEach(function (c) {
      const q = query.replace(new RegExp('\\b' + w + '\\b', 'i'), c.word);
      if (seen[q]) return;
      seen[q] = true;
      out.push({ word: c.word, query: q, count: c.count });
    });
  });
  return out.slice(0, MAX_SUGGESTIONS);
}

// word -> number of times it appears, across columns B..X of every rack.
function getVocab_() {
  const cache = CacheService.getScriptCache();
  const parts = Number(cache.get(VOCAB_CACHE_KEY + '_n') || 0);
  if (parts) {
    const keys = [];
    for (let i = 0; i < parts; i++) keys.push(VOCAB_CACHE_KEY + '_' + i);
    const got = cache.getAll(keys);
    if (keys.every(function (k) { return got[k] != null; })) {
      return JSON.parse(keys.map(function (k) { return got[k]; }).join(''));
    }
  }

  const ss = book_();
  const ranges = rackSheets_(ss).map(function (sh) {
    return "'" + sh.getName().replace(/'/g, "''") + "'!B2:X";
  });
  const resp = Sheets.Spreadsheets.Values.batchGet(sheetId_(), { ranges: ranges });
  const vocab = {};
  resp.valueRanges.forEach(function (vr) {
    (vr.values || []).forEach(function (row) {
      row.forEach(function (cell) {
        (String(cell).toLowerCase().match(/[a-z]{3,}/g) || []).forEach(function (w) {
          vocab[w] = (vocab[w] || 0) + 1;
        });
      });
    });
  });

  // Cache entries are capped at 100 KB, so store the list in chunks.
  const json = JSON.stringify(vocab);
  const chunk = 90000, put = {};
  let n = 0;
  for (let i = 0; i < json.length; i += chunk) put[VOCAB_CACHE_KEY + '_' + n++] = json.slice(i, i + chunk);
  put[VOCAB_CACHE_KEY + '_n'] = String(n);
  cache.putAll(put, VOCAB_CACHE_SECS);
  return vocab;
}

// Rough "sounds like" key tuned for Indian name spellings:
// sowmya / soumya / sowmiya -> sumia, prabhu / prabu -> prabu, deepa / dheepa -> dipa.
function phoneticKey_(w) {
  return w.toLowerCase()
    .replace(/[^a-z]/g, '')
    .replace(/ph/g, 'f')
    .replace(/([bdgkt])h/g, '$1')      // bh dh gh kh th -> b d g k t
    .replace(/sh/g, 's').replace(/z/g, 's')
    .replace(/ck|q/g, 'k')
    .replace(/w/g, 'v')
    .replace(/ov|ou|oo/g, 'u')         // sowmya, soumya -> sumya
    .replace(/ee/g, 'i')
    .replace(/y/g, 'i')
    .replace(/(.)\1+/g, '$1');         // collapse doubles: aa, ii, nn, ...
}

// Edit distance with swapped-letter support; stops early once it exceeds max.
function editDistance_(a, b, max) {
  const m = a.length, n = b.length;
  let prev2 = null, prev = [], cur;
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    cur = [i];
    let rowMin = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      let d = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      if (prev2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d = Math.min(d, prev2[j - 2] + 1);
      }
      cur[j] = d;
      if (d < rowMin) rowMin = d;
    }
    if (rowMin > max) return max + 1;
    prev2 = prev;
    prev = cur;
  }
  return prev[n];
}

// ---------- Racks ----------

// Every rack with its status, next free file number and when it was last used.
function getRacks() {
  const ss = book_();
  const sheets = rackSheets_(ss);
  const ranges = [];
  sheets.forEach(function (sh) {
    const q = "'" + sh.getName().replace(/'/g, "''") + "'";
    ranges.push(q + '!A2:H', q + '!' + STATUS_CELL);
  });

  // One API call for all 70-odd tabs instead of two calls per tab.
  const resp = Sheets.Spreadsheets.Values.batchGet(sheetId_(), { ranges: ranges });
  const recent = getRecent_();

  return sheets.map(function (sh, i) {
    const rows = resp.valueRanges[i * 2].values || [];
    const statusRaw = ((resp.valueRanges[i * 2 + 1].values || [[]])[0] || [])[0] || '';
    const slot = nextSlot_(rows, sh.getName());
    return {
      name: sh.getName(),
      status: normaliseStatus_(statusRaw),
      used: slot.used,
      nextFileNo: slot.fileNo,
      lastUsedAt: recent[sh.getName()] || 0
    };
  });
}

// Files each rack already holds for a client / party name.
// Returns { counts: { rack: n }, files: { rack: [{ row, cells: [A..H] }] } }.
function racksForName(name) {
  name = String(name || '').trim();
  if (name.length < 3) return { counts: {}, files: {} };
  const ss = book_();

  const rowsBySheet = {};
  ss.createTextFinder(name).matchCase(false).findAll().forEach(function (r) {
    const col = r.getColumn();
    if (r.getRow() === 1 || (col !== 5 && col !== 6)) return;   // Cause Title, Client Name
    const s = r.getSheet().getName();
    if (isSystemSheet_(s)) return;
    (rowsBySheet[s] = rowsBySheet[s] || new Set()).add(r.getRow());
  });

  // Read each matching rack once, so the popup can list the files without another trip.
  const counts = {}, files = {};
  Object.keys(rowsBySheet).forEach(function (s) {
    const rows = Array.from(rowsBySheet[s]).sort(function (a, b) { return a - b; });
    const values = ss.getSheetByName(s).getRange(1, 1, rows[rows.length - 1], DATA_COLS).getDisplayValues();
    counts[s] = rows.length;
    files[s] = rows.map(function (row) { return { row: row, cells: values[row - 1] }; });
  });
  return { counts: counts, files: files };
}

// Saves a new file into the next slot of a rack. fields = [Place, Court, Case No.,
// Cause Title, Client Name, Case Status, Date]. Returns the assigned file number.
function addFile(rackName, fields) {
  if (!fields || !fields.some(function (v) { return String(v || '').trim(); })) {
    throw new Error('Fill in at least one field.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);   // two people saving at once must not get the same slot
  try {
    const ss = book_();
    const sheet = ss.getSheetByName(rackName);
    if (!sheet) throw new Error('Rack "' + rackName + '" not found.');
    if (normaliseStatus_(sheet.getRange(STATUS_CELL).getDisplayValue()) === 'full') {
      throw new Error('Rack "' + rackName + '" is marked Full.');
    }

    const lastRow = Math.max(sheet.getLastRow(), 2);
    const rows = sheet.getRange(2, 1, lastRow - 1, DATA_COLS).getDisplayValues();
    const slot = nextSlot_(rows, rackName);
    // A note row ("IKEA BOX RACK 53") sits where the next file goes: push it down.
    if (slot.insert) sheet.insertRowBefore(slot.index + 2);

    const values = [slot.fileNo].concat(fields.slice(0, DATA_COLS - 1).map(function (v) {
      return String(v || '').trim();
    }));
    while (values.length < DATA_COLS) values.push('');

    const target = sheet.getRange(slot.index + 2, 1, 1, DATA_COLS);
    target.setNumberFormat('@');   // keep "1/2023" style case numbers from turning into dates
    target.setValues([values]);

    logActivity_(ss, 'Added', slot.fileNo, rackName, '', [values[4], values[5], values[3]].filter(String).join(' · '));

    const recent = getRecent_();
    recent[rackName] = Date.now();
    PropertiesService.getScriptProperties().setProperty(RECENT_KEY, JSON.stringify(recent));
    CacheService.getScriptCache().remove(VOCAB_CACHE_KEY + '_n');   // new words show up in suggestions

    return { fileNo: slot.fileNo, rack: rackName, row: slot.index + 2 };
  } finally {
    lock.releaseLock();
  }
}

const NEW_RACK_DIGITS = 3;   // empty racks start at e.g. 23-001

// Updates Place .. Date (B..H) of an existing file. expectedFileNo guards against the
// row having moved (e.g. a row inserted) since the page loaded it.
function updateFile(rackName, row, expectedFileNo, fields) {
  row = Number(row);
  if (!(row >= 2)) throw new Error('Invalid row.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    const sheet = ss.getSheetByName(rackName);
    if (!sheet || isSystemSheet_(rackName)) throw new Error('Rack "' + rackName + '" not found.');
    const current = sheet.getRange(row, 1).getDisplayValue().trim();
    if (current !== String(expectedFileNo || '').trim()) {
      throw new Error('This file has moved in the sheet since you searched. Search again and retry.');
    }
    const values = fields.slice(0, DATA_COLS - 1).map(function (v) { return String(v || '').trim(); });
    while (values.length < DATA_COLS - 1) values.push('');
    const target = sheet.getRange(row, 2, 1, DATA_COLS - 1);
    const before = target.getDisplayValues()[0];
    target.setNumberFormat('@');
    target.setValues([values]);
    CacheService.getScriptCache().remove(VOCAB_CACHE_KEY + '_n');

    // Log what changed, e.g. "Case Status: Pending → Disposed".
    const changes = values.map(function (v, i) {
      return v === String(before[i] || '').trim() ? '' :
        DEFAULT_HEADERS[i + 1] + ': ' + (String(before[i] || '').trim() || '(blank)') + ' → ' + (v || '(blank)');
    }).filter(String);
    if (changes.length) logActivity_(ss, 'Edited', current, rackName, '', changes.join('; '));
    return sheet.getRange(row, 1, 1, DATA_COLS).getDisplayValues()[0];
  } finally {
    lock.releaseLock();
  }
}

// Existing files with the same Case No. or Cause Title, ignoring case, spaces and
// punctuation. Returns [{ sheet, row, cells: [A..H], why }].
function findDuplicates(caseNo, title) {
  const norm = function (s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  const c = norm(caseNo), t = norm(title);
  if (c.length < 4 && t.length < 6) return [];

  const ss = book_();
  const sheets = rackSheets_(ss);
  const resp = Sheets.Spreadsheets.Values.batchGet(sheetId_(), {
    ranges: sheets.map(function (sh) { return "'" + sh.getName().replace(/'/g, "''") + "'!A2:H"; })
  });
  const out = [];
  resp.valueRanges.forEach(function (vr, i) {
    (vr.values || []).forEach(function (r, j) {
      const why = [];
      if (c.length >= 4 && norm(r[3]) === c) why.push('same Case No.');
      if (t.length >= 6 && norm(r[4]) === t) why.push('same Cause Title');
      if (!why.length) return;
      const cells = [];
      for (let k = 0; k < DATA_COLS; k++) cells.push(r[k] || '');
      out.push({ sheet: sheets[i].getName(), row: j + 2, cells: cells, why: why.join(' & ') });
    });
  });
  return out.slice(0, 20);
}

// rows = A2:H values. Returns { index, fileNo, used, insert }: the 0-based row for the
// next file, its number, how many rows hold files, and whether a row must be inserted.
// The slot is the row after the last one with details; gaps in the middle are left
// alone, as they may be real files not yet typed in.
function nextSlot_(rows, sheetName) {
  const a = analyseRack_(rows);
  const parsed = a.parsed, isNote = a.isNote, isFile = a.isFile;

  let lastUsed = -1, used = 0;
  rows.forEach(function (r, i) {
    if (isNote[i]) return;
    if (hasDetails_(r)) { lastUsed = i; used++; }
  });

  let index = lastUsed + 1;
  if (lastUsed < 0 && isFile.indexOf(true) >= 0) index = isFile.indexOf(true);   // nothing used yet

  // Continue from the nearest good number above, e.g. 04-023 -> 04-024.
  let fileNo;
  let j = -1;
  for (let i = index - 1; i >= 0; i--) if (isFile[i]) { j = i; break; }
  if (j >= 0) {
    const p = parsed[j];
    fileNo = p.prefix + String(p.n + (index - j)).padStart(p.digits.length, '0');
  } else if (isFile[index]) {
    fileNo = parsed[index].raw;
  } else {
    fileNo = prefixFromName_(sheetName) + String(1).padStart(NEW_RACK_DIGITS, '0');
  }
  // Keep the sheet's own spelling when the pre-numbered row already says the same.
  if (isFile[index] && parsed[index].n === parseInt(fileNo.match(/(\d+)$/)[1], 10)) fileNo = parsed[index].raw;

  return { index: index, fileNo: fileNo, used: used, insert: !!isNote[index] };
}

function hasDetails_(r) {
  return r.slice(1, DATA_COLS).some(function (v) { return String(v || '').trim(); });
}

// Reads a rack's column A: parsed numbers, which rows are notes ("IKEA BOX RACK 53",
// a different prefix from the rack's own) and which are good file numbers to count on
// (a number going backwards, like "13-00" after 13-0051, is a typo).
function analyseRack_(rows) {
  const parsed = rows.map(function (r) {
    const a = String(r[0] || '').trim();
    if (/^IKEA\b|\bBOX\b/i.test(a)) return { raw: a, key: null };   // box label, even if it ends in digits
    const m = a.match(/^(.*?)(\d+)$/);
    if (m) return { raw: a, prefix: m[1], key: m[1].replace(/\s+/g, '').toUpperCase(), digits: m[2], n: parseInt(m[2], 10) };
    return a ? { raw: a, key: null } : null;
  });

  // The rack's numbering is its most common prefix ("53-", "D1-"). Column A text
  // with another prefix ("IKEA BOX RACK 53") is a note row, not a file.
  const freq = {};
  let key;
  parsed.forEach(function (p) {
    if (!p || p.key === null) return;
    freq[p.key] = (freq[p.key] || 0) + 1;
    if (key === undefined || freq[p.key] > freq[key]) key = p.key;
  });
  const isNote = parsed.map(function (p) { return !!p && (p.key === null || p.key !== key); });

  let max = -1;
  const isFile = parsed.map(function (p, i) {
    if (!p || isNote[i] || p.n < max) return false;
    max = p.n;
    return true;
  });
  return { parsed: parsed, key: key, isNote: isNote, isFile: isFile };
}

// ---------- Activity log ----------

function logSheet_(ss) {
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET, ss.getNumSheets());
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 2, sh.getMaxRows(), LOG_HEADERS.length - 1).setNumberFormat('@');   // "08-09" stays text
    sh.hideSheet();
  }
  return sh;
}

// action: Added | Edited | Taken out | Returned. details: free text (what changed, title).
function logActivity_(ss, action, fileNo, rack, person, details, note) {
  const by = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  logSheet_(ss).appendRow([new Date(), by, action, String(fileNo || ''), rack || '', person || '', details || '', note || '']);
}

// ---------- File check-out / return ----------

// fileKey -> { fileNo, rack, person, since (ISO), note, title, by } for files taken out
// and not yet returned, replayed from the activity log.
function filesOut_(ss) {
  const sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEADERS.length).getValues();
  const out = {};
  rows.forEach(function (r) {
    const k = fileKey_(r[3]);
    if (!k) return;
    if (r[2] === 'Taken out') {
      out[k] = { fileNo: String(r[3]), rack: String(r[4]), person: String(r[5]), title: String(r[6]),
                 note: String(r[7]), by: String(r[1]),
                 since: r[0] instanceof Date ? r[0].toISOString() : String(r[0]) };
    } else if (r[2] === 'Returned') {
      delete out[k];
    }
  });
  return out;
}

function checkOut(fileNo, rack, person, note, title) {
  person = String(person || '').trim();
  if (!fileKey_(fileNo)) throw new Error('This file has no file number.');
  if (!person) throw new Error('Enter who is taking the file.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    const already = filesOut_(ss)[fileKey_(fileNo)];
    if (already) throw new Error(fileNo + ' is already out with ' + already.person + '.');
    logActivity_(ss, 'Taken out', fileNo, rack, person, title, note);
    return filesOut_(ss)[fileKey_(fileNo)];
  } finally {
    lock.releaseLock();
  }
}

function checkIn(fileNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    const o = filesOut_(ss)[fileKey_(fileNo)];
    if (!o) throw new Error(fileNo + ' is not marked as taken out.');
    logActivity_(ss, 'Returned', o.fileNo, o.rack, o.person, o.title, '');
    return null;
  } finally {
    lock.releaseLock();
  }
}

// The "Files out" screen: every file out, longest first, plus names used before.
function getFilesOut() {
  const ss = book_();
  const out = filesOut_(ss);
  const now = Date.now();
  const files = Object.keys(out).map(function (k) {
    out[k].days = Math.floor((now - new Date(out[k].since).getTime()) / 86400000);
    return out[k];
  }).sort(function (a, b) { return a.since < b.since ? -1 : 1; });
  return { files: files, people: recentPeople_(ss) };
}

function recentPeople_(ss) {
  const sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(sh.getLastRow() - 1, 500);
  const rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, LOG_HEADERS.length).getValues();
  const seen = [];
  for (let i = rows.length - 1; i >= 0 && seen.length < 20; i--) {
    const p = String(rows[i][5] || '').trim();
    if (rows[i][2] === 'Taken out' && p && seen.indexOf(p) < 0) seen.push(p);
  }
  return seen;
}

// ---------- Clean-up report ----------

// Problems in the rack sheet worth fixing by hand. Each item: { rack, fileNo, text }.
function cleanupReport() {
  const ss = book_();
  const sheets = rackSheets_(ss);
  const ranges = [];
  sheets.forEach(function (sh) {
    const q = "'" + sh.getName().replace(/'/g, "''") + "'";
    ranges.push(q + '!A2:H', q + '!' + STATUS_CELL);
  });
  const resp = Sheets.Spreadsheets.Values.batchGet(sheetId_(), { ranges: ranges });

  const missing = [], typos = [], styles = [], unset = [];
  const seen = {};                       // fileKey -> [{ fileNo, rack }]
  const spell = { place: {}, court: {} };  // normalised -> { raw spelling: count }
  const summary = function (r) { return [r[4], r[5], r[3]].map(function (v) { return String(v || '').trim(); }).filter(String).join(' · '); };

  sheets.forEach(function (sh, i) {
    const name = sh.getName();
    const rows = resp.valueRanges[i * 2].values || [];
    const status = normaliseStatus_(((resp.valueRanges[i * 2 + 1].values || [[]])[0] || [])[0]);
    if (status === 'unset') unset.push({ rack: name, fileNo: '', text: 'Rack status not set (A1)' });

    const a = analyseRack_(rows);
    const widths = {};
    let lastGood = '';
    rows.forEach(function (r, j) {
      const p = a.parsed[j];
      const has = hasDetails_(r);
      if (!p && has) {
        missing.push({ rack: name, fileNo: '', text: 'A file with no number' + (lastGood ? ', just after ' + lastGood : ', at the top') + ': ' + summary(r) });
      } else if (p && !a.isNote[j] && !a.isFile[j]) {
        const s = summary(r);
        typos.push({ rack: name, fileNo: p.raw, text: p.raw + ' looks mistyped (it comes after ' + lastGood + ')' + (s ? ': ' + s : '') });
      }
      if (a.isFile[j]) {
        lastGood = p.raw;
        widths[p.digits.length] = (widths[p.digits.length] || []).concat(p.raw);
      }
      if (p && !a.isNote[j]) (seen[fileKey_(p.raw)] = seen[fileKey_(p.raw)] || []).push({ fileNo: p.raw, rack: name, status: String(r[6] || ''),
        caseId: String(r[3] || r[4] || '').toLowerCase().replace(/[^a-z0-9]/g, '') });
      [['place', 1], ['court', 2]].forEach(function (f) {
        const raw = String(r[f[1]] || '');
        const n = raw.toLowerCase().replace(/[.,]+$/, '').replace(/\s+/g, ' ').trim();
        if (!n) return;
        const m = spell[f[0]][n] = spell[f[0]][n] || {};
        m[raw] = (m[raw] || 0) + 1;
      });
    });
    const w = Object.keys(widths);
    if (w.length > 1) {
      styles.push({ rack: name, fileNo: '', text: 'Mixes number styles: ' + w.map(function (d) {
        return widths[d][0] + ' (' + widths[d].length + (widths[d].length === 1 ? ' file)' : ' files)');
      }).join(', ') });
    }
  });

  // One case kept in several bundles shares one File No., with "Bundle 1 of 2" etc. in
  // Case Status. That's intended, so only flag repeats not all marked as bundles — or
  // whose rows name different cases.
  const isBundles = function (list) {
    const cases = new Set(list.map(function (s) { return s.caseId; }).filter(String));
    return cases.size <= 1 && list.every(function (s) { return /\bbundle\b/i.test(s.status); });
  };
  const dups = Object.keys(seen).filter(function (k) { return seen[k].length > 1 && !isBundles(seen[k]); }).map(function (k) {
    const racks = Array.from(new Set(seen[k].map(function (s) { return s.rack; })));
    return { rack: racks[0], fileNo: seen[k][0].fileNo, text: seen[k][0].fileNo + ' is listed ' + seen[k].length + ' times' +
      (racks.length > 1 ? ' (in ' + racks.join(', ') + ')' : '') };
  });

  // The same word written different ways: "Springfield (120), SPRINGFIELD (30), springfield  (8)".
  const variants = function (map) {
    return Object.keys(map).filter(function (n) { return Object.keys(map[n]).length > 1; })
      .map(function (n) {
        const forms = Object.keys(map[n]).sort(function (x, y) { return map[n][y] - map[n][x]; });
        const total = forms.reduce(function (s, f) { return s + map[n][f]; }, 0);
        return { total: total, item: { rack: '', fileNo: '', text: forms.map(function (f) {
          return '"' + f + '" (' + map[n][f] + ')'; }).join(', ') } };
      })
      .sort(function (x, y) { return y.total - x.total; }).slice(0, 60).map(function (v) { return v.item; });
  };

  return {
    generatedAt: new Date().toISOString(),
    sections: [
      { key: 'dups', title: 'File numbers listed more than once', hint: 'Two files may share a number: give one a new number. If they are bundles of one case, write "Bundle 1 of 2", "Bundle 2 of 2" in Case Status.', items: dups },
      { key: 'missing', title: 'Files with no file number', hint: 'Type the missing number into column A.', items: missing },
      { key: 'typos', title: 'File numbers that look mistyped', hint: 'They break the numbering order of the rack.', items: typos },
      { key: 'styles', title: 'Racks mixing number styles', hint: 'Not wrong, but one style per rack is easier to read.', items: styles },
      { key: 'places', title: 'Places spelled different ways', hint: 'Pick one spelling so filters group them.', items: variants(spell.place) },
      { key: 'courts', title: 'Courts spelled different ways', hint: 'Pick one spelling so filters group them.', items: variants(spell.court) },
      { key: 'unset', title: 'Racks with no status', hint: 'Set Space Available / Almost Full / Full in A1.', items: unset }
    ]
  };
}

// "Rack 23" -> "23-", "57 corner rack" -> "57-", "Rack D2" -> "D2-".
function prefixFromName_(name) {
  const s = String(name || '').replace(/^rack\s*/i, '').trim();
  const m = s.match(/^\d+/);
  if (m) return m[0].padStart(2, '0') + '-';
  return (s.split(/\s+/)[0] || 'NEW').toUpperCase() + '-';
}

function normaliseStatus_(s) {
  s = String(s || '').toLowerCase();
  if (/almost/.test(s)) return 'almost';
  if (/full/.test(s)) return 'full';
  if (/space|available|empty/.test(s)) return 'space';
  return 'unset';
}

function getRecent_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(RECENT_KEY) || '{}');
  } catch (e) {
    return {};
  }
}

// ---------- One-time setup (run from the Apps Script editor) ----------

// Puts a status dropdown in A1 of every tab (green / yellow / red), carrying over any
// status already chosen in Z1 and clearing the old Y1/Z1 cells. Safe to run again.
function setupRackStatus() {
  const ss = book_();
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build();
  const skipped = [];

  rackSheets_(ss).forEach(function (sh) {
    const cell = sh.getRange(STATUS_CELL);
    const current = cell.getDisplayValue().trim();
    // A1 should only hold a header ("File No.", "`````") or a status. A file number
    // there means the tab has no header row, so leave it alone.
    if (/^\S*\d+-\d+/.test(current)) {
      skipped.push(sh.getName());
      return;
    }

    const oldStatus = sh.getRange('Z1').getDisplayValue().trim();
    let status = STATUS_OPTIONS.indexOf(current) >= 0 ? current : '';
    if (!status && STATUS_OPTIONS.indexOf(oldStatus) >= 0) status = oldStatus;
    if (!status && /\bfull\b/i.test(sh.getName())) status = 'Full';

    cell.clearDataValidations().setValue(status).setDataValidation(rule)
      .setFontWeight('bold').setHorizontalAlignment('center');

    // Drop the old Y1/Z1 cells and any colour rules on them or on A1, then colour A1.
    OLD_STATUS_CELLS.forEach(function (a1) {
      sh.getRange(a1).clearContent().clearDataValidations().setFontWeight('normal');
    });
    const keep = sh.getConditionalFormatRules().filter(function (r) {
      return !r.getRanges().some(function (rg) {
        const a = rg.getA1Notation();
        return a === STATUS_CELL || OLD_STATUS_CELLS.indexOf(a) >= 0;
      });
    });
    Object.keys(STATUS_COLOURS).forEach(function (v) {
      keep.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(v).setBackground(STATUS_COLOURS[v]).setRanges([cell]).build());
    });
    sh.setConditionalFormatRules(keep);
  });

  Logger.log(skipped.length ? 'Skipped (A1 holds a file number): ' + skipped.join(', ') : 'All racks set up.');
}

// Optional: after you set the dropdown chip colours on Rack 01's A1 by hand
// (Data validation -> each item's colour), run this to copy them to every other tab.
function copyStatusDropdown() {
  const ss = book_();
  const source = ss.getSheetByName(TEMPLATE_RACK).getRange(STATUS_CELL);
  rackSheets_(ss).forEach(function (sh) {
    if (sh.getName() === TEMPLATE_RACK) return;
    const cell = sh.getRange(STATUS_CELL);
    if (cell.getDataValidation() === null) return;   // tabs skipped by setupRackStatus()
    source.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  });
  Logger.log('Dropdown copied from ' + TEMPLATE_RACK + '!A1 to all racks.');
}
