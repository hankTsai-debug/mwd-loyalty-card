/* Behavioral tests for mwd-loyalty-card Code.gs.
 * Mocks Google Apps Script services (SpreadsheetApp/LockService/Utilities/Logger)
 * with an in-memory spreadsheet that REPRODUCES Sheets' string->Date coercion,
 * then runs the real functions from Code.gs and asserts behavior.
 *
 * Run:  node mwd-loyalty-card/test/gs_test.js
 * (exits 0 if all pass, 1 if any fail) */
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const CODE = path.join(__dirname, '..', 'apps-script', 'Code.gs');

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✅ ' + name); }
  else { fail++; console.log('  ❌ ' + name + (extra ? '  -> ' + extra : '')); }
}
function eq(name, a, b) { ok(name + ' (=' + JSON.stringify(b) + ')', JSON.stringify(a) === JSON.stringify(b), 'got ' + JSON.stringify(a)); }

// ---- Mock Sheets ----
const pad = n => String(n).padStart(2, '0');
function fmtDate(date, tz, fmt) {
  const Y = date.getFullYear(), Mo = date.getMonth() + 1, D = date.getDate(),
        H = date.getHours(), Mi = date.getMinutes(), dow = ((date.getDay() + 6) % 7) + 1;
  return String(fmt)
    .replace('yyyy', Y).replace('HH', pad(H)).replace('mm', pad(Mi))
    .replace('MM', pad(Mo)).replace('dd', pad(D))
    .replace(/\bM\b/, Mo).replace(/\bd\b/, D).replace(/\bu\b/, dow);
}
class Range {
  constructor(sheet, r, c, nr, nc) { this.s = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getValue() { return this.s._get(this.r, this.c); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.s._get(this.r + i, this.c + j)); out.push(row); }
    return out;
  }
  setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.s._fmt(this.r + i, this.c + j, f); return this; }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.s._set(this.r + i, this.c + j, v); return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.fmt = []; }
  _grow(r) { while (this.rows.length < r) { this.rows.push([]); this.fmt.push([]); } }
  _get(r, c) { const v = (this.rows[r - 1] || [])[c - 1]; return v === undefined ? '' : v; }
  _fmt(r, c, f) { this._grow(r); this.fmt[r - 1][c - 1] = f; }
  _set(r, c, v) {
    this._grow(r);
    const f = (this.fmt[r - 1] || [])[c - 1];
    let store = v;
    // Reproduce Google Sheets: a plain "yyyy-MM-dd" string becomes a Date unless cell is text('@')
    if (f !== '@' && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
      const p = v.split('-').map(Number); store = new Date(p[0], p[1] - 1, p[2]);
    }
    this.rows[r - 1][c - 1] = store;
  }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  getDataRange() { return new Range(this, 1, 1, this.getLastRow() || 1, this.getLastColumn() || 1); }
  appendRow(arr) { this.rows.push(arr.slice()); this.fmt.push([]); }
  deleteRow(r) { this.rows.splice(r - 1, 1); this.fmt.splice(r - 1, 1); }
  setContents(arr) { this.rows = arr.map(r => r.slice()); this.fmt = arr.map(() => []); }
}
function newSS() {
  const sheets = {};
  return {
    _sheets: sheets,
    getSheetByName(n) { return sheets[n] || null; },
    insertSheet(n) { sheets[n] = new Sheet(n); return sheets[n]; },
    copy(name) { return { _name: name, getUrl() { return 'https://fake/' + encodeURIComponent(name); }, getId() { return 'fakeid-' + name; } }; },
  };
}
const SpreadsheetApp = { _ss: newSS(), getActiveSpreadsheet() { return this._ss; }, getUi() { throw new Error('no UI in test'); } };
const LockService = { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } };
const Utilities = { formatDate: fmtDate, computeDigest() { return new Array(16).fill(0); }, DigestAlgorithm: { MD5: 'MD5' } };
const Logger = { log() {} };

// ---- Load Code.gs into a sandbox ----
const src = fs.readFileSync(CODE, 'utf8') +
  '\n;this.__api = { CFG, ensureSetup_, addStamp, getStatus, previewPhone, flagTestPhone, fixInflatedPoints, backupNow_, ensureLogNoteCol_, findRow_, normPhone_, getSheet_ };';
// Share the outer realm's Date so `v instanceof Date` in Code.gs matches Dates the mock returns
// (single realm mirrors real Apps Script; without this, vm gives the sandbox a different Date).
const sandbox = { SpreadsheetApp, LockService, Utilities, Logger, console, Date };
vm.createContext(sandbox);
vm.runInContext(src, sandbox);
const A = sandbox.__api;
const CFG = A.CFG;

const TODAY = fmtDate(new Date(), 'x', 'yyyy-MM-dd');
function reset() { SpreadsheetApp._ss = newSS(); }
function sheet(n) { return SpreadsheetApp._ss.getSheetByName(n); }

// ================= TESTS =================
console.log('today = ' + TODAY + '\n');

// T1: once-per-day gate (the bug fix). manual code mode to avoid MD5.
console.log('T1 once-per-day gate');
reset(); CFG.CODE_MODE = 'manual'; CFG.REQUIRE_LOCATION = false;
A.ensureSetup_();
let r1 = A.addStamp('0912000111', '1234');
let r2 = A.addStamp('0912000111', '1234');   // same phone, same day
eq('1st stamp ok', r1.ok, true);
eq('1st points', r1.points, 1);
eq('2nd stamp blocked', r2.ok, false);
eq('2nd points still 1', r2.points, 1);
ok("block msg mentions 今天已經蓋過", /今天已經蓋過/.test(r2.msg || ''), r2.msg);
// stored lastStampDate must stay text and equal today
const lastCell = sheet('Points').getRange(2, 3).getValue();
ok('lastStampDate stored as text string', typeof lastCell === 'string' && lastCell === TODAY, JSON.stringify(lastCell));

// T2: legacy row where lastStampDate is a Date object -> still blocked (normDateCell_ handles Date)
console.log('\nT2 legacy Date value still blocks');
reset(); CFG.CODE_MODE = 'manual';
A.ensureSetup_();
const p = sheet('Points');
p.appendRow(['933828417', 3, new Date(), new Date()]); // col3 is a Date == today
const r3 = A.addStamp('0933828417', '1234');
eq('blocked despite Date-typed cell', r3.ok, false);
eq('points unchanged (3)', r3.points, 3);
// wrong code path still works
const r4 = A.addStamp('0900000000', '9999');
eq('wrong code rejected', r4.ok, false);
ok('wrong-code msg', /通關碼不正確/.test(r4.msg || ''), r4.msg);

// T3: previewPhone is read-only and counts rows
console.log('\nT3 previewPhone (read-only)');
reset();
sheet('Config') || SpreadsheetApp._ss.insertSheet('Config');
A.getSheet_('Points').setContents([
  ['phone', 'points', 'lastStampDate', 'updatedAt'],
  ['912345678', 2, '2026-06-27', '2026-06-27'],
  ['912345678', 1, '2026-06-25', '2026-06-25'],  // duplicate row
  ['965106328', 8, '2026-07-19', '2026-07-19'],
]);
A.getSheet_('Log').setContents([
  ['time', 'phone', 'lat', 'lng', 'distM', 'within150m', 'pointsAfter', 'note'],
  ['t1', '912345678', '', '', '', '', 1, ''],
  ['t2', '912345678', '', '', '', '', 1, ''],
  ['t3', '965106328', '', '', '', '', 8, ''],
]);
const beforePreview = JSON.stringify(sheet('Points').rows) + JSON.stringify(sheet('Log').rows);
const pv = A.previewPhone('0912345678');
eq('preview ok', pv.ok, true);
eq('preview points rows = 2', pv.pointsRows.length, 2);
eq('preview log rows = 2', pv.logRows.length, 2);
ok('preview did not modify data', JSON.stringify(sheet('Points').rows) + JSON.stringify(sheet('Log').rows) === beforePreview);

// T4: flagTestPhone deletes in Points, flags (not deletes) in Log, leaves others alone
console.log('\nT4 flagTestPhone');
const fr = A.flagTestPhone('0912345678');
eq('flag ok', fr.ok, true);
eq('points deleted = 2', fr.pointsDeleted, 2);
eq('log flagged = 2', fr.logFlagged, 2);
const pRows = sheet('Points').rows;
ok('Points no longer has 912345678', !pRows.some(r => A.normPhone_(r[0]) === '912345678'));
ok('Points still has 965106328', pRows.some(r => A.normPhone_(r[0]) === '965106328'));
const lRows = sheet('Log').rows;
eq('Log row count unchanged (header+3)', lRows.length, 4);
const noteCol = A.ensureLogNoteCol_() - 1;
ok('912345678 log rows flagged test', lRows.filter(r => A.normPhone_(r[1]) === '912345678').every(r => String(r[noteCol]).indexOf('test') !== -1));
ok('965106328 log row NOT flagged', lRows.filter(r => A.normPhone_(r[1]) === '965106328').every(r => String(r[noteCol] || '') === ''));

// T5: fixInflatedPoints changes Points + appends Log audit row
console.log('\nT5 fixInflatedPoints');
reset();
A.getSheet_('Points').setContents([
  ['phone', 'points', 'lastStampDate', 'updatedAt'],
  ['965106328', 8, '2026-07-19', '2026-07-19'],
]);
A.getSheet_('Log').setContents([
  ['time', 'phone', 'lat', 'lng', 'distM', 'within150m', 'pointsAfter', 'note'],
  ['t', '965106328', '', '', '', '', 8, ''],
]);
const fx = A.fixInflatedPoints('0965106328', 4);
eq('fix ok', fx.ok, true);
eq('before=8', fx.before, 8);
eq('after=4', fx.after, 4);
eq('Points value now 4', sheet('Points')._get(2, 2), 4);
const logNow = sheet('Log').rows;
eq('Log gained 1 audit row (header+1+1)', logNow.length, 3);
const auditNoteCol = A.ensureLogNoteCol_() - 1;
ok('audit row note = 人工修正 8→4', String(logNow[logNow.length - 1][auditNoteCol]).indexOf('人工修正 8→4') !== -1, logNow[logNow.length - 1][auditNoteCol]);
eq('audit row pointsAfter = 4', logNow[logNow.length - 1][6], 4);
// edge cases
eq('fix to same value rejected', A.fixInflatedPoints('0965106328', 4).ok, false);
eq('fix unknown phone rejected', A.fixInflatedPoints('0999999999', 1).ok, false);
eq('fix negative rejected', A.fixInflatedPoints('0965106328', -1).ok, false);

// T6: backupNow_
console.log('\nT6 backupNow_');
const bk = A.backupNow_();
eq('backup ok', bk.ok, true);
ok('backup name prefix', /^MWD集點卡_backup_/.test(bk.name), bk.name);
ok('backup has url', /^https:\/\/fake\//.test(bk.url), bk.url);

// T7: ensureLogNoteCol_ finds existing / creates missing
console.log('\nT7 ensureLogNoteCol_');
reset();
A.getSheet_('Log').setContents([['time', 'phone', 'lat', 'lng', 'distM', 'within150m', 'pointsAfter', 'note']]);
eq('finds existing note col at 8', A.ensureLogNoteCol_(), 8);
reset();
A.getSheet_('Log').setContents([['time', 'phone', 'lat', 'lng', 'distM', 'within150m', 'pointsAfter']]); // no note
const created = A.ensureLogNoteCol_();
eq('creates note col at 8', created, 8);
eq('header now says note', sheet('Log')._get(1, 8), 'note');

console.log('\n==============================');
console.log(`RESULT: ${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
