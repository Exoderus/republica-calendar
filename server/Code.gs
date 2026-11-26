/**
 * Republica | 1963 — Calendar evenimente · serviciul (Google Apps Script)
 * Script legat de Sheet-ul „Republica | 1963 — Calendar evenimente”.
 * Rulează pe contul care îl publică. Primește tokenul Google (doar email) al
 * utilizatorului, verifică în foaia „Echipa” că are „Activ” la „Acces” și abia
 * apoi citește sau scrie în calendar.
 *
 * Implementare: Aplicație web · Execută ca „Eu” · Acces „Oricine”.
 */

const CLIENT_ID = '541619680642-c3rufd8rfgg8mtuikvj0218irm9rohsj.apps.googleusercontent.com';
const SS_ID = '1WC0rnBY_tHXwj_Fxi5X2IucptQx5lWwG92HS9A1a4Pg';
const SHEET_EV = 'Evenimente';
const SHEET_LISTE = 'Liste';
const SHEET_LOG = 'Jurnal';
const SHEET_TEAM = 'Echipa';
const SHEET_TASKS = 'Sarcini';
const SHEET_RIDER = 'Rider';
const SHEET_MODULES = 'Module';
const SHEET_TEMPLATES = 'Șabloane';
const TZ = 'Europe/Bucharest';

// Spații cu reguli speciale la verificarea suprapunerilor (numele trebuie să fie identice cu cele din „Liste”)
const SPATIU_TOT = 'Toată locația';      // se suprapune cu orice spațiu
const SPATIU_EXTERN = 'În afara locației'; // nu se suprapune cu nimic
const STATUS_ANULAT = 'Anulat';
const STATUS_OPTIUNE = 'Opțiune';   // rezervare temporară, până la confirmare
const ZILE_OPTIUNE = 3;            // valabilitatea implicită a unei opțiuni, în zile

const COLS = ['id', 'titlu', 'tip', 'spatiu', 'data', 'start', 'sfarsit', 'montaj', 'demontaj',
  'status', 'responsabil', 'client', 'contact', 'participanti', 'valoare', 'avans',
  'tehnic', 'bar', 'note', 'creatDe', 'creatLa', 'modificatDe', 'modificatLa', 'arhivat', 'optiunePana',
  'module', 'barInterval', 'pachetBar', 'bilPlatforma', 'bilLink', 'capacitate', 'bilVandute', 'spectatoriReali', 'consumBar', 'feedback'];

const LABELS = {
  titlu: 'Titlu', tip: 'Tip', spatiu: 'Spațiu', data: 'Data', start: 'Ora început',
  sfarsit: 'Ora sfârșit', montaj: 'Montaj de la', demontaj: 'Demontaj până la', status: 'Status',
  responsabil: 'Responsabil', client: 'Client / partener', contact: 'Contact',
  participanti: 'Participanți estimați', valoare: 'Valoare (lei)', avans: 'Avans (lei)',
  tehnic: 'Necesar tehnic', bar: 'Bar și consum', note: 'Note',
  optiunePana: 'Opțiune valabilă până la', module: 'Module', barInterval: 'Bar: interval',
  pachetBar: 'Pachet bar inclus (lei)', bilPlatforma: 'Bilete: platformă', bilLink: 'Bilete: link',
  capacitate: 'Capacitate', bilVandute: 'Bilete vândute', spectatoriReali: 'Spectatori reali',
  consumBar: 'Consum bar realizat (lei)', feedback: 'Feedback client'
};

const EDITABLE = Object.keys(LABELS);
const NUMERIC = ['participanti', 'valoare', 'avans', 'pachetBar', 'capacitate', 'bilVandute', 'spectatoriReali', 'consumBar'];

/* ---------- Intrare web ---------- */

let USER_ = null; // { email, name } pentru cererea curentă

function doGet() {
  return json_({ ok: true, service: 'Republica | 1963 · Calendar' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const who = verify_(req.token);
    if (who.error) return json_(who);
    const me = member_(who.email);
    if (!me) return json_({ error: 'not_member', email: who.email });
    USER_ = me;
    if (req.action === 'bootstrap') return json_(getData());
    if (req.action === 'save') return json_(saveEvent(req.event || {}, req.opts || {}));
    if (req.action === 'archive') return json_(archiveEvent(String(req.id || ''), req.expectedModified));
    if (req.action === 'task') return json_(taskUpdate(String(req.id || ''), req.patch || {}));
    if (req.action === 'task_add') return json_(taskAdd(req.task || {}));
    if (req.action === 'rider_save') return json_(riderSave(req.row || {}));
    return json_({ error: 'bad_request' });
  } catch (err) {
    return json_({ error: 'server', message: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- autentificare ---------- */

function verify_(token) {
  if (!token) return { error: 'unauthorized' };
  const cache = CacheService.getScriptCache();
  const ck = 'tok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { error: 'unauthorized' };
  const t = JSON.parse(r.getContentText());
  const okIss = t.iss === 'accounts.google.com' || t.iss === 'https://accounts.google.com';
  if (t.aud !== CLIENT_ID || !okIss || String(t.email_verified) !== 'true' || Number(t.exp) * 1000 < Date.now()) return { error: 'unauthorized' };
  const out = { email: String(t.email).toLowerCase() };
  const ttl = Math.max(30, Math.min(600, Number(t.exp) - Math.floor(Date.now() / 1000)));
  cache.put(ck, JSON.stringify(out), ttl);
  return out;
}

/** Foaia „Echipa”: Nume | Email Google | Acces (Activ / Inactiv) */
function member_(email) {
  const sh = ss_().getSheetByName(SHEET_TEAM);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getDisplayValues();
  const r = rows.find(x => String(x[1]).trim().toLowerCase() === email && String(x[2]).trim().toLowerCase() === 'activ');
  return r ? { email: email, name: String(r[0]).trim() || email } : null;
}

/* ---------- API ---------- */

function getData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const events = readAll_(true).filter(e => e.arhivat !== 'DA');
    events.forEach(e => delete e._row);
    return {
      user: USER_,
      lists: readLists_(),
      modules: readModules_(),
      templates: readTemplates_(),
      events: events,
      tasks: readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM).map(strip_),
      rider: readTable_(SHEET_RIDER, RIDER_COLS, RIDER_NUM).map(strip_),
      sheetUrl: ss_().getUrl(),
      now: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * input: obiect cu câmpurile din LABELS (+ id la editare)
 * opts: { force: bool, expectedModified: string, repeat: { mode, count } }
 */
function saveEvent(input, opts) {
  opts = opts || {};
  const ev = clean_(input);
  const err = validate_(ev);
  if (err) return { error: err };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = readAll_(false);
    const user = currentUser_();
    const now = stamp_();
    const sh = ss_().getSheetByName(SHEET_EV);

    if (ev.id) {
      const cur = all.find(e => e.id === ev.id);
      if (!cur) return { error: 'Evenimentul nu mai există în Sheet.' };
      if (opts.expectedModified !== undefined && String(cur.modificatLa || '') !== String(opts.expectedModified || '')) {
        const fresh = Object.assign({}, cur); delete fresh._row;
        return { staleEdit: true, current: fresh };
      }
      const conflicts = findConflicts_(ev, all, [ev.id]);
      if (conflicts.length && !opts.force) return { needsConfirm: true, conflicts: conflicts };

      const changes = [];
      EDITABLE.forEach(k => {
        const a = String(cur[k] === undefined ? '' : cur[k]);
        const b = String(ev[k] === undefined ? '' : ev[k]);
        if (a !== b) changes.push(LABELS[k] + ': ' + (a || '—') + ' → ' + (b || '—'));
      });
      if (!changes.length) {
        const same = Object.assign({}, cur); delete same._row;
        return { ok: true, events: [same], unchanged: true };
      }
      const merged = Object.assign({}, cur, ev, { modificatDe: user, modificatLa: now });
      writeRow_(sh, cur._row, merged);
      log_(user, 'Modificat', merged.id, merged.titlu, changes.join('; '));
      const tasks = syncTasks_([{ ev: merged, prev: cur }]);
      delete merged._row;
      return { ok: true, events: [merged], conflicts: conflicts, tasks: tasks };
    }

    // Eveniment nou (eventual repetat)
    const dates = repeatDates_(ev.data, opts.repeat);
    const rows = [];
    let conflicts = [];
    dates.forEach(d => {
      const copy = Object.assign({}, ev, { data: d });
      conflicts = conflicts.concat(findConflicts_(copy, all.concat(rows), []).map(c => Object.assign(c, { pentru: d })));
      rows.push(Object.assign(copy, {
        id: newId_(all.concat(rows)), creatDe: user, creatLa: now, modificatDe: user, modificatLa: now, arhivat: ''
      }));
    });
    if (conflicts.length && !opts.force) return { needsConfirm: true, conflicts: conflicts };

    const start = sh.getLastRow() + 1;
    ensureRows_(sh, start + rows.length);
    formatRows_(sh, start, rows.length);
    sh.getRange(start, 1, rows.length, COLS.length).setValues(rows.map(toRow_));
    const det = rows.length > 1 ? 'Serie de ' + rows.length + ': ' + rows.map(r => r.data).join(', ') : summary_(rows[0]);
    log_(user, rows.length > 1 ? 'Adăugat (serie)' : 'Adăugat', rows.map(r => r.id).join(', '), ev.titlu, det);
    const tasks = syncTasks_(rows.map(r => ({ ev: r, prev: null })));
    return { ok: true, events: rows, conflicts: conflicts, tasks: tasks };
  } finally {
    lock.releaseLock();
  }
}

/** Nu ștergem nimic: evenimentul primește „Arhivat = DA” și dispare din aplicație, dar rămâne în Sheet. */
function archiveEvent(id, expectedModified) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = readAll_(false);
    const cur = all.find(e => e.id === id);
    if (!cur) return { error: 'Evenimentul nu mai există în Sheet.' };
    if (expectedModified !== undefined && String(cur.modificatLa || '') !== String(expectedModified || '')) {
      const fresh = Object.assign({}, cur); delete fresh._row;
      return { staleEdit: true, current: fresh };
    }
    const user = currentUser_();
    const merged = Object.assign({}, cur, { arhivat: 'DA', modificatDe: user, modificatLa: stamp_() });
    writeRow_(ss_().getSheetByName(SHEET_EV), cur._row, merged);
    log_(user, 'Arhivat', id, cur.titlu, summary_(cur));
    const tasks = syncTasks_([{ ev: merged, prev: cur }]);
    return { ok: true, id: id, tasks: tasks };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Citire și scriere ---------- */

function ss_() { return SpreadsheetApp.openById(SS_ID); }

function currentUser_() {
  return USER_ ? USER_.name : 'necunoscut';
}

function stamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function readAll_(fillMissingIds) {
  const sh = ss_().getSheetByName(SHEET_EV);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, COLS.length).getValues();
  const out = [];
  const missing = [];
  vals.forEach((r, i) => {
    if (!String(r[0]).trim() && !String(r[1]).trim()) return; // rând gol
    const e = rowToEv_(r);
    e._row = i + 2;
    if (!e.id) missing.push(e);
    out.push(e);
  });
  // Rânduri adăugate de mână în Sheet, fără ID: primesc un ID ca să poată fi editate din aplicație
  if (fillMissingIds && missing.length) {
    missing.forEach(e => {
      e.id = newId_(out);
      sh.getRange(e._row, 1).setNumberFormat('@').setValue(e.id);
    });
  }
  return out;
}

function rowToEv_(r) {
  const e = {};
  COLS.forEach((k, i) => {
    let v = r[i];
    if (k === 'data' || k === 'optiunePana') v = normDate_(v);
    else if (['start', 'sfarsit', 'montaj', 'demontaj'].indexOf(k) >= 0) v = normTime_(v);
    else if (NUMERIC.indexOf(k) >= 0) v = (v === '' || v === null) ? '' : Number(String(v).replace(/\s/g, '').replace(',', '.')) || 0;
    else if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
    else v = String(v === null || v === undefined ? '' : v).trim();
    e[k] = v;
  });
  return e;
}

function toRow_(e) {
  return COLS.map(k => {
    const v = e[k];
    if (NUMERIC.indexOf(k) >= 0) return (v === '' || v === null || v === undefined) ? '' : Number(v);
    return v === null || v === undefined ? '' : String(v);
  });
}

function writeRow_(sh, row, e) {
  formatRows_(sh, row, 1);
  sh.getRange(row, 1, 1, COLS.length).setValues([toRow_(e)]);
}

function ensureRows_(sh, needed) {
  const max = sh.getMaxRows();
  if (needed > max) sh.insertRowsAfter(max, Math.max(200, needed - max));
}

function formatRows_(sh, start, n) {
  formatCols_(sh, start, n, COLS, NUMERIC);
}

/** Pune formatul fiecărei coloane: text pentru tot, număr pentru coloanele numerice. */
function formatCols_(sh, start, n, cols, numeric) {
  let i = 0;
  while (i < cols.length) {
    const isNum = numeric.indexOf(cols[i]) >= 0;
    let j = i;
    while (j + 1 < cols.length && (numeric.indexOf(cols[j + 1]) >= 0) === isNum) j++;
    sh.getRange(start, i + 1, n, j - i + 1).setNumberFormat(isNum ? '#,##0' : '@');
    i = j + 1;
  }
}

function readLists_() {
  const sh = ss_().getSheetByName(SHEET_LISTE);
  const vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 8).getValues();
  const col = i => vals.map(r => String(r[i]).trim()).filter(Boolean);
  const spatii = vals.filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), color: String(r[1]).trim() || '#4A4A4A' }));
  const set = {};
  const lastCol = sh.getLastColumn();
  if (lastCol >= 11) {
    sh.getRange(2, 10, Math.max(sh.getLastRow() - 1, 1), 2).getDisplayValues()
      .forEach(r => { if (String(r[0]).trim()) set[String(r[0]).trim()] = String(r[1]).trim(); });
  }
  return {
    spatii: spatii, tipuri: col(3), responsabili: col(5), statusuri: col(7),
    pragConcession: Number(set['Prag concession (spectatori)']) || 150,
    spatiuConcession: set['Spațiu concession'] || 'Sala mare'
  };
}

function log_(user, action, id, titlu, details) {
  const sh = ss_().getSheetByName(SHEET_LOG);
  const row = sh.getLastRow() + 1;
  ensureRows_(sh, row);
  sh.getRange(row, 1, 1, 6).setNumberFormat('@').setValues([[stamp_(), user + (USER_ ? ' <' + USER_.email + '>' : ''), action, id, titlu, details]]);
}


/* ---------- Module, șabloane, sarcini, rider ---------- */

const TASK_COLS = ['id', 'eventId', 'modul', 'sarcina', 'responsabil', 'zile', 'termen', 'stare', 'sablon', 'note', 'modificatDe', 'modificatLa'];
const TASK_NUM = ['zile'];
const RIDER_COLS = ['id', 'eventId', 'echipament', 'cantitate', 'sursa', 'partener', 'cost', 'pret', 'stare', 'note', 'modificatDe', 'modificatLa'];
const RIDER_NUM = ['cantitate', 'cost', 'pret'];
const T_OPEN = 'De făcut', T_DONE = 'Gata', T_CANCEL = 'Anulată';

function strip_(o) { delete o._row; return o; }

function readModules_() {
  const sh = ss_().getSheetByName(SHEET_MODULES);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      cod: String(r[0]).trim(), nume: String(r[1]).trim() || String(r[0]).trim(),
      tipuri: String(r[2]).split(',').map(x => x.trim()).filter(Boolean)
    }));
}

function readTemplates_() {
  const sh = ss_().getSheetByName(SHEET_TEMPLATES);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim() && String(r[2]).trim())
    .map(r => ({
      cod: String(r[0]).trim(), modul: String(r[1]).trim(), sarcina: String(r[2]).trim(),
      responsabil: String(r[3]).trim(), zile: Number(String(r[4]).replace(',', '.')) || 0
    }));
}

function readTable_(name, cols, numeric) {
  const sh = ss_().getSheetByName(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues()
    .map((r, i) => {
      const o = { _row: i + 2 };
      cols.forEach((k, j) => {
        let v = r[j];
        if (numeric.indexOf(k) >= 0) v = (v === '' || v === null) ? '' : Number(v) || 0;
        else if (k === 'termen') v = normDate_(v);
        else if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
        else v = String(v === null || v === undefined ? '' : v).trim();
        o[k] = v;
      });
      return o;
    })
    .filter(o => o.id);
}

function tableRow_(o, cols, numeric) {
  return cols.map(k => {
    const v = o[k];
    if (numeric.indexOf(k) >= 0) return (v === '' || v === null || v === undefined) ? '' : Number(v);
    return v === null || v === undefined ? '' : String(v);
  });
}

function writeTableRow_(name, cols, numeric, row, o) {
  const sh = ss_().getSheetByName(name);
  formatCols_(sh, row, 1, cols, numeric);
  sh.getRange(row, 1, 1, cols.length).setValues([tableRow_(o, cols, numeric)]);
}

function appendTableRows_(name, cols, numeric, list) {
  if (!list.length) return;
  const sh = ss_().getSheetByName(name);
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + list.length);
  formatCols_(sh, start, list.length, cols, numeric);
  sh.getRange(start, 1, list.length, cols.length).setValues(list.map(o => tableRow_(o, cols, numeric)));
}

function newSimpleId_(prefix, existing) {
  const used = {};
  existing.forEach(e => used[e.id] = true);
  let id;
  do { id = prefix + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase(); } while (used[id]);
  return id;
}

function modsOf_(ev) {
  return String(ev.module || '').split(',').map(x => x.trim()).filter(Boolean);
}

/**
 * Ține sarcinile evenimentelor la zi față de modulele bifate:
 * modul bifat → sarcinile din șabloane (create o singură dată);
 * modul debifat, eveniment anulat sau arhivat → sarcinile nefăcute devin „Anulată”;
 * data sau responsabilul schimbat → termenele și responsabilii sarcinilor nefăcute se actualizează.
 * Întoarce sarcinile modificate sau create.
 */
function syncTasks_(pairs) {
  const templates = readTemplates_();
  const all = readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM);
  const user = currentUser_(), now = stamp_();
  const changed = [], created = [];
  pairs.forEach(({ ev, prev }) => {
    const dead = ev.arhivat === 'DA' || ev.status === STATUS_ANULAT;
    const mods = dead ? [] : modsOf_(ev);
    const mine = all.filter(t => t.eventId === ev.id);
    templates.forEach(tp => {
      const want = mods.indexOf(tp.modul) >= 0;
      const t = mine.find(x => x.sablon === tp.cod);
      const resp = tp.responsabil === '@responsabil' ? (ev.responsabil || '') : tp.responsabil;
      const termen = addDays_(ev.data, -tp.zile);
      if (want && !t) {
        const nt = {
          id: newSimpleId_('T', all.concat(created)), eventId: ev.id, modul: tp.modul, sarcina: tp.sarcina,
          responsabil: resp, zile: tp.zile, termen: termen, stare: T_OPEN, sablon: tp.cod, note: '',
          modificatDe: user, modificatLa: now
        };
        created.push(nt);
        return;
      }
      if (!t) return;
      let dirty = false;
      if (want && t.stare === T_CANCEL) { t.stare = T_OPEN; dirty = true; }
      if (!want && t.stare === T_OPEN) { t.stare = T_CANCEL; dirty = true; }
      if (t.stare === T_OPEN) {
        if (t.termen !== termen && prev && prev.data !== ev.data) { t.termen = termen; dirty = true; }
        if (tp.responsabil === '@responsabil' && prev && prev.responsabil !== ev.responsabil && t.responsabil === (prev.responsabil || '')) {
          t.responsabil = ev.responsabil || ''; dirty = true;
        }
      }
      if (dirty) { t.modificatDe = user; t.modificatLa = now; changed.push(t); }
    });
    // sarcinile adăugate de mână rămân, dar se anulează odată cu evenimentul
    if (dead) mine.filter(t => !t.sablon && t.stare === T_OPEN).forEach(t => {
      t.stare = T_CANCEL; t.modificatDe = user; t.modificatLa = now; changed.push(t);
    });
  });
  changed.forEach(t => writeTableRow_(SHEET_TASKS, TASK_COLS, TASK_NUM, t._row, t));
  appendTableRows_(SHEET_TASKS, TASK_COLS, TASK_NUM, created);
  return changed.concat(created).map(t => strip_(Object.assign({}, t)));
}

function taskUpdate(id, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM);
    const t = all.find(x => x.id === id);
    if (!t) return { error: 'Sarcina nu mai există.' };
    const before = t.stare + ' · ' + t.responsabil + ' · ' + t.termen;
    if (patch.stare !== undefined && [T_OPEN, T_DONE, T_CANCEL].indexOf(patch.stare) >= 0) t.stare = patch.stare;
    if (patch.responsabil !== undefined) t.responsabil = String(patch.responsabil).trim().slice(0, 100);
    if (patch.termen !== undefined) t.termen = normDate_(patch.termen);
    if (patch.note !== undefined) t.note = safeText_(patch.note);
    t.modificatDe = currentUser_(); t.modificatLa = stamp_();
    writeTableRow_(SHEET_TASKS, TASK_COLS, TASK_NUM, t._row, t);
    log_(t.modificatDe, 'Sarcină', t.eventId, t.sarcina, before + ' → ' + t.stare + ' · ' + t.responsabil + ' · ' + t.termen);
    return { ok: true, task: strip_(t) };
  } finally {
    lock.releaseLock();
  }
}

function taskAdd(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sarcina = safeText_(input.sarcina);
    if (!sarcina || !input.eventId) return { error: 'Scrie ce trebuie făcut.' };
    const all = readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM);
    const t = {
      id: newSimpleId_('T', all), eventId: String(input.eventId), modul: String(input.modul || ''), sarcina: sarcina,
      responsabil: String(input.responsabil || '').trim(), zile: '', termen: normDate_(input.termen), stare: T_OPEN,
      sablon: '', note: safeText_(input.note), modificatDe: currentUser_(), modificatLa: stamp_()
    };
    appendTableRows_(SHEET_TASKS, TASK_COLS, TASK_NUM, [t]);
    log_(t.modificatDe, 'Sarcină nouă', t.eventId, t.sarcina, t.responsabil + ' · ' + t.termen);
    return { ok: true, task: t };
  } finally {
    lock.releaseLock();
  }
}

function riderSave(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = readTable_(SHEET_RIDER, RIDER_COLS, RIDER_NUM);
    const num = v => { v = String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'); return v === '' || isNaN(Number(v)) ? '' : Number(v); };
    const o = {
      eventId: String(input.eventId || ''), echipament: safeText_(input.echipament), cantitate: num(input.cantitate),
      sursa: input.sursa === 'Partener' ? 'Partener' : 'În casă', partener: safeText_(input.partener),
      cost: num(input.cost), pret: num(input.pret),
      stare: ['Cerut', 'Confirmat', 'Livrat', 'Montat', 'Anulat'].indexOf(input.stare) >= 0 ? input.stare : 'Cerut',
      note: safeText_(input.note), modificatDe: currentUser_(), modificatLa: stamp_()
    };
    if (!o.echipament || !o.eventId) return { error: 'Scrie echipamentul.' };
    if (input.id) {
      const cur = all.find(x => x.id === input.id);
      if (!cur) return { error: 'Rândul nu mai există.' };
      const merged = Object.assign(cur, o);
      writeTableRow_(SHEET_RIDER, RIDER_COLS, RIDER_NUM, cur._row, merged);
      log_(o.modificatDe, 'Rider', o.eventId, o.echipament, o.stare + ' · ' + o.sursa + (o.partener ? ' (' + o.partener + ')' : '') + ' · cost ' + o.cost + ' · preț ' + o.pret);
      return { ok: true, row: strip_(merged) };
    }
    o.id = newSimpleId_('R', all);
    appendTableRows_(SHEET_RIDER, RIDER_COLS, RIDER_NUM, [o]);
    log_(o.modificatDe, 'Rider nou', o.eventId, o.echipament, o.sursa + (o.partener ? ' (' + o.partener + ')' : '') + ' · cost ' + o.cost + ' · preț ' + o.pret);
    return { ok: true, row: o };
  } finally {
    lock.releaseLock();
  }
}

function safeText_(v) {
  v = String(v === undefined || v === null ? '' : v).trim().slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/* ---------- Validare și normalizare ---------- */

function clean_(input) {
  const ev = {};
  if (input.id) ev.id = String(input.id).trim();
  EDITABLE.forEach(k => {
    let v = input[k];
    if (v === undefined || v === null) v = '';
    if (NUMERIC.indexOf(k) >= 0) {
      v = String(v).replace(/\s/g, '').replace(',', '.');
      v = v === '' ? '' : (isNaN(Number(v)) ? '' : Number(v));
    } else {
      v = String(v).trim().slice(0, 2000);
      if (/^[=+\-@]/.test(v)) v = "'" + v; // textul nu devine formulă în Sheet
    }
    ev[k] = v;
  });
  ev.data = normDate_(ev.data);
  // Opțiunea (rezervarea temporară) are mereu o dată de expirare; celelalte statusuri nu
  if (ev.status === STATUS_OPTIUNE) {
    ev.optiunePana = normDate_(ev.optiunePana) || addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), ZILE_OPTIUNE);
  } else {
    ev.optiunePana = '';
  }
  ['start', 'sfarsit', 'montaj', 'demontaj'].forEach(k => ev[k] = normTime_(ev[k]));
  return ev;
}

function validate_(ev) {
  if (!ev.titlu) return 'Lipsește titlul.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ev.data)) return 'Data nu e validă.';
  if (!ev.spatiu) return 'Alege spațiul.';
  if (!ev.start) return 'Lipsește ora de început.';
  return '';
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/);
  if (m) return m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);
  return s;
}

function normTime_(v) {
  if (v instanceof Date) return pad_(v.getHours()) + ':' + pad_(v.getMinutes());
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/^(\d{1,2})(?:[:.](\d{2}))?$/);
  if (m && Number(m[1]) < 24 && Number(m[2] || 0) < 60) return pad_(m[1]) + ':' + (m[2] || '00');
  return s;
}

function addDays_(iso, n) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());
}

function pad_(n) { return ('0' + Number(n)).slice(-2); }

function newId_(existing) {
  const used = {};
  existing.forEach(e => used[e.id] = true);
  const base = 'E' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-';
  let id;
  do { id = base + Math.random().toString(36).slice(2, 6).toUpperCase(); } while (used[id]);
  return id;
}

function summary_(e) {
  return [e.data, e.start + (e.sfarsit ? '–' + e.sfarsit : ''), e.spatiu, e.tip, e.status, e.responsabil]
    .filter(Boolean).join(' · ');
}

/* ---------- Suprapuneri ---------- */

function toMin_(t) {
  const m = String(t || '').match(/^(\d{2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

/** Intervalul ocupat: de la montaj (sau start) până la demontaj (sau sfârșit; implicit start + 2h). */
function interval_(e) {
  let s = toMin_(e.montaj);
  if (s === null) s = toMin_(e.start);
  if (s === null) s = 0;
  let f = toMin_(e.demontaj);
  if (f === null) f = toMin_(e.sfarsit);
  if (f === null) f = Math.min(s + 120, 24 * 60);
  if (f <= s) f += 24 * 60; // trece de miezul nopții
  return [s, f];
}

function spacesClash_(a, b) {
  if (!a || !b || a === SPATIU_EXTERN || b === SPATIU_EXTERN) return false;
  return a === b || a === SPATIU_TOT || b === SPATIU_TOT;
}

function findConflicts_(ev, all, excludeIds) {
  if (ev.status === STATUS_ANULAT) return [];
  const [s, f] = interval_(ev);
  return all.filter(o =>
    excludeIds.indexOf(o.id) < 0 &&
    o.arhivat !== 'DA' &&
    o.status !== STATUS_ANULAT &&
    o.data === ev.data &&
    spacesClash_(o.spatiu, ev.spatiu) &&
    (() => { const [os, of] = interval_(o); return os < f && s < of; })()
  ).map(o => ({
    id: o.id, titlu: o.titlu, data: o.data, start: o.start, sfarsit: o.sfarsit,
    spatiu: o.spatiu, status: o.status, responsabil: o.responsabil, optiunePana: o.optiunePana
  }));
}

/* ---------- Repetare ---------- */

/** mode: '' | 'saptamanal' | 'doua' | 'lunar_zi' | 'lunar_pozitie'; count = numărul total de apariții */
function repeatDates_(first, repeat) {
  const mode = repeat && repeat.mode;
  const count = Math.min(Math.max(Number(repeat && repeat.count) || 1, 1), 52);
  if (!mode || count === 1) return [first];
  const p = first.split('-').map(Number);
  const d0 = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  const out = [];
  const fmt = d => d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());

  if (mode === 'saptamanal' || mode === 'doua') {
    const step = mode === 'saptamanal' ? 7 : 14;
    for (let i = 0; i < count; i++) out.push(fmt(new Date(d0.getTime() + i * step * 86400000)));
    return out;
  }
  if (mode === 'lunar_zi') {
    for (let i = 0; out.length < count && i < count * 2; i++) {
      const d = new Date(Date.UTC(p[0], p[1] - 1 + i, p[2]));
      if (d.getUTCDate() === p[2]) out.push(fmt(d)); // sare peste lunile fără ziua respectivă
    }
    return out;
  }
  if (mode === 'lunar_pozitie') {
    const wd = d0.getUTCDay();
    const daysInMonth = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    const isLast = p[2] + 7 > daysInMonth(p[0], p[1] - 1);
    const nth = Math.ceil(p[2] / 7);
    for (let i = 0; i < count; i++) {
      const y = new Date(Date.UTC(p[0], p[1] - 1 + i, 1));
      const yy = y.getUTCFullYear(), mm = y.getUTCMonth();
      let day;
      if (isLast) {
        const lastD = daysInMonth(yy, mm);
        const lastWd = new Date(Date.UTC(yy, mm, lastD)).getUTCDay();
        day = lastD - ((lastWd - wd + 7) % 7);
      } else {
        const firstWd = new Date(Date.UTC(yy, mm, 1)).getUTCDay();
        day = 1 + ((wd - firstWd + 7) % 7) + (nth - 1) * 7;
      }
      out.push(fmt(new Date(Date.UTC(yy, mm, day))));
    }
    return out;
  }
  return [first];
}
