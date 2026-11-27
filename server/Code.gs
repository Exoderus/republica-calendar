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
const ZILE_GRATIE = 2;             // după atâtea zile de la expirare, opțiunea se anulează singură
const APP_URL = 'https://exoderus.github.io/republica-calendar/';

const COLS = ['id', 'titlu', 'tip', 'spatiu', 'data', 'start', 'sfarsit', 'montaj', 'demontaj',
  'status', 'responsabil', 'client', 'contact', 'participanti', 'valoare', 'avans',
  'tehnic', 'bar', 'note', 'creatDe', 'creatLa', 'modificatDe', 'modificatLa', 'arhivat', 'optiunePana',
  'module', 'barInterval', 'pachetBar', 'bilPlatforma', 'bilLink', 'capacitate', 'bilVandute', 'spectatoriReali', 'consumBar', 'feedback',
  'deschide', 'inchide', 'echipOrg', 'configurare', 'contactTehnic', 'acces', 'proba', 'montajData'];

const LABELS = {
  titlu: 'Titlu', tip: 'Tip', spatiu: 'Spațiu', data: 'Data', start: 'Ora început',
  sfarsit: 'Ora sfârșit', montaj: 'Montaj de la', demontaj: 'Demontaj până la', status: 'Status',
  responsabil: 'Responsabil', client: 'Client / partener', contact: 'Contact',
  participanti: 'Participanți estimați', valoare: 'Valoare (lei)', avans: 'Avans (lei)',
  tehnic: 'Necesar tehnic', bar: 'Bar și consum', note: 'Note',
  optiunePana: 'Opțiune valabilă până la', module: 'Module', barInterval: 'Bar: interval',
  pachetBar: 'Pachet bar inclus (lei)', bilPlatforma: 'Bilete: platformă', bilLink: 'Bilete: link',
  capacitate: 'Capacitate', bilVandute: 'Bilete vândute', spectatoriReali: 'Spectatori reali',
  consumBar: 'Consum bar realizat (lei)', feedback: 'Feedback client',
  deschide: 'Deschide la montaj', inchide: 'Închide după demontaj', echipOrg: 'Vin cu echipamentul lor',
  configurare: 'Configurare sală', contactTehnic: 'Contact tehnic organizator',
  acces: 'Acces public de la', proba: 'Probă de sunet', montajData: 'Montaj din ziua'
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
    if (me.readonly) return json_({ error: 'readonly', message: 'Ai acces doar pentru citire.' });
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
  const ok = ['activ', 'doar citire'];
  const r = rows.find(x => String(x[1]).trim().toLowerCase() === email && ok.indexOf(String(x[2]).trim().toLowerCase()) >= 0);
  return r ? { email: email, name: String(r[0]).trim() || email, readonly: String(r[2]).trim().toLowerCase() === 'doar citire' } : null;
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
      notifyChange_(cur, merged);
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
    notifyChange_(cur, merged);
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
  const nc = Math.min(Math.max(sh.getLastColumn(), 5), 7);
  return sh.getRange(2, 1, sh.getLastRow() - 1, nc).getDisplayValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim() && String(r[2]).trim())
    .map(r => ({
      cod: String(r[0]).trim(), modul: String(r[1]).trim(), sarcina: String(r[2]).trim(),
      responsabil: String(r[3]).trim(),
      // „R+2” = la 2 zile după rezervare; un număr simplu = zile înainte de eveniment (negativ = după)
      rez: /^R/i.test(String(r[4]).trim()),
      zile: Number(String(r[4]).trim().replace(/^R\+?/i, '').replace(',', '.')) || 0,
      critica: /^(da|x|1|true)$/i.test(String(r[5] || '').trim()),
      prelungire: Number(r[6]) || 0
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
  const m = String(ev.module || '').split(',').map(x => x.trim()).filter(Boolean);
  // dacă au oră de montaj sau demontaj, cineva trebuie să deschidă și să închidă
  if ((ev.montaj || ev.demontaj) && m.indexOf('montaj') < 0) m.push('montaj');
  return m;
}

/** Responsabilul unei sarcini din șablon: „@câmp” = persoana din acel câmp al evenimentului. */
function tplResp_(tp, ev) {
  return tp.responsabil.charAt(0) === '@' ? String(ev[tp.responsabil.slice(1)] || '') : tp.responsabil;
}

/** Textul sarcinii, cu {montaj}, {demontaj}, {start}, {sfarsit} înlocuite. */
function tplText_(tp, ev) {
  const val = k => k === 'montaj' && ev.montaj && ev.montajData ? fmtZi_(ev.montajData) + ' ' + ev.montaj : (ev[k] || '');
  return tp.sarcina
    .replace(/\s*\(\{(montaj|demontaj|start|sfarsit|acces|proba)\}\)/g, (m, k) => val(k) ? ' (' + val(k) + ')' : '')
    .replace(/\{(montaj|demontaj|start|sfarsit|acces|proba)\}/g, (_, k) => val(k) || '?');
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
      const resp = tplResp_(tp, ev), text = tplText_(tp, ev);
      let termen = addDays_(ev.data, -tp.zile);
      if (tp.rez) {
        const rezData = String(ev.creatLa || stamp_()).slice(0, 10);
        termen = addDays_(rezData, tp.zile);
        if (termen > ev.data) termen = ev.data;
      }
      if (want && !t) {
        const nt = {
          id: newSimpleId_('T', all.concat(created)), eventId: ev.id, modul: tp.modul, sarcina: text,
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
        if (t.termen !== termen && prev && prev.data !== ev.data && !tp.rez) { t.termen = termen; dirty = true; }
        if (tp.responsabil.charAt(0) === '@' && prev && t.responsabil === tplResp_(tp, prev) && t.responsabil !== resp) {
          t.responsabil = resp; dirty = true;
        }
        if (tp.sarcina.indexOf('{') >= 0 && t.sarcina !== text) { t.sarcina = text; dirty = true; }
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
    let event = null;
    if (patch.stare === T_DONE) {
      const tp = readTemplates_().find(x => x.cod === t.sablon);
      if (tp && tp.prelungire > 0) event = extendOption_(t.eventId, tp.prelungire, t.sarcina);
    }
    return { ok: true, task: strip_(t), event: event };
  } finally {
    lock.releaseLock();
  }
}

/** Prelungește opțiunea unui eveniment (ex. după ce contractul a fost trimis clientului). */
function extendOption_(eventId, zile, motiv) {
  const all = readAll_(false);
  const cur = all.find(e => e.id === eventId);
  if (!cur || cur.status !== STATUS_OPTIUNE) return null;
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const nou = addDays_(today, zile);
  if (cur.optiunePana && cur.optiunePana >= nou) return null;
  const merged = Object.assign({}, cur, { optiunePana: nou, modificatDe: currentUser_(), modificatLa: stamp_() });
  writeRow_(ss_().getSheetByName(SHEET_EV), cur._row, merged);
  log_(currentUser_(), 'Opțiune prelungită', cur.id, cur.titlu, (cur.optiunePana || '—') + ' → ' + nou + ' (' + motiv + ')');
  delete merged._row;
  return merged;
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
  ['start', 'sfarsit', 'montaj', 'demontaj', 'acces', 'proba'].forEach(k => ev[k] = normTime_(ev[k]));
  ev.montajData = normDate_(ev.montajData);
  if (ev.montajData && (ev.montajData >= ev.data || !ev.montaj)) ev.montajData = ''; // doar o zi anterioară, și doar cu oră
  ev.module = modsOf_(ev).join(',');
  return ev;
}

function validate_(ev) {
  if (!ev.titlu) return 'Lipsește titlul.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ev.data)) return 'Data nu e validă.';
  if (!ev.spatiu) return 'Alege spațiul.';
  if (!ev.start) return 'Lipsește ora de început.';
  if (ev.montajData && daysBetween_(ev.montajData, ev.data) > 7) return 'Montajul poate începe cu cel mult 7 zile înainte.';
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

/** Ziua ca număr (zile de la 1970), ca să comparăm intervale peste mai multe zile. */
function dayNum_(d) {
  const p = String(d || '').split('-').map(Number);
  return p.length === 3 ? Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5) : 0;
}
function daysBetween_(a, b) { return dayNum_(b) - dayNum_(a); }

/**
 * Intervalul ocupat, în minute absolute: de la montaj (sau start) până la demontaj
 * (sau sfârșit; implicit start + 2h). Montajul poate fi într-o zi anterioară,
 * iar demontajul poate trece de miezul nopții.
 */
function interval_(e) {
  const base = dayNum_(e.data) * 1440;
  let st = toMin_(e.start); if (st === null) st = 0;
  let s = toMin_(e.montaj);
  let sBase = base;
  if (s === null) s = st;
  else if (e.montajData) sBase = dayNum_(e.montajData) * 1440;
  else if (s > st) s = st; // montaj trecut după start în aceeași zi: îl tratăm ca start
  let f = toMin_(e.demontaj);
  if (f === null) f = toMin_(e.sfarsit);
  if (f === null) f = Math.min(st + 120, 24 * 60);
  if (f <= st) f += 24 * 60; // trece de miezul nopții
  return [sBase + s, base + f];
}

function spacesClash_(a, b) {
  if (!a || !b || a === SPATIU_EXTERN || b === SPATIU_EXTERN) return false;
  return a === b || a === SPATIU_TOT || b === SPATIU_TOT;
}

/** Ferestrele în care oamenii noștri sunt prinși la un eveniment: montaj (deschide) și demontaj (închide). */
function staffWindows_(e) {
  const base = dayNum_(e.data) * 1440, st = toMin_(e.start);
  if (st === null) return [];
  const out = [];
  if (e.deschide && toMin_(e.montaj) !== null) {
    const s0 = (e.montajData ? dayNum_(e.montajData) * 1440 : base) + toMin_(e.montaj);
    out.push({ who: e.deschide, rol: 'montaj', s: s0, f: base + st });
  }
  if (e.inchide && toMin_(e.demontaj) !== null) {
    let fin = toMin_(e.sfarsit); if (fin === null) fin = st + 120; if (fin <= st) fin += 1440;
    let d = toMin_(e.demontaj); if (d <= st) d += 1440;
    out.push({ who: e.inchide, rol: 'demontaj', s: base + fin, f: base + Math.max(d, fin + 1) });
  }
  return out;
}

function findConflicts_(ev, all, excludeIds) {
  if (ev.status === STATUS_ANULAT) return [];
  const [s, f] = interval_(ev);
  const mine = staffWindows_(ev);
  const people = mine.length ? all.filter(o =>
    excludeIds.indexOf(o.id) < 0 && o.arhivat !== 'DA' && o.status !== STATUS_ANULAT &&
    Math.abs(daysBetween_(o.data, ev.data)) <= 8 && !spacesClash_(o.spatiu, ev.spatiu)
  ).map(o => {
    const hit = [];
    staffWindows_(o).forEach(w => mine.forEach(m => { if (w.who === m.who && w.s < m.f && m.s < w.f) hit.push(m.who + ' e la ' + m.rol + ' aici și la ' + w.rol + ' acolo'); }));
    return hit.length ? { id: o.id, titlu: o.titlu, data: o.data, start: o.start, sfarsit: o.sfarsit, spatiu: o.spatiu, status: o.status, responsabil: o.responsabil, optiunePana: o.optiunePana, motiv: hit.join('; ') } : null;
  }).filter(Boolean) : [];
  return all.filter(o =>
    excludeIds.indexOf(o.id) < 0 &&
    o.arhivat !== 'DA' &&
    o.status !== STATUS_ANULAT &&
    Math.abs(daysBetween_(o.data, ev.data)) <= 8 &&
    spacesClash_(o.spatiu, ev.spatiu) &&
    (() => { const [os, of] = interval_(o); return os < f && s < of; })()
  ).map(o => ({
    id: o.id, titlu: o.titlu, data: o.data, start: o.start, sfarsit: o.sfarsit,
    spatiu: o.spatiu, status: o.status, responsabil: o.responsabil, optiunePana: o.optiunePana
  })).concat(people);
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


/* ---------- Zilnic: anulare opțiuni expirate + rezumat pe email ---------- */

/** Rulează o singură dată din editor: pornește rularea zilnică (în jur de 7:30). */
function setupTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyRun').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyRun').timeBased().atHour(7).nearMinute(30).everyDays(1).inTimezone(TZ).create();
  return 'Rularea zilnică e pornită.';
}

/**
 * Pentru test: trimite rezumatul de azi al unei persoane (implicit Ioana) la adresa lui Mihai din Echipa,
 * fără să anuleze nimic și fără să scrie altcuiva.
 */
function testDigest(nume) {
  const team = team_();
  const mihai = team.find(p => p.name === 'Mihai');
  if (!mihai) return 'Mihai nu are adresă în Echipa.';
  return dailyRun({ onlyTo: nume || 'Ioana', previewTo: mihai.email, dryCancel: true });
}

function dailyRun(o) {
  o = (o && o.onlyTo !== undefined) ? o : {};
  USER_ = { email: 'automat', name: 'Calendar (automat)' };
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!o.dryCancel) backup_(today);
  const cancelled = o.dryCancel ? [] : cancelExpiredOptions_(today);
  return sendDigests_(today, cancelled, o.onlyTo || '', o.previewTo || '');
}

/** Opțiunile expirate de mai mult de ZILE_GRATIE zile trec la „Anulat”. */
function cancelExpiredOptions_(today) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const all = readAll_(false);
    const sh = ss_().getSheetByName(SHEET_EV);
    const limit = addDays_(today, -ZILE_GRATIE);
    const out = [];
    all.filter(e => e.arhivat !== 'DA' && e.status === STATUS_OPTIUNE && e.optiunePana && e.optiunePana < limit)
      .forEach(cur => {
        const note = (cur.note ? cur.note + '\n' : '') + 'Opțiune expirată pe ' + fmtZi_(cur.optiunePana) + ', anulată automat pe ' + fmtZi_(today) + '.';
        const merged = Object.assign({}, cur, { status: STATUS_ANULAT, optiunePana: '', note: note, modificatDe: USER_.name, modificatLa: stamp_() });
        writeRow_(sh, cur._row, merged);
        log_(USER_.name, 'Anulat automat', cur.id, cur.titlu, 'Opțiune expirată pe ' + cur.optiunePana);
        syncTasks_([{ ev: merged, prev: cur }]);
        out.push(merged);
      });
    return out;
  } finally {
    lock.releaseLock();
  }
}

function team_() {
  const sh = ss_().getSheetByName(SHEET_TEAM);
  return sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getDisplayValues()
    .filter(r => String(r[0]).trim() && /@/.test(r[1]) && String(r[2]).trim().toLowerCase() === 'activ')
    .map(r => ({ name: String(r[0]).trim(), email: String(r[1]).trim() }));
}

const ZILE_SC_ = ['dum', 'lun', 'mar', 'mie', 'joi', 'vin', 'sâm'];
const LUNI_SC_ = ['ian', 'feb', 'mar', 'apr', 'mai', 'iun', 'iul', 'aug', 'sep', 'oct', 'noi', 'dec'];
function fmtZi_(d) {
  const p = String(d || '').split('-').map(Number);
  if (p.length !== 3) return String(d || '');
  const dt = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return ZILE_SC_[dt.getUTCDay()] + ' ' + p[2] + ' ' + LUNI_SC_[p[1] - 1];
}
function h_(s) { return String(s == null ? '' : s).replace(/^'/, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function sendDigests_(today, cancelled, onlyName, previewTo) {
  const events = readAll_(true).filter(e => e.arhivat !== 'DA');
  const byId = {}; events.forEach(e => byId[e.id] = e);
  const allTasks = readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM), templates = readTemplates_();
  const tasks = allTasks.filter(t => t.stare === T_OPEN && byId[t.eventId] && byId[t.eventId].status !== STATUS_ANULAT);
  const tomorrow = addDays_(today, 1);
  const live = events.filter(e => e.status !== STATUS_ANULAT);
  let sent = 0;
  team_().forEach(p => {
    if (onlyName && p.name !== onlyName) return;
    const mine = tasks.filter(t => t.responsabil === p.name && t.termen && t.termen <= tomorrow)
      .sort((a, b) => a.termen.localeCompare(b.termen));
    const late = mine.filter(t => t.termen < today), now = mine.filter(t => t.termen >= today);
    const roles = e => [e.responsabil === p.name && 'responsabil', e.deschide === p.name && 'montaj', e.inchide === p.name && 'demontaj'].filter(Boolean);
    const myEv = live.filter(e => (e.data === today || e.data === tomorrow || e.montajData === today || e.montajData === tomorrow) && roles(e).length)
      .sort((a, b) => (a.data + a.start).localeCompare(b.data + b.start));
    const opts = live.filter(e => e.status === STATUS_OPTIUNE && e.responsabil === p.name && e.optiunePana && e.optiunePana <= tomorrow);
    const lipsa = live.filter(e => e.responsabil === p.name && (e.montaj || e.demontaj) && (!e.deschide || !e.inchide) && e.data >= today && e.data <= addDays_(today, 7));
    const anulate = cancelled.filter(e => e.responsabil === p.name);
    const nepreg = live.filter(e => e.responsabil === p.name && e.data >= today && e.data <= addDays_(today, 3))
      .map(e => ({ e: e, r: readiness_(e, allTasks, templates) })).filter(x => x.r.total && x.r.gata < x.r.total);
    if (!late.length && !now.length && !myEv.length && !opts.length && !lipsa.length && !anulate.length && !nepreg.length) return;

    const evLine = e => `<b>${h_(e.titlu)}</b> · ${fmtZi_(e.data)}, ${h_(e.start)}${e.sfarsit ? '–' + h_(e.sfarsit) : ''} · ${h_(e.spatiu)}`;
    const taskLine = t => `${h_(t.sarcina)}<br><span style="color:#6B706C">${byId[t.eventId] ? evLine(byId[t.eventId]) : ''} · termen ${fmtZi_(t.termen)}</span>`;
    const sec = (title, items, color) => items.length ? `<h3 style="font:600 16px Georgia,serif;color:${color || '#24382E'};margin:22px 0 8px">${title}</h3><ul style="padding-left:18px;margin:0">${items.map(x => `<li style="margin:0 0 8px">${x}</li>`).join('')}</ul>` : '';
    let html = `<div style="font:14px/1.5 Arial,sans-serif;color:#1E2420;max-width:560px">`;
    html += `<p style="margin:0 0 4px">Bună dimineața, ${h_(p.name)}.</p><p style="margin:0;color:#6B706C">Ce ai de urmărit azi, ${fmtZi_(today)}, în calendarul Republica | 1963.</p>`;
    html += sec('Nu sunt gata pentru public', nepreg.map(x => evLine(x.e) + `<br><span style="color:#6B706C">gata ${x.r.gata} din ${x.r.total}; lipsește: ${x.r.lipsa.map(h_).join('; ')}</span>`), '#CB2125');
    html += sec('Întârziate' + (late.length > 15 ? ' (primele 15 din ' + late.length + ')' : ''), late.slice(0, 15).map(taskLine), '#CB2125');
    html += sec('De făcut azi și mâine', now.map(taskLine));
    html += sec('Evenimentele tale, azi și mâine', myEv.map(e => {
      const r = roles(e), extra = [];
      if (r.indexOf('montaj') >= 0 && e.montaj) extra.push('montaj de la ' + (e.montajData ? fmtZi_(e.montajData) + ' ' : '') + h_(e.montaj));
      if (r.indexOf('demontaj') >= 0 && e.demontaj) extra.push('demontaj până la ' + h_(e.demontaj));
      if (e.proba) extra.push('probă de sunet ' + h_(e.proba));
      if (e.acces) extra.push('acces public ' + h_(e.acces));
      return evLine(e) + `<br><span style="color:#6B706C">Tu: ${r.join(', ')}${extra.length ? ' · ' + extra.join(' · ') : ''}</span>`;
    }));
    html += sec('Opțiuni care expiră', opts.map(e => evLine(e) + `<br><span style="color:#6B706C">valabilă până ${fmtZi_(e.optiunePana)}. Confirmă clientul sau anulează; după ${ZILE_GRATIE} zile se anulează singură.</span>`), '#B24732');
    html += sec('Lipsește omul de la montaj sau demontaj', lipsa.map(e => evLine(e) + `<br><span style="color:#6B706C">${!e.deschide ? 'montaj: nimeni ales' : ''}${!e.deschide && !e.inchide ? ' · ' : ''}${!e.inchide ? 'demontaj: nimeni ales' : ''}</span>`), '#B24732');
    html += sec('Anulate automat azi-noapte', anulate.map(e => evLine(e) + `<br><span style="color:#6B706C">opțiunea a expirat și n-a fost confirmată</span>`), '#6B706C');
    html += `<p style="margin:26px 0 0"><a href="${APP_URL}" style="background:#24382E;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;display:inline-block">Deschide calendarul</a></p></div>`;
    const parts = [nepreg.length && nepreg.length + ' nepregătite', late.length && late.length + ' întârziate', now.length && now.length + ' de făcut', myEv.length && (myEv.length === 1 ? '1 eveniment' : myEv.length + ' evenimente'),
      opts.length && (opts.length === 1 ? '1 opțiune expiră' : opts.length + ' opțiuni expiră'), anulate.length && anulate.length + ' anulate automat', lipsa.length && 'lipsesc oameni la montaj'].filter(Boolean);
    MailApp.sendEmail({ to: previewTo || p.email, subject: (previewTo ? '[Test, rezumatul lui ' + p.name + '] ' : '') + 'Calendar Republica: ' + (parts.join(', ') || 'rezumatul zilei'), htmlBody: html, name: 'Calendar Republica | 1963' });
    sent++;
  });
  return { sent: sent, cancelled: cancelled.length };
}


/* ---------- Notificări imediate la schimbări importante ---------- */

const CRITICE_ = ['data', 'spatiu', 'start', 'sfarsit', 'montajData', 'montaj', 'proba', 'acces', 'demontaj', 'deschide', 'inchide', 'responsabil', 'status'];

/** Trimite imediat email oamenilor implicați când se schimbă ceva important la un eveniment. */
function notifyChange_(cur, nou) {
  try {
    const arhivat = nou.arhivat === 'DA' && cur.arhivat !== 'DA';
    const diffs = CRITICE_.filter(k => String(cur[k] || '') !== String(nou[k] || ''));
    if (!arhivat && !diffs.length) return;
    if (!arhivat && diffs.length === 1 && diffs[0] === 'status' && nou.status !== STATUS_ANULAT && cur.status !== STATUS_ANULAT) return; // Opțiune → Confirmat nu e o alarmă
    const names = {};
    [cur, nou].forEach(e => ['responsabil', 'deschide', 'inchide'].forEach(k => { if (e[k]) names[e[k]] = 1; }));
    readTable_(SHEET_TASKS, TASK_COLS, TASK_NUM).filter(t => t.eventId === nou.id && t.stare === T_OPEN && t.responsabil).forEach(t => names[t.responsabil] = 1);
    const me = USER_ ? USER_.name : '';
    delete names[me];
    const team = team_().filter(p => names[p.name]);
    if (!team.length) return;
    const val = (e, k) => {
      const v = String(e[k] || '').replace(/^'/, '');
      if (!v) return '—';
      return (k === 'data' || k === 'montajData') ? fmtZi_(v) : v;
    };
    const rows = arhivat ? ['<li><b>Evenimentul a fost scos din calendar (arhivat).</b></li>']
      : diffs.map(k => `<li>${h_(LABELS[k])}: <s style="color:#9C9286">${h_(val(cur, k))}</s> → <b>${h_(val(nou, k))}</b></li>`);
    const anulat = arhivat || (nou.status === STATUS_ANULAT && cur.status !== STATUS_ANULAT);
    const titlu = String(nou.titlu || '').replace(/^'/, '');
    const subj = (anulat ? 'Anulat: ' : 'Modificare: ') + titlu + ' · ' + fmtZi_(nou.data);
    const html = `<div style="font:14px/1.5 Arial,sans-serif;color:#1E2420;max-width:560px">
      <p style="margin:0 0 6px">${h_(me || 'Cineva')} a modificat un eveniment la care ești implicat:</p>
      <h3 style="font:600 17px Georgia,serif;color:${anulat ? '#CB2125' : '#24382E'};margin:10px 0 4px">${h_(titlu)}</h3>
      <p style="margin:0;color:#6B706C">${fmtZi_(nou.data)}, ${h_(nou.start)}${nou.sfarsit ? '–' + h_(nou.sfarsit) : ''} · ${h_(nou.spatiu)}</p>
      <ul style="padding-left:18px">${rows.join('')}</ul>
      <p style="margin:20px 0 0"><a href="${APP_URL}" style="background:#24382E;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;display:inline-block">Deschide calendarul</a></p></div>`;
    team.forEach(p => MailApp.sendEmail({ to: p.email, subject: subj, htmlBody: html, name: 'Calendar Republica | 1963' }));
    log_(me, 'Notificare trimisă', nou.id, titlu, team.map(p => p.name).join(', ') + ': ' + (arhivat ? 'arhivat' : diffs.map(k => LABELS[k]).join(', ')));
  } catch (err) {
    try { log_(USER_ ? USER_.name : '', 'Eroare notificare', nou.id, nou.titlu, String(err && err.message || err)); } catch (_) {}
  }
}

/* ---------- Backup zilnic ---------- */

/** O copie a Sheet-ului în fiecare zi, în folderul „Backup calendar” de lângă el. Nu șterge nimic. */
function backup_(today) {
  try {
    const file = DriveApp.getFileById(SS_ID);
    const parent = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
    const it = parent.getFoldersByName('Backup calendar');
    const folder = it.hasNext() ? it.next() : parent.createFolder('Backup calendar');
    const name = 'Calendar backup ' + today;
    if (folder.getFilesByName(name).hasNext()) return;
    file.makeCopy(name, folder);
  } catch (err) {
    try { log_('Calendar (automat)', 'Eroare backup', '', '', String(err && err.message || err)); } catch (_) {}
  }
}

/** Gata pentru public: sarcinile critice ale evenimentului (marcate „DA” în Șabloane, coloana F). */
function readiness_(ev, tasks, templates) {
  const crit = {}; templates.filter(t => t.critica).forEach(t => crit[t.cod] = 1);
  const ts = tasks.filter(t => t.eventId === ev.id && crit[t.sablon] && t.stare !== T_CANCEL);
  return { total: ts.length, gata: ts.filter(t => t.stare === T_DONE).length, lipsa: ts.filter(t => t.stare !== T_DONE).map(t => t.sarcina) };
}
