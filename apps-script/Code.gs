const CONFIG = {
  SHEET_F3: 'Fase 3 (2025-)',
  SHEET_F2: 'Fase 2 (2019-2024)',
  SHEET_F1: 'Fase 1 (2018-2019)',
  SHEET_EV: 'Eventos',
  SHEET_TB: 'Taboom',
  SHEET_FAV: 'Jogos Favoritos',
  CACHE_SECONDS: 300
};

const DIAS = ['Domingo', 'Segunda-Feira', 'Terça-Feira', 'Quarta-Feira', 'Quinta-Feira', 'Sexta-Feira', 'Sábado'];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'data';
    if (action === 'ping') return json_({ ok: true, ts: new Date().toISOString() });
    const fresh = e && e.parameter && e.parameter.fresh === '1';
    return ContentService.createTextOutput(getData(fresh)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const user = verifyIdToken_(body.idToken);

    if (body.action === 'whoami') return json_({ ok: true, email: user.email, admin: user.admin });

    if (body.action === 'add') {
      if (!user.admin) throw new Error('Só o Ludo pode inserir jogos. GAME OVER 💀');
      return json_(addGame_(body.game || {}, user.email));
    }
    throw new Error('Ação desconhecida');
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('Faça login com o Google primeiro.');
  const clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (!clientId) throw new Error('GOOGLE_CLIENT_ID não configurado nas Script Properties.');

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Login inválido ou expirado. Entre de novo.');
  const t = JSON.parse(res.getContentText());

  if (t.aud !== clientId) throw new Error('Token emitido para outro site.');
  if (t.iss !== 'accounts.google.com' && t.iss !== 'https://accounts.google.com') throw new Error('Emissor inválido.');
  if (Number(t.exp) * 1000 < Date.now()) throw new Error('Login expirado. Entre de novo.');
  if (String(t.email_verified) !== 'true') throw new Error('Email não verificado.');

  const email = String(t.email || '').toLowerCase();
  const admins = String(PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '')
    .split(/[,;\s]+/).map(a => a.trim().toLowerCase()).filter(String);
  return { email: email, admin: admins.indexOf(email) > -1 };
}

function authorize() {
  ss_().getName();
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo', { muteHttpExceptions: true });
  CacheService.getScriptCache().get('x');
  LockService.getScriptLock();
  Logger.log('Autorizado! Planilha: ' + ss_().getName());
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getData(fresh) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('ludo_n') || 0);
  const builtAt = Number(cache.get('ludo_t') || 0);
  const skipCache = fresh && Date.now() - builtAt > 20000;
  if (n && !skipCache) {
    const keys = []; for (let i = 0; i < n; i++) keys.push('ludo_' + i);
    const parts = cache.getAll(keys);
    if (keys.every(k => parts[k] !== undefined)) return keys.map(k => parts[k]).join('');
  }
  const json = JSON.stringify(buildData_());
  try {
    const map = {}, SIZE = 90000; let i = 0;
    for (let p = 0; p < json.length; p += SIZE) map['ludo_' + (i++)] = json.slice(p, p + SIZE);
    map.ludo_n = String(i);
    map.ludo_t = String(Date.now());
    cache.putAll(map, CONFIG.CACHE_SECONDS);
  } catch (e) {  }
  return json;
}

function buildData_() {
  const ss = ss_();
  const tz = ss.getSpreadsheetTimeZone();
  const games = [];

  [[CONFIG.SHEET_F3, 'F3'], [CONFIG.SHEET_F2, 'F2']].forEach(([name, tag]) => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    sh.getRange(1, 1, sh.getLastRow(), 14).getValues().forEach(r => {
      const n = clean_(r[1]);
      if (!n || n === 'Jogo' || /^\d{4}\/\d{4}/.test(n)) return;
      games.push(row_(tag, n, r[4], r[6], r[7], r[8], r[11], r[12], r[13], tz));
    });
  });

  (function () {
    const sh = ss.getSheetByName(CONFIG.SHEET_F1); if (!sh) return;
    sh.getRange(1, 1, sh.getLastRow(), 13).getValues().forEach(r => {
      const n = clean_(r[1]);
      if (!n || n === 'Jogo' || /^\d{4}\/\d{4}/.test(n)) return;
      const notes = [clean_(r[8]), clean_(r[9])].filter(String).join(' • ');
      games.push(row_('F1', n, r[4], r[6], r[7], notes, 'LIVE - TWITCH', r[11], '', tz));
    });
  })();

  (function () {
    const sh = ss.getSheetByName(CONFIG.SHEET_EV); if (!sh) return;
    sh.getRange(1, 1, sh.getLastRow(), 13).getValues().forEach(r => {
      const n = clean_(r[1]);
      if (!n || n === 'Jogo' || /^Eventos/.test(n)) return;
      games.push(row_('EV', n, r[4], r[6], r[7], r[8], 'EVENTO', r[11], r[12], tz));
    });
  })();

  (function () {
    const sh = ss.getSheetByName(CONFIG.SHEET_TB); if (!sh) return;
    sh.getRange(1, 1, sh.getLastRow(), 6).getValues().forEach(r => {
      const n = clean_(r[1]);
      if (!n || n === 'Jogo' || /^\d{4}\/\d{4}/.test(n)) return;
      games.push(row_('TB', n, r[5], '', '', 'Live na Taboom (2021/2022)', 'TABOOM', '', '', tz));
    });
  })();

  const anual = [];
  const f3 = ss.getSheetByName(CONFIG.SHEET_F3);
  if (f3) {
    f3.getRange(1, 16, 40, 4).getValues().forEach(r => {
      if (typeof r[0] === 'number' && r[0] > 2000) anual.push({ ano: r[0], total: num_(r[1]), unicos: num_(r[2]), off: num_(r[3]) });
    });
  }

  const favoritos = [];
  const fav = ss.getSheetByName(CONFIG.SHEET_FAV);
  if (fav) {
    fav.getRange(2, 1, Math.max(fav.getLastRow() - 1, 1), 6).getValues().forEach(r => {
      const n = clean_(r[0]); if (n) favoritos.push({ n: n, s: clean_(r[3]), y: num_(r[5]) });
    });
  }

  return { updated: new Date().toISOString(), games: games, anual: anual, favoritos: favoritos };
}

function row_(tag, n, sis, tempo, nota, notes, onde, data, dia, tz) {
  const t = tempo_(tempo, tz);
  const d = data_(data, tz);
  let o = String(onde || '').toUpperCase();
  o = o.indexOf('OFF') > -1 ? 'OFF' : o.indexOf('TABOOM') > -1 ? 'TABOOM' : o.indexOf('EVENTO') > -1 ? 'EVENTO' : 'LIVE';
  return {
    f: tag, n: n, s: clean_(sis), h: t.h, t: t.txt, r: nota_(nota), x: clean_(notes),
    o: o, d: d, w: clean_(dia).replace(/[\[\]]/g, '') || (d.length === 10 ? DIAS[new Date(d + 'T12:00:00').getDay()] : '')
  };
}

function clean_(v) { return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim(); }
function num_(v) { return typeof v === 'number' ? v : null; }

function nota_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').toUpperCase();
  if (s.indexOf('S') > -1 || s.indexOf('11') > -1) return 11;
  const n = parseFloat(s); return isNaN(n) ? null : n;
}

function tempo_(v, tz) {
  if (v instanceof Date) {
    const s = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
    const p = s.split(' '), dp = p[0].split('-').map(Number), tp = p[1].split(':').map(Number);
    const days = Math.round((Date.UTC(dp[0], dp[1] - 1, dp[2]) - Date.UTC(1899, 11, 30)) / 864e5);
    const h = Math.max(0, days) * 24 + tp[0] + tp[1] / 60 + tp[2] / 3600;
    return { h: Math.round(h * 100) / 100, txt: fmtH_(h) };
  }
  if (typeof v === 'number') { const h = v * 24; return { h: Math.round(h * 100) / 100, txt: fmtH_(h) }; }
  const m = String(v || '').match(/^(\d+):(\d{2})(?::(\d{2}))?$/);
  if (m) { const h = +m[1] + m[2] / 60 + (m[3] ? m[3] / 3600 : 0); return { h: Math.round(h * 100) / 100, txt: fmtH_(h) }; }
  return { h: null, txt: '' };
}

function fmtH_(h) { let H = Math.floor(h), M = Math.round((h - H) * 60); if (M === 60) { H++; M = 0; } return H + 'h' + (M ? ('0' + M).slice(-2) : ''); }

function data_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (typeof v === 'number') return String(v);
  return clean_(v);
}

function addGame_(g, email) {
  const jogo = clean_(g.jogo).slice(0, 150);
  if (!jogo) throw new Error('Faltou o nome do jogo!');
  const notaN = Math.round(Number(g.nota));
  if (!(notaN >= 1 && notaN <= 11)) throw new Error('Nota inválida.');
  if (g.data && !/^\d{4}-\d{2}-\d{2}$/.test(g.data)) throw new Error('Data inválida.');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const ss = ss_();
    const sh = ss.getSheetByName(CONFIG.SHEET_F3);
    const B = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();

    let first = -1;
    for (let i = 1; i < B.length; i++) if (clean_(B[i][0])) { first = i + 1; break; }
    if (first === -1) first = sh.getLastRow() + 1;

    let row = first - 1;
    if (row < 2 || clean_(sh.getRange(row, 2).getValue())) {

      sh.getRange('A2:N2').insertCells(SpreadsheetApp.Dimension.ROWS);
      row = 2; first = 3;
    }
    if (!sh.getRange(row, 1).getFormula()) sh.getRange(row, 1).setFormula('=A' + (row + 1) + '+1');

    const dt = g.data ? new Date(g.data + 'T12:00:00') : new Date();
    const tempo = normTempo_(g.tempo);

    sh.getRange(first, 2, 1, 13).copyFormatToRange(sh, 2, 14, row, row);
    sh.getRange(row, 2).setValue(jogo);
    sh.getRange(row, 5).setValue(clean_(g.sistema).slice(0, 60));
    if (tempo) sh.getRange(row, 7).setValue(tempo);
    sh.getRange(row, 8).setValue(notaN === 11 ? '11/S' : notaN);
    sh.getRange(row, 9).setValue(clean_(g.notas).slice(0, 250));
    sh.getRange(row, 12).setValue(g.onde === 'OFF' ? 'OFF STREAM' : 'LIVE - TWITCH');
    sh.getRange(row, 13).setValue(new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()));
    sh.getRange(row, 14).setValue('[' + DIAS[dt.getDay()] + ']');

    if (g.placar !== false) bumpPlacar_(sh, dt.getFullYear(), jogo, g.onde === 'OFF');

    CacheService.getScriptCache().remove('ludo_n');
    console.log('Jogo inserido por ' + email + ': ' + jogo + ' (linha ' + row + ')');
    return { ok: true, row: row, numero: sh.getRange(row, 1).getValue() };
  } finally {
    lock.releaseLock();
  }
}

function normTempo_(t) {
  const m = String(t || '').trim().match(/^(\d+)(?:[:hH](\d{1,2}))?(?::(\d{1,2}))?m?$/);
  if (!m) return '';
  return m[1] + ':' + ('0' + (m[2] || 0)).slice(-2) + ':' + ('0' + (m[3] || 0)).slice(-2);
}

function bumpPlacar_(sh, ano, jogo, off) {
  const v = sh.getRange(1, 16, 40, 4).getValues();
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] === ano) {
      const r = i + 1;
      if (typeof v[i][1] === 'number') sh.getRange(r, 17).setValue(v[i][1] + 1);
      if (typeof v[i][2] === 'number' && !jaZerado_(jogo)) sh.getRange(r, 18).setValue(v[i][2] + 1);
      if (off) sh.getRange(r, 19).setValue(typeof v[i][3] === 'number' ? v[i][3] + 1 : 1);
      return;
    }
  }
}

function jaZerado_(jogo) {
  const key = String(jogo).toLowerCase().replace(/[^a-z0-9]/g, '');
  return buildData_().games.filter(g => g.n.toLowerCase().replace(/[^a-z0-9]/g, '') === key).length > 1;
}
