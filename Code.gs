/**
 * IDP CAFC — Backend (Google Apps Script)
 * Databáze + živé seznamy + ukládání + e-mail + data pro dashboard
 *
 * ► setup()  — jednorázově naplní listy Hráči a Trenéři reálnými daty
 *              (POZOR: přepíše obsah listů Hráči a Trenéři!)
 */

const SHEETS = { PLAYERS: 'Hráči', COACHES: 'Trenéři', RESPONSES: 'Odpovědi' };
const CLUB = 'ČAFC';   // zatím jeden klub

// ─────────────────────────────────────────────────────────────
// SETUP — spusť JEDNOU z editoru (naplní hráče a trenéry)
// ─────────────────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Hráči ---
  let hraci = ss.getSheetByName(SHEETS.PLAYERS) || ss.insertSheet(SHEETS.PLAYERS);
  hraci.clear();
  hraci.getRange(1, 1, 1, 9).setValues([[
    'Jméno', 'Příjmení', 'Přezdívka', 'Datum narození', 'Ročník', 'Tým', 'Pozice', 'Kód', 'Aktivní'
  ]]).setFontWeight('bold');
  const players = [
    ['Albert','Hora','Berťa','8.11.2012',2012,'U15','O','','ano'],
    ['Bohumil','Dráb','Bohy','24.4.2012',2012,'U15','O','','ano'],
    ['Dominik','Hrůša','Domča','24.6.2012',2012,'U15','BR','','ano'],
    ['Filip','Kadlec','Káca','17.11.2012',2012,'U15','O','','ano'],
    ['Filip','Šácha','Šachy','31.1.2012',2012,'U15','BR','','ano'],
    ['Filip','Šandor','Šandy','21.11.2012',2012,'U15','O','','ano'],
    ['Jakub','Jeníček','Jéňa','3.11.2012',2012,'U15','Ú','','ano'],
    ['Jakub','Sláma','Slámič','13.6.2012',2012,'U15','SZ','','ano'],
    ['Jan','Navrátil','Navi','1.6.2012',2012,'U15','SZ','','ano'],
    ['Jeroným','Fiala','Jeroš','10.7.2012',2012,'U15','Ú','','ano'],
    ['Lukáš','Mareš','Luky','16.9.2012',2012,'U15','Ú','','ano'],
    ['Matyáš','Urban','Urbi','3.6.2012',2012,'U15','KZ','','ano'],
    ['Matyáš','Zimandl','Zimi','13.10.2012',2012,'U15','SZ','','ano'],
    ['Nicolas','Steranka','Nico','6.6.2012',2012,'U15','O','','ano'],
    ['Ondřej','Lipavský','Lipi','15.3.2013',2013,'U15','KZ','','ano'],
    ['Šimon','Mádr','Šíma','31.7.2012',2012,'U15','SZ','','ano'],
    ['Štěpán','Hodr','Hody','12.12.2013',2013,'U15','Ú','','ano'],
    ['Tobiáš','Kostinec','Kosťa','29.8.2012',2012,'U15','O','','ano'],
    ['Tobiáš','Mareš','Toby (Skála)','16.9.2012',2012,'U15','O','','ano'],
    ['Vojtěch','Čanda','Vojta','4.8.2012',2012,'U15','BR','','ano'],
  ];
  hraci.getRange(2, 1, players.length, 9).setValues(players);
  hraci.setFrozenRows(1);
  hraci.setColumnWidth(3, 120); hraci.setColumnWidth(4, 120); hraci.setColumnWidth(8, 200);

  // --- Trenéři ---
  let tren = ss.getSheetByName(SHEETS.COACHES) || ss.insertSheet(SHEETS.COACHES);
  tren.clear();
  tren.getRange(1, 1, 1, 5).setValues([[
    'Jméno', 'Tým', 'Role', 'ID FAČR', 'Aktivní'
  ]]).setFontWeight('bold');
  const coaches = [
    ['Avetis Švamberk','U15','hlavní','90110504','ano'],
    ['Radka Drábová','U15','asistent','80050443','ano'],
    ['Robert Tlustý','U15','asistent','93090846','ano'],
  ];
  tren.getRange(2, 1, coaches.length, 5).setValues(coaches);
  tren.setFrozenRows(1);
  tren.setColumnWidth(4, 120);

  // --- Odpovědi (nezasahujeme, jen zajistíme existenci) ---
  let odp = ss.getSheetByName(SHEETS.RESPONSES) || ss.insertSheet(SHEETS.RESPONSES);
  if (odp.getLastRow() === 0) {
    odp.getRange(1, 1, 1, 12).setValues([[
      'Časová značka', 'Typ', 'ID hodnocení', 'Hráč', 'Ročník',
      'Vyplnil', 'Pozice hlavní', 'Pozice druhá', 'Nejlepší pozice',
      'Období', 'Sezóna', 'Data (JSON)',
    ]]).setFontWeight('bold');
    odp.setFrozenRows(1);
    odp.setColumnWidth(12, 400);
  }

  SpreadsheetApp.getUi().alert('Hotovo! Nahráno ' + players.length + ' hráčů a ' + coaches.length + ' trenéři.');
}

// ─────────────────────────────────────────────────────────────
// doGet — seznamy (formuláře) + data (dashboard). Podpora JSONP.
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'lists';
  let result;
  try {
    if (action === 'lists') {
      result = { ok: true, players: getPlayers(), coaches: getCoachNames() };
    } else if (action === 'data') {
      result = { ok: true, club: CLUB, players: getPlayers(), coaches: getCoachesFull(), responses: getResponses() };
    } else {
      result = { ok: false, error: 'Neznámá akce: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return respond(result, e);
}

// ─────────────────────────────────────────────────────────────
// doPost — uložení formuláře + odeslání e-mailu
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'email') {
      MailApp.sendEmail({ to: data.to, subject: data.subject || 'Tvůj IDP výsledek', htmlBody: data.html || '' });
      return respond({ ok: true, emailed: true }, e);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const odp = ss.getSheetByName(SHEETS.RESPONSES);
    const fullData = {
      answers: data.answers || {}, open: data.open || {}, fun: data.fun || {}, verbal: data.verbal || '', xp: data.xp || 0,
    };
    odp.appendRow([
      new Date(), data.type || '', data.assessmentId || '', data.playerName || '', data.playerYear || '',
      data.filledBy || '', data.posMain || '', data.posSecond || '', data.bestPos || '',
      data.period || '', data.season || '', JSON.stringify(fullData),
    ]);
    result = { ok: true, saved: true };
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return respond(result, e);
}

// ─────────────────────────────────────────────────────────────
// POMOCNÉ
// ─────────────────────────────────────────────────────────────
function getPlayers() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PLAYERS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  return rows
    .filter(r => r[0] && String(r[8]).toLowerCase() !== 'ne')
    .map(r => ({
      first: r[0], last: r[1], nick: r[2], dob: r[3], year: r[4],
      team: r[5], pos: r[6], hasCode: !!String(r[7]).trim(),
    }));
}

function getCoachNames() {  // pro formuláře (jen jména)
  return getCoachesFull().map(c => c.name);
}
function getCoachesFull() {  // pro dashboard (role, tým, id)
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.COACHES);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return rows
    .filter(r => r[0] && String(r[4]).toLowerCase() !== 'ne')
    .map(r => ({ name: r[0], team: r[1], role: r[2], id: r[3] }));
}

function getResponses() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RESPONSES);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  return rows.filter(r => r[1]).map(r => {
    let data = {};
    try { data = JSON.parse(r[11] || '{}'); } catch (e) { data = {}; }
    return {
      ts: r[0], type: r[1], assessmentId: r[2], player: r[3], year: r[4],
      filledBy: r[5], posMain: r[6], posSecond: r[7], bestPos: r[8],
      period: r[9], season: r[10], data: data,
    };
  });
}

function respond(obj, e) {
  const json = JSON.stringify(obj);
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
