const SHEET_DIPENDENTI = 'Dipendenti';
const SHEET_TIMBRATURE = 'Timbrature';
const SHEET_CAUSALI = 'Causali';
const SHEET_REGOLE = 'Regole';

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.recordId = (e && e.parameter && e.parameter.recordId) ? String(e.parameter.recordId) : '';

  return tpl.evaluate()
    .setTitle('Timbrature')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSessionUser_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error(
      'Utente non riconosciuto. Pubblica la Web App come "Esegui come: Utente che accede" ' +
      'e "Chi ha accesso: Solo dominio" (o utenti autenticati), poi accedi con account Google aziendale.'
    );
  }
  return email.toLowerCase().trim();
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Foglio mancante: ${name}`);
  return sh;
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defs = [
    {
      name: SHEET_DIPENDENTI,
      headers: ['id_dipendente', 'nome', 'email', 'ruolo', 'attivo', 'responsabile_email']
    },
    {
      name: SHEET_TIMBRATURE,
      headers: [
        'id_record', 'data', 'email_dipendente', 'id_dipendente', 'nome_dipendente',
        'entrata', 'uscita', 'pausa_minuti', 'causale', 'note',
        'ore_lavorate', 'ore_ordinarie', 'ore_straordinario', 'stato', 'responsabile_email',
        'creato_il', 'aggiornato_il', 'approvato_da', 'approvato_il', 'motivo_rifiuto'
      ]
    },
    {
      name: SHEET_CAUSALI,
      headers: ['codice', 'descrizione', 'tipo', 'richiede_note', 'attiva']
    },
    {
      name: SHEET_REGOLE,
      headers: ['chiave', 'valore']
    }
  ];

  defs.forEach(def => {
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sh.setFrozenRows(1);
    }
  });

  const regole = getSheet_(SHEET_REGOLE);
  if (regole.getLastRow() === 1) {
    regole.getRange(2, 1, 4, 2).setValues([
      ['ARROTONDAMENTO_MINUTI', 15],
      ['SOGLIA_ORDINARIO_GIORNO_ORE', 8],
      ['PAUSA_MIN_OBBLIGATORIA_MIN', 0],
      ['TIMEZONE', 'Europe/Rome']
    ]);
  }

  const causali = getSheet_(SHEET_CAUSALI);
  if (causali.getLastRow() === 1) {
    causali.getRange(2, 1, 8, 5).setValues([
      ['LAV', 'Lavoro ordinario', 'LAVORATIVA', 'NO', 'SI'],
      ['FER', 'Ferie', 'ASSENZA', 'NO', 'SI'],
      ['MAL', 'Malattia', 'ASSENZA', 'SI', 'SI'],
      ['PIO', 'Pioggia', 'ASSENZA', 'SI', 'SI'],
      ['PER', 'Permesso', 'ASSENZA', 'SI', 'SI'],
      ['TRA', 'Trasferta', 'LAVORATIVA', 'NO', 'SI'],
      ['FES', 'Festivo', 'ASSENZA', 'NO', 'SI'],
      ['SW', 'Smart working', 'LAVORATIVA', 'NO', 'SI']
    ]);
  }

  return { ok: true, message: 'Setup completato. Verifica e popola il foglio Dipendenti.' };
}

function getRulesMap_() {
  const sh = getSheet_(SHEET_REGOLE);
  const data = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    if (!k) continue;
    map[k] = data[i][1];
  }
  return map;
}

function getEmployeeByEmail_(email) {
  const sh = getSheet_(SHEET_DIPENDENTI);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[2]).toLowerCase().trim() === email && String(r[4]).toUpperCase() === 'SI') {
      return {
        id: r[0],
        nome: r[1],
        email: r[2],
        ruolo: String(r[3] || '').toLowerCase(),
        attivo: r[4],
        responsabileEmail: r[5] || ''
      };
    }
  }
  throw new Error('Dipendente non trovato o non attivo.');
}

function getActiveCausali() {
  const sh = getSheet_(SHEET_CAUSALI);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).toUpperCase() === 'SI') {
      out.push({
        codice: rows[i][0],
        descrizione: rows[i][1],
        tipo: rows[i][2],
        richiedeNote: String(rows[i][3]).toUpperCase() === 'SI'
      });
    }
  }
  return out;
}

function parseTimeToMinutes_(hhmm) {
  if (!hhmm) return null;
  const parts = String(hhmm).split(':').map(Number);
  if (parts.length !== 2) return null;
  if (Number.isNaN(parts[0]) || Number.isNaN(parts[1])) return null;
  return parts[0] * 60 + parts[1];
}

function roundMinutes_(minutes, step) {
  if (!step || step <= 0) return minutes;
  return Math.round(minutes / step) * step;
}

function computeHours_(entrata, uscita, pausaMinuti, causale, regole) {
  if (causale && causale !== 'LAV' && causale !== 'SW' && causale !== 'TRA') {
    return { worked: 0, ord: 0, stra: 0 };
  }

  const inMin = parseTimeToMinutes_(entrata);
  const outMin = parseTimeToMinutes_(uscita);
  if (inMin == null || outMin == null) return { worked: 0, ord: 0, stra: 0 };

  let workedMin = outMin - inMin - (Number(pausaMinuti) || 0);
  if (workedMin < 0) workedMin = 0;

  const rounded = roundMinutes_(workedMin, Number(regole.ARROTONDAMENTO_MINUTI || 0));
  const sogliaMin = Number(regole.SOGLIA_ORDINARIO_GIORNO_ORE || 8) * 60;

  const ordMin = Math.min(rounded, sogliaMin);
  const straMin = Math.max(rounded - sogliaMin, 0);

  return {
    worked: +(rounded / 60).toFixed(2),
    ord: +(ordMin / 60).toFixed(2),
    stra: +(straMin / 60).toFixed(2)
  };
}

function nowIso_() {
  const tz = getRulesMap_().TIMEZONE || Session.getScriptTimeZone() || 'Europe/Rome';
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
}

function todayIso_() {
  const tz = getRulesMap_().TIMEZONE || Session.getScriptTimeZone() || 'Europe/Rome';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

function getDashboard() {
  const email = getSessionUser_();
  const user = getEmployeeByEmail_(email);
  return {
    user,
    causali: getActiveCausali(),
    today: todayIso_()
  };
}

function createOrUpdateDay(payload) {
  const email = getSessionUser_();
  const user = getEmployeeByEmail_(email);
  const regole = getRulesMap_();

  const data = payload.data || todayIso_();
  const entrata = payload.entrata || '';
  const uscita = payload.uscita || '';
  const pausaMinuti = Number(payload.pausaMinuti || 0);
  const causale = payload.causale || 'LAV';
  const note = payload.note || '';
  const stato = payload.stato || 'BOZZA';

  const calc = computeHours_(entrata, uscita, pausaMinuti, causale, regole);
  const sh = getSheet_(SHEET_TIMBRATURE);
  const rows = sh.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(data) && String(rows[i][2]).toLowerCase().trim() === email) {
      rowIndex = i + 1;
      break;
    }
  }

  const record = [
    '', data, email, user.id, user.nome,
    entrata, uscita, pausaMinuti, causale, note,
    calc.worked, calc.ord, calc.stra,
    stato, user.responsabileEmail || '',
    '', nowIso_(), '', '', ''
  ];

  if (rowIndex === -1) {
    record[0] = Utilities.getUuid();
    record[15] = nowIso_();
    sh.appendRow(record);
  } else {
    const existing = sh.getRange(rowIndex, 1, 1, 20).getValues()[0];
    if (String(existing[13]).toUpperCase() === 'APPROVATO' && user.ruolo === 'dipendente') {
      throw new Error('Record approvato: non modificabile dal dipendente.');
    }

    record[0] = existing[0] || Utilities.getUuid();
    record[15] = existing[15] || nowIso_();
    record[17] = existing[17] || '';
    record[18] = existing[18] || '';
    record[19] = existing[19] || '';
    sh.getRange(rowIndex, 1, 1, 20).setValues([record]);
  }

  return { ok: true, calcolo: calc };
}

function getMyMonthData(yyyyMm) {
  const email = getSessionUser_();
  const sh = getSheet_(SHEET_TIMBRATURE);
  const rows = sh.getDataRange().getValues();
  const out = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const dateVal = String(r[1] || '');
    if (String(r[2]).toLowerCase().trim() === email && dateVal.startsWith(yyyyMm)) {
      out.push({
        id: r[0],
        data: r[1],
        entrata: r[5],
        uscita: r[6],
        pausa: r[7],
        causale: r[8],
        note: r[9],
        ore: r[10],
        ord: r[11],
        stra: r[12],
        stato: r[13]
      });
    }
  }

  return out;
}

function getPendingForManager(yyyyMm) {
  const email = getSessionUser_();
  const me = getEmployeeByEmail_(email);
  if (!['responsabile', 'admin'].includes(me.ruolo)) return [];

  const sh = getSheet_(SHEET_TIMBRATURE);
  const rows = sh.getDataRange().getValues();
  const out = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const data = String(r[1] || '');
    const stato = String(r[13] || '').toUpperCase();
    const resp = String(r[14] || '').toLowerCase().trim();

    if (data.startsWith(yyyyMm) && stato === 'INVIATO') {
      if (me.ruolo === 'admin' || resp === email) {
        out.push({
          id: r[0],
          data: r[1],
          dipendente: r[4],
          email: r[2],
          causale: r[8],
          ore: r[10],
          ord: r[11],
          stra: r[12],
          stato: r[13],
          note: r[9]
        });
      }
    }
  }

  return out;
}

function getMyRecordDetail(recordId) {
  const email = getSessionUser_();
  const me = getEmployeeByEmail_(email);
  const sh = getSheet_(SHEET_TIMBRATURE);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]) !== String(recordId)) continue;

    const ownerEmail = String(r[2] || '').toLowerCase().trim();
    const responsabileEmail = String(r[14] || '').toLowerCase().trim();
    const isOwner = ownerEmail === email;
    const isAdmin = me.ruolo === 'admin';
    const isAssignedManager = me.ruolo === 'responsabile' && responsabileEmail === email;

    if (!isOwner && !isAdmin && !isAssignedManager) {
      throw new Error('Non autorizzato a visualizzare questa timbratura.');
    }

    return {
      id: r[0],
      data: r[1],
      emailDipendente: r[2],
      idDipendente: r[3],
      nomeDipendente: r[4],
      entrata: r[5],
      uscita: r[6],
      pausa: r[7],
      causale: r[8],
      note: r[9],
      ore: r[10],
      ord: r[11],
      stra: r[12],
      stato: r[13],
      responsabileEmail: r[14],
      creatoIl: r[15],
      aggiornatoIl: r[16],
      approvatoDa: r[17],
      approvatoIl: r[18],
      motivoRifiuto: r[19]
    };
  }

  throw new Error('Timbratura non trovata.');
}

function approveRecord(idRecord, approve, motivo) {
  const email = getSessionUser_();
  const me = getEmployeeByEmail_(email);
  if (!['responsabile', 'admin'].includes(me.ruolo)) {
    throw new Error('Non autorizzato.');
  }

  const sh = getSheet_(SHEET_TIMBRATURE);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(idRecord)) continue;

    const responsabileEmail = String(rows[i][14] || '').toLowerCase().trim();
    if (me.ruolo === 'responsabile' && responsabileEmail && responsabileEmail !== email) {
      throw new Error('Puoi approvare solo timbrature dei tuoi collaboratori.');
    }

    rows[i][13] = approve ? 'APPROVATO' : 'RIFIUTATO';
    rows[i][17] = email;
    rows[i][18] = nowIso_();
    rows[i][19] = approve ? '' : (motivo || 'Rifiutato');
    sh.getRange(i + 1, 1, 1, 20).setValues([rows[i]]);

    return { ok: true };
  }

  throw new Error('Record non trovato.');
}
