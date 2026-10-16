/**
 * The Revival Is Here — registration + Square payment backend.
 * Paste this into the Google Sheet: Extensions → Apps Script.
 * Secrets go in Project Settings → Script properties (never in this file):
 *   SQUARE_ACCESS_TOKEN  = your Square production access token (EAAA...)
 *   SQUARE_LOCATION_ID   = your Square location ID
 *   CHECKIN_PIN          = a PIN your door team types once on the scanner page
 *   NTFY_TOPIC           = (optional) private ntfy.sh topic for phone push alerts
 */
const CAP = 250;
const PRICE_CENTS = 3000; // $30.00
const SHEET_NAME = 'Inscripciones';
const HEADERS = ['Fecha', 'Nombre y apellido', 'Correo electrónico', 'Iglesia', 'Ciudad', 'Monto', 'Estado del pago', 'ID de pago Square', 'Código', 'Check-in'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function paidCount_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return 0;
  return sh.getRange(2, 7, n, 1).getValues().filter(r => r[0] === 'COMPLETED').length;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'checkin') return json_(checkin_(p));
  if (p.action === 'list') return json_(list_(p));
  if (p.action === 'undo') return json_(undo_(p));
  return json_({ count: paidCount_(), cap: CAP });
}

function pinOk_(p) {
  const pin = PropertiesService.getScriptProperties().getProperty('CHECKIN_PIN');
  return !!pin && String(p.pin || '') === pin;
}

/** Staff app: every registration with its check-in status. */
function list_(p) {
  if (!pinOk_(p)) return { ok: false, status: 'pin' };
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return { ok: true, rows: [], cap: CAP };
  const tz = 'America/New_York';
  const fmt = (v, f) => (v ? Utilities.formatDate(new Date(v), tz, f) : '');
  const rows = sh.getRange(2, 1, n, 10).getValues().map(r => ({
    fecha: fmt(r[0], 'MMM d, h:mm a'), nombre: r[1], email: r[2], iglesia: r[3], ciudad: r[4],
    monto: r[5], estado: r[6], codigo: r[8], checkin: fmt(r[9], 'h:mm a')
  }));
  return { ok: true, rows: rows, cap: CAP };
}

/** Staff app: undo a check-in made by mistake. */
function undo_(p) {
  if (!pinOk_(p)) return { ok: false, status: 'pin' };
  const code = String(p.code || '').trim().toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n <= 0) return { ok: false };
    const codes = sh.getRange(2, 9, n, 1).getValues();
    for (let i = 0; i < codes.length; i++) {
      if (String(codes[i][0]).toUpperCase() === code) { sh.getRange(i + 2, 10).setValue(''); return { ok: true }; }
    }
    return { ok: false };
  } finally {
    lock.releaseLock();
  }
}

/** Door check-in: verifies a ticket code and marks it used. */
function checkin_(p) {
  const pin = PropertiesService.getScriptProperties().getProperty('CHECKIN_PIN');
  if (!pin || String(p.pin || '') !== pin) return { ok: false, status: 'pin', error: 'PIN incorrecto' };
  const code = String(p.code || '').trim().toUpperCase();
  if (!code) return { ok: false, status: 'invalid' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n <= 0) return { ok: false, status: 'invalid' };
    const rows = sh.getRange(2, 1, n, 10).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][8]).toUpperCase() !== code) continue;
      const r = rows[i];
      const who = { nombre: r[1], iglesia: r[3], ciudad: r[4] };
      if (r[6] !== 'COMPLETED') return Object.assign({ ok: false, status: 'unpaid' }, who);
      if (r[9]) return Object.assign({ ok: false, status: 'used', at: Utilities.formatDate(new Date(r[9]), 'America/New_York', 'h:mm a') }, who);
      sh.getRange(i + 2, 10).setValue(new Date());
      return Object.assign({ ok: true, status: 'valid' }, who);
    }
    return { ok: false, status: 'invalid' };
  } finally {
    lock.releaseLock();
  }
}

function newCode_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let c = 'AV-';
  for (let i = 0; i < 8; i++) c += chars[Math.floor(Math.random() * chars.length)];
  return c;
}

/** Emails the ticket with a QR code that the door team scans. */
function sendTicket_(nombre, email, code) {
  nombre = String(nombre).replace(/[<>&"]/g, '');
  const qr = UrlFetchApp.fetch('https://quickchart.io/qr?size=400&margin=2&ecLevel=H&dark=8e0f19&text=' + encodeURIComponent(code)).getBlob().setName('entrada.png');
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;text-align:center;color:#222">' +
    '<h2 style="color:#8e0f19;margin-bottom:4px">The Revival Is Here</h2>' +
    '<p style="margin-top:0">Iglesia Presencia de Dios</p>' +
    '<p>Hola <b>' + nombre + '</b>, tu inscripción está confirmada y recibimos tu ofrenda de $30.</p>' +
    '<p>Muestra este código en la entrada:</p>' +
    '<img src="cid:qr" width="260" height="260" alt="Código QR de entrada">' +
    '<p style="font-size:20px;letter-spacing:2px"><b>' + code + '</b></p>' +
    '<p><b>Sábado 14 de noviembre</b><br>13505 Dulles Technology Dr., Suite #2A, Herndon, VA 20171</p>' +
    '<p style="color:#666;font-size:13px">Predicador invitado: Josué Bernier · Almuerzo incluido</p></div>';
  MailApp.sendEmail({
    to: email,
    subject: 'Tu entrada: The Revival Is Here · 14 de noviembre',
    htmlBody: html,
    inlineImages: { qr: qr },
    attachments: [qr],
    name: 'Iglesia Presencia de Dios'
  });
}

/** Alerts the organizer on every paid registration: email to the script owner + optional ntfy push. */
function notifyOwner_(nombre, iglesia, ciudad, count) {
  const title = 'Nueva inscripción (' + count + '/' + CAP + ')';
  const text = nombre + ' · ' + iglesia + ' · ' + ciudad + '\nQuedan ' + (CAP - count) + ' cupos.';
  const owner = Session.getEffectiveUser().getEmail();
  if (owner) MailApp.sendEmail({ to: owner, subject: '✅ ' + title + ' – ' + nombre, body: text, name: 'Revival Inscripciones' });
  const topic = PropertiesService.getScriptProperties().getProperty('NTFY_TOPIC');
  if (topic) UrlFetchApp.fetch('https://ntfy.sh/' + encodeURIComponent(topic), {
    method: 'post', payload: text, muteHttpExceptions: true,
    headers: { Title: title, Tags: 'tada', Priority: 'high' }
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const clean = v => String(v || '').trim().slice(0, 200);
    const nombre = clean(d.nombre), email = clean(d.email), iglesia = clean(d.iglesia), ciudad = clean(d.ciudad);
    if (!nombre || !email || !iglesia || !ciudad || !d.token) return json_({ ok: false, error: 'Faltan datos. Completa todos los campos.' });
    if (paidCount_() >= CAP) return json_({ ok: false, error: 'Lo sentimos, los 250 cupos ya están llenos.' });

    const props = PropertiesService.getScriptProperties();
    const res = UrlFetchApp.fetch('https://connect.squareup.com/v2/payments', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + props.getProperty('SQUARE_ACCESS_TOKEN'), 'Square-Version': '2024-10-17' },
      payload: JSON.stringify({
        source_id: d.token,
        idempotency_key: Utilities.getUuid(),
        amount_money: { amount: PRICE_CENTS, currency: 'USD' },
        location_id: props.getProperty('SQUARE_LOCATION_ID'),
        buyer_email_address: email,
        verification_token: d.verificationToken || undefined,
        note: 'The Revival Is Here - ' + nombre
      })
    });
    const body = JSON.parse(res.getContentText());
    const p = body.payment;
    if (!p || p.status !== 'COMPLETED') {
      const e0 = (body.errors && body.errors[0]) || {};
      const code = e0.code || ('HTTP ' + res.getResponseCode());
      const why = 'ERROR: ' + code + (e0.detail ? ' - ' + e0.detail : '') + (e0.field ? ' (' + e0.field + ')' : '');
      sheet_().appendRow([new Date(), nombre, email, iglesia, ciudad, 30, 'FALLIDO', why.slice(0, 300), '', '']);
      const declined = ['CARD_DECLINED', 'GENERIC_DECLINE', 'INSUFFICIENT_FUNDS', 'CVV_FAILURE', 'ADDRESS_VERIFICATION_FAILURE', 'INVALID_EXPIRATION', 'CARD_EXPIRED'];
      const msg = declined.indexOf(code) >= 0 ? 'Tu tarjeta fue rechazada. Prueba con otra.' : 'No se pudo completar el pago. Intenta de nuevo.';
      return json_({ ok: false, error: msg + ' (' + code + ')' });
    }
    const code = newCode_();
    sheet_().appendRow([new Date(), nombre, email, iglesia, ciudad, 30, 'COMPLETED', p.id, code, '']);
    try { sendTicket_(nombre, email, code); } catch (mailErr) {}
    try { notifyOwner_(nombre, iglesia, ciudad, paidCount_()); } catch (nErr) {}
    return json_({ ok: true, count: paidCount_(), code: code });
  } catch (err) {
    try { sheet_().appendRow([new Date(), '', '', '', '', '', 'ERROR SCRIPT', String(err).slice(0, 300), '', '']); } catch (e2) {}
    return json_({ ok: false, error: 'Ocurrió un error. Intenta de nuevo en un momento.' });
  } finally {
    lock.releaseLock();
  }
}
