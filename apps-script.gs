/**
 * The Revival Is Here — registration + Square payment backend.
 * Paste this into the Google Sheet: Extensions → Apps Script.
 * Secrets go in Project Settings → Script properties (never in this file):
 *   SQUARE_ACCESS_TOKEN  = your Square production access token (EAAA...)
 *   SQUARE_LOCATION_ID   = your Square location ID
 */
const CAP = 250;
const PRICE_CENTS = 3000; // $30.00
const SHEET_NAME = 'Inscripciones';
const HEADERS = ['Fecha', 'Nombre y apellido', 'Correo electrónico', 'Iglesia', 'Ciudad', 'Monto', 'Estado del pago', 'ID de pago Square'];

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

function doGet() {
  return json_({ count: paidCount_(), cap: CAP });
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
        note: 'The Revival Is Here - ' + nombre
      })
    });
    const body = JSON.parse(res.getContentText());
    const p = body.payment;
    if (!p || p.status !== 'COMPLETED') {
      sheet_().appendRow([new Date(), nombre, email, iglesia, ciudad, 30, 'FALLIDO', '']);
      const code = body.errors && body.errors[0] && body.errors[0].code;
      const msg = code === 'CARD_DECLINED' || code === 'GENERIC_DECLINE' ? 'Tu tarjeta fue rechazada. Prueba con otra.' : 'No se pudo completar el pago. Intenta de nuevo.';
      return json_({ ok: false, error: msg });
    }
    sheet_().appendRow([new Date(), nombre, email, iglesia, ciudad, 30, 'COMPLETED', p.id]);
    return json_({ ok: true, count: paidCount_() });
  } catch (err) {
    return json_({ ok: false, error: 'Ocurrió un error. Intenta de nuevo en un momento.' });
  } finally {
    lock.releaseLock();
  }
}
