// Check — servidor en Google Apps Script.
// Guarda tus datos en Drive (carpeta "Check" > check_data.json), lee fotos con el OCR de Google
// y revisa tu Gmail para agregar compras y transferencias.
//
// Implementar como: Aplicación web · Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario.
// La app solo funciona con la CLAVE de abajo, así que nadie más puede leer ni cambiar tus datos.

const TOKEN = '__TOKEN__';
const FOLDER_NAME = 'Check';
const DATA_FILE = 'check_data.json';
const INBOX_FILE = 'check_inbox.json';
const OWNER_RE = /garc[ií]a\s*-?\s*huidobro/i;   // tu nombre: transferencias a/desde ti mismo son traspasos
const BANK_SENDERS = 'from:bci.cl OR from:bancochile.cl OR from:mercadopago.cl OR from:mercadopago.com OR from:santander.cl OR from:bancoestado.cl OR from:scotiabank.cl OR from:itau.cl OR from:bice.cl OR from:security.cl OR from:bancofalabella.cl OR from:machbank.cl OR from:tenpo.cl OR from:bancoripley.cl OR from:coopeuch.cl OR from:bancointernacional.cl OR from:consorcio.cl';
function mailQueries_(days) {
  return [
    'from:contacto@bci.cl subject:(uso de tu tarjeta) newer_than:' + days + 'd',
    'newer_than:' + days + 'd (' + BANK_SENDERS + ') (transferencia OR transferido OR transferiste OR abono)'
  ];
}
const COLLECTIONS = ['transactions', 'recurring', 'installments', 'receivables', 'investments'];

function doGet() {
  return ContentService.createTextOutput('Check API funcionando.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!req || req.token !== TOKEN) return json_({ ok: false, error: 'unauthorized' });
  try {
    switch (req.action) {
      case 'load':
        return json_({ ok: true, data: readFile_(DATA_FILE), inbox: JSON.parse(readFile_(INBOX_FILE) || '[]') });
      case 'save':
        return json_({ ok: true, data: saveMerged_(JSON.parse(req.data)), savedAt: Date.now() });
      case 'ocr':
        return json_({ ok: true, text: ocr_(req.name || 'imagen', req.mimeType || 'image/jpeg', req.b64) });
      case 'checkMail':
        return json_({ ok: true, added: checkMail() });
      case 'rescanMail':
        withLock_(() => writeRaw_(INBOX_FILE, '[]'));
        return json_({ ok: true, added: checkMail(Math.min(90, Math.max(1, +req.days || 7))) });
      default:
        return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Archivos en Drive ----------
function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function file_(name) {
  const it = folder_().getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}
function readFile_(name) {
  const f = file_(name);
  return f ? f.getBlob().getDataAsString() : null;
}
function writeRaw_(name, text) {
  const f = file_(name);
  if (f) f.setContent(text); else folder_().createFile(name, text, MimeType.PLAIN_TEXT);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- Mezcla de datos entre aparatos ----------
// Cada registro lleva "u" (cuándo se cambió). Gana el más reciente; lo borrado queda en "deleted".
function mergeStates_(a, b) {
  a = a || {}; b = b || {};
  const out = Object.assign({}, a, b);
  const deleted = Object.assign({}, a.deleted || {});
  Object.keys(b.deleted || {}).forEach(id => { deleted[id] = Math.max(deleted[id] || 0, b.deleted[id]); });
  COLLECTIONS.forEach(k => {
    const map = {};
    (a[k] || []).forEach(r => { map[String(r.id)] = r; });
    (b[k] || []).forEach(r => { const id = String(r.id), cur = map[id]; if (!cur && !r.u && (a[k] || []).length) return; if (!cur || (r.u || 0) >= (cur.u || 0)) map[id] = r; });  // registros antiguos sin fecha solo sobreviven si el otro lado también los tiene
    out[k] = Object.keys(map).map(id => map[id]).filter(r => !(deleted[String(r.id)] >= (r.u || 0)));
  });
  out.categories = ((b.catsU || 0) >= (a.catsU || 0) ? b.categories : a.categories) || a.categories || b.categories;
  out.catsU = Math.max(a.catsU || 0, b.catsU || 0);
  out.mailSeen = Array.from(new Set((a.mailSeen || []).concat(b.mailSeen || []))).slice(-1500);
  out.deleted = deleted;
  out.updatedAt = Math.max(a.updatedAt || 0, b.updatedAt || 0);
  return out;
}

function saveMerged_(incoming) {
  return withLock_(() => {
    const stored = JSON.parse(readFile_(DATA_FILE) || 'null');
    const merged = stored ? mergeStates_(stored, incoming) : incoming;
    const text = JSON.stringify(merged);
    writeRaw_(DATA_FILE, text);
    return text;
  });
}

// ---------- OCR de fotos y PDF escaneados ----------
function ocr_(name, mimeType, b64) {
  if (typeof Drive === 'undefined' || !Drive.Files) throw new Error('DRIVE_API_OFF');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, name);
  const folderId = folder_().getId();
  let id;
  if (typeof Drive.Files.create === 'function') {
    id = Drive.Files.create({ name: 'check_ocr_tmp', mimeType: 'application/vnd.google-apps.document', parents: [folderId] }, blob, { ocrLanguage: 'es' }).id;
  } else {
    id = Drive.Files.insert({ title: 'check_ocr_tmp', mimeType: 'application/vnd.google-apps.document', parents: [{ id: folderId }] }, blob, { ocr: true, ocrLanguage: 'es' }).id;
  }
  try {
    return DocumentApp.openById(id).getBody().getText();
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  }
}

// ---------- Correo: compras y transferencias ----------
// Se ejecuta solo cada 15 minutos y deja los movimientos nuevos en check_inbox.json.
// La app los agrega al abrirse, con un id fijo por correo, así nunca se duplican.
function checkMail(days) {
  days = typeof days === 'number' ? days : 7;
  return withLock_(() => {
    const inbox = JSON.parse(readFile_(INBOX_FILE) || '[]');
    const known = {}, keys = {};
    inbox.forEach(x => { known[x.mailId] = true; if (x.key) keys[x.key] = true; });
    let added = 0;
    const seenThreads = {};
    mailQueries_(days).forEach(q => {
      GmailApp.search(q, 0, 200).forEach(thread => {
        if (seenThreads[thread.getId()]) return;
        seenThreads[thread.getId()] = true;
        thread.getMessages().forEach(msg => {
          const mailId = msg.getId();
          if (known[mailId]) return;
          known[mailId] = true;
          const row = parseMail_(msg.getFrom(), msg.getSubject(), msg.getPlainBody() || '', msg.getDate());
          if (!row) { inbox.push({ mailId: mailId, skip: true }); return; }
          if (keys[row.key]) { inbox.push({ mailId: mailId, skip: true, dupOf: row.key }); return; }
          if (row.kind === 'anulacion') {
            const orig = inbox.filter(x => x.kind === 'compra' && x.card === row.card && x.amount === -row.amount && Math.abs(new Date(x.date) - new Date(row.date)) <= 3 * 864e5).pop();
            row.desc = 'Anulación · ' + (orig ? orig.desc : 'compra ' + row.card);
            if (orig) row.of = orig.mailId;
          }
          keys[row.key] = true;
          inbox.push(Object.assign({ mailId: mailId }, row));
          added++;
        });
      });
    });
    writeRaw_(INBOX_FILE, JSON.stringify(inbox.slice(-1500)));
    return added;
  });
}

function clean_(s) { return String(s || '').replace(/ /g, ' ').replace(/\t/g, ' ').replace(/\r/g, ''); }
function money_(s) { const n = parseInt(String(s).replace(/[^\d]/g, ''), 10); return isNaN(n) ? 0 : n; }
function pad2_(n) { return ('0' + n).slice(-2); }
const MONTHS_ = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12 };
function dateOf_(text, fallback) {
  let m = text.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  m = text.match(/(\d{1,2}) de (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre) de (\d{4})/i);
  if (m) return m[3] + '-' + pad2_(MONTHS_[m[2].toLowerCase()]) + '-' + pad2_(m[1]);
  return Utilities.formatDate(fallback, 'America/Santiago', 'yyyy-MM-dd');
}
function field_(text, labelRe) {
  const m = text.match(new RegExp('(?:^|\\n)[ \\t]*(?:' + labelRe + ')[ \\t]*:?[ \\t]*([^\\n]*)', 'i'));
  if (!m) return '';
  let v = m[1].trim();
  if (!v) {
    const after = text.slice(m.index + m[0].length).split('\n').map(s => s.trim()).filter(Boolean)[0] || '';
    v = /^(Datos|Nombre|Rut|Banco|Monto|Fecha|Cuenta|Tipo|Email|Mail|N[uú]mero|N°|Transacci|ID|Origen|Destino|Hora|Comercio|Si no|Te recomendamos)\b/i.test(after) ? '' : after;
  }
  return v;
}

function parseMail_(from, subject, body, date) {
  const text = clean_(body);
  if (/contacto@bci\.cl/i.test(from) && /uso de tu tarjeta/i.test(subject)) return parseCard_(text, subject, date);
  if (/transferencia|transferido|transferiste|abono/i.test(subject + ' ' + text)) return parseTransfer_(from, subject, text, date);
  return null;
}

// Compras, anulaciones y giros con tarjetas BCI.
function parseCard_(text, subject, date) {
  const amount = money_((text.match(/Monto\s*:?\s*\$\s?([\d.]+)/i) || [])[1] || '');
  if (!amount) return null;                                   // verificaciones de $0, etc.
  const actionM = text.match(/Realizaste\s+(?:un\(a\)|una|un)\s+([^\n]+?)(?:\s+con tu tarjeta|\n|$)/i);
  const action = actionM ? actionM[1].trim() : 'compra';
  const shop = field_(text, 'Comercio').replace(/[.\s]+$/, '');
  const time = field_(text, 'Hora').replace(/\s*horas?/i, '');
  const iso = dateOf_(field_(text, 'Fecha') || text, date);
  const card = /cr[eé]dito/i.test(subject + ' ' + text) ? 'crédito' : 'débito';
  const refund = /anulaci|devoluci|revers|reembolso/i.test(action);
  let desc = shop || (refund ? 'Anulación' : /giro|avance/i.test(action) ? 'Giro ' + card : 'Compra ' + card);
  if (refund) desc = 'Anulación · ' + desc.replace(/^Anulación\s*·?\s*/, '');
  return {
    date: iso, desc: desc.slice(0, 90), amount: refund ? amount : -amount,
    kind: refund ? 'anulacion' : 'compra', card: card,
    key: (refund ? 'anul|' : 'card|') + card + '|' + iso + '|' + time + '|' + amount + '|' + (shop || '').toLowerCase()
  };
}

// Transferencias de cualquier banco: enviadas, recibidas y entre tus cuentas.
function parseTransfer_(from, subject, text, date) {
  const all = subject + '\n' + text;
  const amountRaw = field_(text, 'Monto transferido|Monto abonado|Monto recibido|Monto') ||
    ((all.match(/transferencia de\s*\$\s?([\d.]+)/i) || [])[1] || '');
  const amount = money_((amountRaw.match(/\$?\s?([\d.]+)/) || [])[1] || '');
  if (!amount) return null;

  const outgoing = /Realizaste una transferencia|has realizado una transferencia|usted ha efectuado una transferencia|Enviamos tu transferencia|tu transferencia (fue enviada|se realiz)|Transferiste/i.test(all);
  const incoming = !outgoing && /a tu cuenta|te transfiri|te ha transferido|transferencia recibida|Recibiste una transferencia|ha efectuado una transferencia de fondos a/i.test(all);
  if (!outgoing && !incoming) return null;

  let who = '';
  if (outgoing) {
    who = field_(text, 'Nombre del destinatario|Nombre y apellido|Nombre y Apellido|Destinatario|Nombre');
    const m = all.match(/transferencia (?:de fondos )?a ([^,\n]+?)(?:,| el d[ií]a|\n|$)/i);
    if (!who && m) who = m[1];
  } else {
    const m = all.match(/(?:cliente|cliente\(a\))\s+([^\n]+?)\s+ha efectuado/i) || all.match(/([A-ZÁÉÍÓÚÑ][^\n]{2,60}?)\s+te (?:ha )?transferi/);
    who = m ? m[1] : field_(text, 'Nombre del originador|Nombre del remitente|Remitente|Ordenante|Origen');
  }
  who = who.replace(/\s{2,}/g, ' ').trim().slice(0, 60);
  const msg = field_(text, 'Mensaje|Asunto|Comentario').trim();
  const ref = (field_(text, 'N[uú]mero de comprobante|Transacci[oó]n|ID|N[uú]mero de operaci[oó]n') || '').replace(/\s/g, '').slice(0, 40);
  const iso = dateOf_(field_(text, 'Fecha de abono|Fecha y Hora|Fecha') + '\n' + text, date);

  const selfTransfer = OWNER_RE.test(who) && !(/para ti|a tu nombre/i.test(msg));
  const rail = /^(a|fintoc|medio de pago fintoc|khipu|flow|mercado ?pago|transbank|webpay)$/i.test(who) || /khipu|fintoc/i.test(text.slice(0, 1500)) && /^a$/i.test(who);
  const usefulMsg = msg && !/^[A-Za-z0-9]{18,}$/.test(msg) && msg.length < 80;

  let desc, kind, amountSigned;
  if (selfTransfer) {
    kind = 'traspaso';
    const bank = /bancochile/i.test(from) ? 'Banco de Chile' : /mercadopago/i.test(from) ? 'Mercado Pago' : /bci\.cl/i.test(from) ? 'BCI' : 'otra cuenta';
    desc = 'Traspaso entre mis cuentas · ' + bank;
    amountSigned = outgoing ? -amount : amount;
  } else if (outgoing) {
    kind = 'transfer_out';
    desc = rail && usefulMsg ? msg.replace(/^\d+\s+KHIPU\s+\S+\s+/i, '') : 'Transferencia a ' + (who || 'tercero');
    if (!rail && usefulMsg) desc += ' · ' + msg;
    amountSigned = -amount;
  } else {
    kind = 'transfer_in';
    desc = 'Transferencia de ' + (who || 'tercero') + (usefulMsg ? ' · ' + msg : '');
    amountSigned = amount;
  }
  return {
    date: iso, desc: desc.replace(/\s{2,}/g, ' ').slice(0, 90), amount: amountSigned, kind: kind,
    key: 'tr|' + (outgoing ? 'out' : 'in') + '|' + iso + '|' + amount + '|' + (ref || who.toLowerCase().replace(/[^a-z]/g, ''))
  };
}

// Ejecuta esta función UNA vez desde el editor para que Check revise tu correo cada 15 minutos.
function instalarRevisionCorreo() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'checkMail') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkMail').timeBased().everyMinutes(15).create();
  checkMail();
}
