/**
 * ZonaMarcas · UbikMe GPS — base de datos en Google Sheets
 * Versión 2 del servidor: guardado por acciones (sin pérdida de datos entre varios usuarios).
 *
 * Cómo funciona:
 *  - La app ya no envía la lista completa: envía solo la acción (registrar, marcar, completar etapa,
 *    deshacer, eliminar). Este programa la aplica sobre lo último que hay en la hoja, todo dentro
 *    de un candado, así dos personas que guardan al mismo tiempo no se pisan.
 *  - Lectura con candado: nunca se entrega un valor a medio escribir.
 *  - Una sola consulta trae todos los datos (más rápido para todos).
 *  - Numeración de remisiones sin repetir.
 *  - Rechaza guardados de páginas desactualizadas y borrados masivos.
 *  - Copia de respaldo en la pestaña "respaldos" (como máximo una cada 3 minutos por lista; últimas 300).
 */
const SHEET_NAME = 'datos';
const BACKUP_SHEET = 'respaldos';
const CHUNK = 45000;
const SERVIDOR_VERSION = 2;
const VERSION_MINIMA = 4;                 // la app debe ser versión 4 o superior para guardar
const LISTAS = ['marcacion_registros', 'gps_registros'];
const MAX_BORRADOS_POR_GUARDADO = 3;
const MAX_RESPALDOS = 300;
const MINUTOS_ENTRE_RESPALDOS = 3;

/* ---------------- Hoja ---------------- */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['clave', 'valor']);
  }
  return sh;
}

function getBackupSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(BACKUP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BACKUP_SHEET);
    sh.appendRow(['fecha', 'clave', 'registros', 'valor']);
  }
  return sh;
}

function findRow_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === key) return i + 2;
  }
  return -1;
}

function readValue_(sh, row) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return '';
  return sh.getRange(row, 2, 1, lastCol - 1).getValues()[0].join('');
}

function trocear_(value) {
  const chunks = [];
  for (let i = 0; i < value.length; i += CHUNK) chunks.push(value.substring(i, i + CHUNK));
  if (chunks.length === 0) chunks.push('');
  return chunks;
}

function leer_(sh, key) {
  const row = findRow_(sh, key);
  return row === -1 ? null : readValue_(sh, row);
}

function escribir_(sh, key, value) {
  let row = findRow_(sh, key);
  if (row === -1) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue(key);
  }
  const chunks = trocear_(value);
  const lastCol = sh.getLastColumn();
  if (lastCol > 1) sh.getRange(row, 2, 1, lastCol - 1).clearContent();
  sh.getRange(row, 2, 1, chunks.length).setNumberFormat('@').setValues([chunks]);
  SpreadsheetApp.flush();
}

function contar_(texto) {
  try {
    const a = JSON.parse(texto);
    return Array.isArray(a) ? a.length : null;
  } catch (err) {
    return null;
  }
}

/* ---------------- Respaldos ---------------- */
function respaldar_(key, valorAnterior) {
  if (!valorAnterior) return;
  const props = PropertiesService.getScriptProperties();
  const prop = 'ultimo_respaldo_' + key;
  const ultimo = Number(props.getProperty(prop) || 0);
  if (Date.now() - ultimo < MINUTOS_ENTRE_RESPALDOS * 60000) return;
  const sh = getBackupSheet_();
  const chunks = trocear_(valorAnterior);
  const row = sh.getLastRow() + 1;
  const n = contar_(valorAnterior);
  sh.getRange(row, 1, 1, 3).setValues([[new Date(), key, n === null ? '' : n]]);
  sh.getRange(row, 4, 1, chunks.length).setNumberFormat('@').setValues([chunks]);
  const sobran = sh.getLastRow() - 1 - MAX_RESPALDOS;
  if (sobran > 0) sh.deleteRows(2, sobran);
  props.setProperty(prop, String(Date.now()));
}

/* ---------------- Acciones sobre un registro ---------------- */
function getPath_(obj, path) {
  const ks = path.split('.');
  let o = obj;
  for (let i = 0; i < ks.length; i++) {
    if (o === null || typeof o !== 'object') return undefined;
    o = o[ks[i]];
  }
  return o;
}
function setPath_(obj, path, val) {
  const ks = path.split('.');
  let o = obj;
  for (let i = 0; i < ks.length - 1; i++) {
    if (o[ks[i]] === null || typeof o[ks[i]] !== 'object') o[ks[i]] = {};
    o = o[ks[i]];
  }
  o[ks[ks.length - 1]] = val;
}
function unsetPath_(obj, path) {
  const ks = path.split('.');
  let o = obj;
  for (let i = 0; i < ks.length - 1; i++) {
    if (o === null || typeof o !== 'object') return;
    o = o[ks[i]];
  }
  if (o && typeof o === 'object') delete o[ks[ks.length - 1]];
}
/* cambio = { id, setSiFalta:{ruta:valor}, set:{ruta:valor}, unset:[ruta] } */
function aplicarCambio_(it, c) {
  if (c.setSiFalta) Object.keys(c.setSiFalta).forEach(function (p) {
    const v = getPath_(it, p);
    if (v === undefined || v === null || v === '') setPath_(it, p, c.setSiFalta[p]);
  });
  if (c.set) Object.keys(c.set).forEach(function (p) { setPath_(it, p, c.set[p]); });
  if (c.unset) c.unset.forEach(function (p) { unsetPath_(it, p); });
}

/* ---------------- Respuestas ---------------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function rechazo_(msg) {
  return json_({ ok: false, rechazo: true, error: msg });
}

/* ---------------- Lectura ---------------- */
function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = e.parameter || {};
    if (p.action === 'info') return json_({ ok: true, servidor: SERVIDOR_VERSION });
    lock.waitLock(30000);
    const sh = getSheet_();
    if (p.action === 'getall') {
      const keys = String(p.keys || '').split(',').filter(String);
      const values = {};
      keys.forEach(function (k) { values[k] = leer_(sh, k); });
      return json_({ ok: true, values: values });
    }
    const key = p.key;
    if (!key) return json_({ ok: false, error: 'Falta la clave' });
    return json_({ ok: true, value: leer_(sh, key) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------------- Escritura ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || 'set');
    const key = String(body.key || '');
    if (!key) return json_({ ok: false, error: 'Falta la clave' });
    if (Number(body.version || 0) < VERSION_MINIMA) {
      return rechazo_('Esta página está desactualizada y no puede guardar. Ciérrala y vuelve a abrirla.');
    }
    lock.waitLock(30000);
    const sh = getSheet_();

    // --- Acciones sobre una lista (registrar, marcar, completar, deshacer, eliminar) ---
    if (action === 'ops') {
      if (LISTAS.indexOf(key) < 0) return rechazo_('Lista desconocida.');
      const borrar = body.borrar || [];
      if (borrar.length > MAX_BORRADOS_POR_GUARDADO) {
        return rechazo_('Se bloqueó un guardado que iba a borrar ' + borrar.length + ' registros.');
      }
      const anterior = leer_(sh, key) || '';
      let lista = [];
      if (anterior) {
        try { lista = JSON.parse(anterior); } catch (x) { return json_({ ok: false, error: 'La hoja tiene datos dañados; no se guardó.' }); }
        if (!Array.isArray(lista)) return json_({ ok: false, error: 'La hoja tiene datos dañados; no se guardó.' });
      }
      const borrarSet = {};
      borrar.forEach(function (id) { borrarSet[id] = true; });
      lista = lista.filter(function (x) { return !borrarSet[x.id]; });
      const porId = {};
      lista.forEach(function (x) { porId[x.id] = x; });
      (body.cambios || []).forEach(function (c) { if (porId[c.id]) aplicarCambio_(porId[c.id], c); });
      (body.nuevos || []).forEach(function (it) {
        if (it && it.id && !porId[it.id] && !borrarSet[it.id]) { lista.push(it); porId[it.id] = it; }
      });
      const nuevo = JSON.stringify(lista);
      respaldar_(key, anterior);
      escribir_(sh, key, nuevo);
      return json_({ ok: true, value: nuevo });
    }

    // --- Consecutivo (número de remisión) sin repetir ---
    if (action === 'incr') {
      const actual = parseInt(leer_(sh, key) || '', 10);
      const siguiente = (isNaN(actual) ? Number(body.inicio || 0) : actual) + 1;
      escribir_(sh, key, String(siguiente));
      return json_({ ok: true, value: siguiente });
    }

    // --- Valor suelto (lista de vendedores, etc.) ---
    if (LISTAS.indexOf(key) >= 0) {
      return rechazo_('Esta página está desactualizada y no puede guardar. Ciérrala y vuelve a abrirla.');
    }
    const value = String(body.value == null ? '' : body.value);
    respaldar_(key, leer_(sh, key) || '');
    escribir_(sh, key, value);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
