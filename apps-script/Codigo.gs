/**
 * ZonaMarcas · UbikMe GPS — base de datos en Google Sheets
 * Versión con protección contra pérdida de datos.
 *
 * Protecciones:
 *  1. Lectura y escritura con candado: nadie lee mientras otro está guardando (evita lecturas vacías).
 *  2. Solo acepta guardados de la app versión 3 o superior (bloquea equipos con la versión vieja abierta).
 *  3. Rechaza guardados que borren muchos registros de una vez.
 *  4. Antes de cada guardado copia lo que había en la pestaña "respaldos" (últimas 300 copias).
 */
const SHEET_NAME = 'datos';
const BACKUP_SHEET = 'respaldos';
const CHUNK = 45000;
const VERSION_MINIMA = 3;
const MAX_RESPALDOS = 300;
const LISTAS = ['marcacion_registros', 'gps_registros'];
const MAX_BORRADOS_POR_GUARDADO = 3;

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

function contar_(texto) {
  try {
    const a = JSON.parse(texto);
    return Array.isArray(a) ? a.length : null;
  } catch (err) {
    return null;
  }
}

function respaldar_(key, valorAnterior) {
  if (!valorAnterior) return;
  const sh = getBackupSheet_();
  const chunks = trocear_(valorAnterior);
  const row = sh.getLastRow() + 1;
  const n = contar_(valorAnterior);
  sh.getRange(row, 1, 1, 3).setValues([[new Date(), key, n === null ? '' : n]]);
  sh.getRange(row, 4, 1, chunks.length).setNumberFormat('@').setValues([chunks]);
  const sobran = sh.getLastRow() - 1 - MAX_RESPALDOS;
  if (sobran > 0) sh.deleteRows(2, sobran);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);   // espera si alguien está guardando: nunca devuelve un valor a medio escribir
    const key = e.parameter.key;
    if (!key) return json_({ ok: false, error: 'Falta la clave' });
    const sh = getSheet_();
    const row = findRow_(sh, key);
    return json_({ ok: true, value: row === -1 ? null : readValue_(sh, row) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    const key = String(body.key || '');
    const value = String(body.value == null ? '' : body.value);
    if (!key) return json_({ ok: false, error: 'Falta la clave' });

    // 2. Bloquear versiones viejas de la app
    if (Number(body.version || 0) < VERSION_MINIMA) {
      return json_({ ok: false, error: 'Esta página está desactualizada y no puede guardar. Ciérrala y vuelve a abrirla.' });
    }

    const sh = getSheet_();
    let row = findRow_(sh, key);
    const anterior = row === -1 ? '' : readValue_(sh, row);

    // 3. Bloquear borrados masivos en las listas
    if (LISTAS.indexOf(key) >= 0) {
      const nNuevo = contar_(value);
      if (nNuevo === null) return json_({ ok: false, error: 'Datos inválidos: no se guardó.' });
      const nAnterior = anterior ? contar_(anterior) : null;
      if (nAnterior !== null && nNuevo < nAnterior - MAX_BORRADOS_POR_GUARDADO) {
        return json_({ ok: false, error: 'Se bloqueó un guardado que iba a borrar ' + (nAnterior - nNuevo) + ' registros. Recarga la página.' });
      }
    }

    // 4. Copia de respaldo de lo que había
    respaldar_(key, anterior);

    if (row === -1) {
      row = sh.getLastRow() + 1;
      sh.getRange(row, 1).setValue(key);
    }
    const chunks = trocear_(value);
    const lastCol = sh.getLastColumn();
    if (lastCol > 1) sh.getRange(row, 2, 1, lastCol - 1).clearContent();
    sh.getRange(row, 2, 1, chunks.length).setNumberFormat('@').setValues([chunks]);
    SpreadsheetApp.flush();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
