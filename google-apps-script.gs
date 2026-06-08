/****************************************************************************
 *  Puente Google Sheets — Registro de Mantenciones MP 2026
 *  --------------------------------------------------------------------------
 *  Cómo usarlo (configuración única):
 *   1. Crea (o abre) una Google Sheet.
 *   2. Menú: Extensiones -> Apps Script.
 *   3. Borra el contenido y pega ESTE archivo completo. Guarda (icono 💾).
 *   4. Implementar -> Nueva implementación -> tipo "Aplicación web".
 *        - Ejecutar como: Yo
 *        - Quién tiene acceso: Cualquier persona
 *      Implementar y autorizar los permisos.
 *   5. Copia la "URL de la aplicación web" (termina en /exec) y pégala en la
 *      app (sección "Guardar en Google Sheets").
 *
 *  La app envía las mantenciones registradas; cada una se identifica por un
 *  UID (columna A) y se ACTUALIZA si ya existe (no se duplica).
 ****************************************************************************/

var SHEET_NAME = 'Mantenciones';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    var headers = body.headers || [];
    if (sh.getLastRow() === 0 && headers.length) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    // Mapa UID -> fila existente (columna A) para actualizar en vez de duplicar
    var existing = {};
    if (sh.getLastRow() >= 2) {
      var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      ids.forEach(function (r, i) { if (r[0] !== '') existing[r[0]] = i + 2; });
    }

    var added = 0, updated = 0;
    (body.rows || []).forEach(function (row) {
      var uid = row[0];
      if (uid !== '' && existing[uid]) {
        sh.getRange(existing[uid], 1, 1, row.length).setValues([row]);
        updated++;
      } else {
        sh.appendRow(row);
        added++;
      }
    });

    return json({ ok: true, added: added, updated: updated });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  var n = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  return json({ ok: true, count: n });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
