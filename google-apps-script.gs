/****************************************************************************
 *  Puente Google Sheets — Programación MP 2026
 *  --------------------------------------------------------------------------
 *  Crea/actualiza en tu planilla las MISMAS hojas que el Excel:
 *  "Eventos", "Catalogos" y "Resumen".
 *
 *  Cómo usarlo (configuración única):
 *   1. Crea (o abre) una Google Sheet.
 *   2. Menú: Extensiones -> Apps Script.
 *   3. Borra el contenido y pega ESTE archivo completo. Guarda (icono de disco).
 *   4. Implementar -> Nueva implementación -> tipo "Aplicación web".
 *        - Ejecutar como: Yo
 *        - Quién tiene acceso: Cualquier persona
 *      Implementar y autorizar los permisos.
 *   5. Copia la "URL de la aplicación web" (termina en /exec) y pégala en la
 *      app (sección "Guardar en Google Sheets").
 *
 *  La app envía un libro completo; cada hoja se REEMPLAZA con el contenido
 *  enviado (las demás hojas de la planilla no se tocan).
 ****************************************************************************/

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = body.sheets || [];
    var written = [];

    sheets.forEach(function (spec) {
      var name = spec.name || 'Hoja';
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clearContents();
      var rows = spec.rows || [];
      if (rows.length > 0) {
        // Todas las filas deben tener el mismo número de columnas para setValues
        var maxc = 1;
        rows.forEach(function (r) { if (r.length > maxc) maxc = r.length; });
        rows.forEach(function (r) { while (r.length < maxc) r.push(''); });
        sh.getRange(1, 1, rows.length, maxc).setValues(rows);
        sh.setFrozenRows(1);
      }
      written.push(name);
    });

    return json({ ok: true, sheets: written });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Eventos');
  var n = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  return json({ ok: true, count: n });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
