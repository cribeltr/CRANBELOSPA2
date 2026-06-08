/****************************************************************************
 *  Puente Google Sheets — Programación MP 2026
 *  --------------------------------------------------------------------------
 *  Crea/actualiza en tu planilla las MISMAS hojas que el Excel:
 *  "Eventos", "Catalogos" y "Resumen", con desplegables (Resultado, Estado
 *  Final, Ejecutor) y coloreado por Estado en la hoja Eventos.
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
 *  Cada hoja se REEMPLAZA con el contenido enviado (las demás hojas no se tocan).
 ****************************************************************************/

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = [];

    (body.sheets || []).forEach(function (spec) {
      var name = spec.name || 'Hoja';
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);

      // Limpiar contenido, validaciones y formato condicional previos
      sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
      sh.setConditionalFormatRules([]);
      sh.clearContents();

      var rows = spec.rows || [];
      if (rows.length > 0) {
        var maxc = 1;
        rows.forEach(function (r) { if (r.length > maxc) maxc = r.length; });
        rows.forEach(function (r) { while (r.length < maxc) r.push(''); });
        sh.getRange(1, 1, rows.length, maxc).setValues(rows);
        sh.setFrozenRows(1);
        applyValidations(sh, spec.validations);
        applyColors(sh, spec.colors);
      }
      written.push(name);
    });

    // Snapshot de datos de la app (para poder LEERLOS al reabrir el programa)
    if (body.data !== undefined && body.data !== null) {
      var ds = ss.getSheetByName('_datos') || ss.insertSheet('_datos');
      ds.getRange(1, 1).setValue(typeof body.data === 'string' ? body.data : JSON.stringify(body.data));
      try { ds.hideSheet(); } catch (e2) {}
    }

    return json({ ok: true, sheets: written });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Columna (1-based) cuyo encabezado (fila 1) coincide con 'header'
function colByHeader(sh, header) {
  var lastC = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastC).getValues()[0];
  for (var i = 0; i < hdr.length; i++) { if (String(hdr[i]) === header) return i + 1; }
  return 0;
}

// Desplegables (validación de datos por lista) en las columnas indicadas
function applyValidations(sh, validations) {
  if (!validations) return;
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  validations.forEach(function (v) {
    var c = colByHeader(sh, v.header);
    if (!c) return;
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(v.values, true)
      .setAllowInvalid(true)
      .build();
    sh.getRange(2, c, n, 1).setDataValidation(rule);
  });
}

// Número de columna -> letra (1 -> A, 27 -> AA)
function columnToLetter(col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - 1) / 26); }
  return s;
}

// Coloreado por valor (formato condicional). Si la regla es wholeRow, colorea
// TODA la fila usando una fórmula que mira la columna del encabezado indicado.
function applyColors(sh, colors) {
  if (!colors) return;
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var lastCol = sh.getLastColumn();
  var rules = [];
  colors.forEach(function (cf) {
    var c = colByHeader(sh, cf.header);
    if (!c) return;
    if (cf.wholeRow) {
      var rowRange = sh.getRange(2, 1, n, lastCol);    // todas las columnas, filas de datos
      var colL = columnToLetter(c);
      (cf.rules || []).forEach(function (r) {
        var formula = '=$' + colL + '2="' + r.value + '"';
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula).setBackground(r.color).setRanges([rowRange]).build());
      });
    } else {
      var rng = sh.getRange(2, c, n, 1);
      (cf.rules || []).forEach(function (r) {
        var b = SpreadsheetApp.newConditionalFormatRule().setBackground(r.color).setRanges([rng]);
        b = (r.mode === 'startsWith') ? b.whenTextStartsWith(r.value) : b.whenTextEqualTo(r.value);
        rules.push(b.build());
      });
    }
  });
  if (rules.length) sh.setConditionalFormatRules(rules);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Eventos');
  var n = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  var ds = ss.getSheetByName('_datos');
  var data = ds ? ds.getRange(1, 1).getValue() : '';
  return json({ ok: true, count: n, data: data });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
