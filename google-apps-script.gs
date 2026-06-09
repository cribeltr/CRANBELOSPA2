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
 *
 *  -- OPCIONAL: servir la app DESDE aquí (el link /exec abre el programa) --
 *   a. En el editor de Apps Script: "+" -> HTML -> nómbralo EXACTAMENTE "index".
 *   b. Pega dentro TODO el contenido de "index-appsscript.html" (NO el index.html
 *      grande: dentro de Apps Script el ExcelJS incrustado no se ejecuta bien;
 *      la versión -appsscript lo carga por CDN y sí funciona).
 *   c. Implementar -> Nueva versión. La URL /exec abre la app ya conectada a
 *      esta planilla (escribe y lee con google.script.run, sin pegar la URL).
 *   Nota: dentro de Apps Script el botón "Descargar Excel" puede quedar
 *   bloqueado por el iframe protegido; para descargar usa la app alojada aparte.
 *
 *  -- Subir archivos a Drive (opcional) --
 *   La app puede subir archivos del equipo a tu Drive: crea la carpeta
 *   "MP 2026 - Archivos" con una subcarpeta por equipo y agrega el ENLACE a la
 *   hoja "Archivos". La PRIMERA vez pedirá un permiso adicional de Drive: vuelve
 *   a "Implementar -> Nueva versión" y autoriza cuando lo solicite.
 ****************************************************************************/

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body && body.action === 'upload') return json(uploadArchivo(body));   // subir archivo a Drive
    return json(writeAll(body));
  }
  catch (err) { return json({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

// Llamado desde la app cuando se sirve DESDE Apps Script (google.script.run)
function appPush(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try { return writeAll(typeof body === 'string' ? JSON.parse(body) : body); }
  finally { lock.releaseLock(); }
}
function appPull() { return readAll(); }
// Subir archivo a Drive desde la app servida en Apps Script (google.script.run)
function appUpload(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try { return uploadArchivo(typeof body === 'string' ? JSON.parse(body) : body); }
  finally { lock.releaseLock(); }
}

/* ------------------------- Subida de archivos a Drive -------------------------
 *  Crea (si no existe) la carpeta raíz "MP 2026 - Archivos" y, dentro, una
 *  subcarpeta por equipo. Guarda el archivo, lo comparte como "cualquiera con el
 *  enlace puede ver" y agrega una fila con el ENLACE a la hoja "Archivos".
 *  payload: { equipoId, equipo, inv, serie, servicio, categoria, descripcion, nombre, mime, dataBase64 }
 *  Requiere autorización de Drive (al implementar pedirá el permiso una vez).
 * ----------------------------------------------------------------------------- */
var ARCHIVOS_ROOT = 'MP 2026 - Archivos';

function getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function sanitizeName(s) { return String(s == null ? '' : s).replace(/[\\/:*?"<>|]/g, '-').trim() || 'Equipo'; }

function uploadArchivo(p) {
  if (!p || !p.dataBase64) return { ok: false, error: 'Sin datos de archivo.' };
  var root = getOrCreateFolder(DriveApp.getRootFolder(), ARCHIVOS_ROOT);
  var carpetaEquipo = sanitizeName(
    (p.equipoId ? p.equipoId + ' - ' : '') + (p.equipo || 'Equipo') +
    (p.inv ? ' (Inv ' + p.inv + ')' : (p.serie ? ' (Serie ' + p.serie + ')' : ''))
  );
  var folder = getOrCreateFolder(root, carpetaEquipo);

  var nombre = sanitizeName(p.nombre || 'archivo');
  var bytes = Utilities.base64Decode(p.dataBase64);
  var blob = Utilities.newBlob(bytes, p.mime || 'application/octet-stream', nombre);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
  var url = file.getUrl();

  // Registrar el enlace en la hoja "Archivos"
  var HEADERS = ['ID', 'N° Inventario', 'N° Serie', 'Equipo', 'Servicio', 'Categoría', 'Descripción', 'Nombre del archivo', 'Enlace', 'Fecha de carga'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Archivos');
  if (!sh) {
    sh = ss.insertSheet('Archivos');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#15616d').setFontColor('#ffffff');
  }
  // Si la hoja ya existía sin la columna "Descripción", insertarla tras "Categoría".
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  if (hdr.indexOf('Descripción') === -1) {
    var catIdx = hdr.indexOf('Categoría');
    var at = (catIdx >= 0 ? catIdx + 2 : sh.getLastColumn() + 1);
    sh.insertColumnBefore(at);
    sh.getRange(1, at).setValue('Descripción').setFontWeight('bold').setBackground('#15616d').setFontColor('#ffffff');
    hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  }
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Santiago', 'yyyy-MM-dd HH:mm');
  // Escribir en el ORDEN real del encabezado (el enlace como texto: Google Sheets
  // lo hace clicable y queda legible al volver a leer la hoja, round-trip).
  var valByHeader = {
    'ID': p.equipoId || '', 'N° Inventario': p.inv || '', 'N° Serie': p.serie || '', 'Equipo': p.equipo || '',
    'Servicio': p.servicio || '', 'Categoría': p.categoria || '', 'Descripción': p.descripcion || '',
    'Nombre del archivo': nombre, 'Enlace': url, 'Fecha de carga': fecha
  };
  sh.appendRow(hdr.map(function (h) { return (h in valByHeader) ? valByHeader[h] : ''; }));

  return { ok: true, url: url, name: nombre, folder: folder.getUrl(), categoria: p.categoria || '', descripcion: p.descripcion || '', fecha: fecha };
}

function writeAll(body) {
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

    return { ok: true, sheets: written };
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

function sheetVals(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  return (lr >= 1 && lc >= 1) ? sh.getRange(1, 1, lr, lc).getValues() : [];
}
function readAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Eventos');
  var n = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  var ds = ss.getSheetByName('_datos');
  var data = ds ? ds.getRange(1, 1).getValue() : '';
  return {
    ok: true, count: n, data: data,
    tablas: {
      Pendientes: sheetVals(ss, 'Pendientes'),
      Correctivos: sheetVals(ss, 'Correctivos'),
      Tareas: sheetVals(ss, 'Tareas'),
      Bitacora: sheetVals(ss, 'Bitacora'),
      Archivos: sheetVals(ss, 'Archivos')
    }
  };
}
// doGet: con ?api=1 devuelve los datos (JSON); sin parámetros sirve la app.
function doGet(e) {
  if (e && e.parameter && e.parameter.api) return json(readAll());
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gestión MP 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
