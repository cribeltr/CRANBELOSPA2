/* ============================================================================
 *  build.js — Ensambla un index.html autocontenido (sin dependencias externas)
 *  Inserta ExcelJS (vendor) + core.js + output.js + app.js dentro del template.
 *  Uso:  node build.js
 * ==========================================================================*/
const fs = require('fs');
const path = require('path');

const root = __dirname;
const read = p => fs.readFileSync(path.join(root, p), 'utf8');

// Evita que un "</script>" dentro del JS cierre el bloque antes de tiempo,
// y elimina los comentarios sourceMappingURL (los .map no se distribuyen).
function inlineScript(jsPath) {
  let js = read(jsPath);
  js = js.replace(/\/\/[#@]\s*sourceMappingURL=.*$/gm, '');
  js = js.replace(/<\/script>/gi, '<\\/script>');
  return '<script>\n' + js + '\n</script>';
}

// Texto a insertar dentro de un <pre> (escapa &, < y >).
function inlineText(p) {
  return read(p).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

let html = read('src/template.html');
const replacements = {
  '<!--EXCELJS-->': inlineScript('vendor/exceljs.min.js'),
  '<!--CORE-->':    inlineScript('src/core.js'),
  '<!--OUTPUT-->':  inlineScript('src/output.js'),
  '<!--APP-->':     inlineScript('src/app.js'),
  '<!--GASCRIPT-->': inlineText('google-apps-script.gs')
};
for (const [marker, content] of Object.entries(replacements)) {
  if (!html.includes(marker)) throw new Error('Marcador no encontrado en template: ' + marker);
  // Reemplazo con función: evita la interpretación de patrones "$" del JS minificado.
  html = html.replace(marker, () => content);
}

const out = path.join(root, 'index.html');
fs.writeFileSync(out, html);
const kb = (fs.statSync(out).size / 1024).toFixed(0);
console.log('index.html generado (' + kb + ' KB).');
