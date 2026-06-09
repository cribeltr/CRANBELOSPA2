/* ============================================================================
 *  build.js — Ensambla index.html autocontenido (sin dependencias externas).
 *  Inserta ExcelJS (vendor) + core.js + output.js + app.js dentro del template.
 *
 *  Genera DOS archivos:
 *   - index.html               : ExcelJS INCRUSTADA (funciona 100% sin internet;
 *                                ideal para GitHub Pages / Netlify / abrir local).
 *   - index-appsscript.html    : ExcelJS desde CDN (archivo liviano y fiable para
 *                                PEGAR dentro de Apps Script; el iframe protegido
 *                                de Apps Script no ejecuta bien el script gigante
 *                                incrustado, así que ahí cargamos ExcelJS por CDN).
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

// ExcelJS desde CDN, con un segundo CDN de respaldo si el primero falla.
const EXCELJS_CDN =
  '<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>\n' +
  '<script>if(typeof ExcelJS==="undefined"){document.write(\'<scr\'+\'ipt src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"><\\/scr\'+\'ipt>\')}</script>';

const template = read('src/template.html');
const shared = {
  '<!--CORE-->':     inlineScript('src/core.js'),
  '<!--OUTPUT-->':   inlineScript('src/output.js'),
  '<!--APP-->':      inlineScript('src/app.js'),
  '<!--GASCRIPT-->': inlineText('google-apps-script.gs')
};

function assemble(exceljsBlock) {
  let html = template;
  const reps = Object.assign({ '<!--EXCELJS-->': exceljsBlock }, shared);
  for (const [marker, content] of Object.entries(reps)) {
    if (!html.includes(marker)) throw new Error('Marcador no encontrado en template: ' + marker);
    html = html.replace(marker, () => content);   // función: no interpreta "$" del JS minificado
  }
  return html;
}

function write(name, html) {
  const out = path.join(root, name);
  fs.writeFileSync(out, html);
  const kb = (fs.statSync(out).size / 1024).toFixed(0);
  console.log(name + ' generado (' + kb + ' KB).');
}

write('index.html', assemble(inlineScript('vendor/exceljs.min.js')));
write('index-appsscript.html', assemble(EXCELJS_CDN));
