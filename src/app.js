/* ============================================================================
 *  Interfaz de usuario — Generador de Eventos MP 2026
 *  Usa los globales: ExcelJS, MP (core), MPOUT (output)
 * ==========================================================================*/
(function () {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const drop = $('#drop');
  const fileInput = $('#file');
  const pickBtn = $('#pick');
  const statusBox = $('#status');
  const resultBox = $('#result');
  const dlBtn = $('#download');
  const fileNameEl = $('#fileName');

  let lastParse = null;     // { events, equipos, warnings, sourceName }

  function setStatus(msg, kind) {
    statusBox.className = 'status ' + (kind || '');
    statusBox.innerHTML = msg;
    statusBox.style.display = msg ? 'block' : 'none';
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
  function pad2(n) { return ('0' + n).slice(-2); }
  function stamp() {
    const d = new Date();
    return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate());
  }

  /* ---------- Lectura del archivo subido ---------- */
  function readFile(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result);
      fr.onerror = () => reject(new Error('No se pudo leer el archivo.'));
      fr.readAsArrayBuffer(file);
    });
  }

  async function handleFile(file) {
    if (!file) return;
    const name = file.name || 'archivo';
    if (!/\.(xlsm|xlsx)$/i.test(name)) {
      setStatus('⚠️ El archivo debe ser <b>.xlsm</b> o <b>.xlsx</b> (Programación MP 2026).', 'warn');
      return;
    }
    if (typeof ExcelJS === 'undefined') {
      setStatus('❌ No se cargó la librería ExcelJS. Verifique su conexión o el archivo vendor/exceljs.min.js.', 'error');
      return;
    }
    resultBox.style.display = 'none';
    dlBtn.disabled = true;
    lastParse = null;
    setStatus('⏳ Procesando <b>' + esc(name) + '</b>… (puede tardar unos segundos)', 'info');

    // Dar un respiro al navegador para pintar el estado antes de trabajar.
    await new Promise(r => setTimeout(r, 50));

    try {
      const buf = await readFile(file);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);

      if (!wb.getWorksheet('PMP_2026')) {
        throw new Error('La hoja "PMP_2026" no existe en el archivo. ¿Es el archivo de Programación MP 2026 correcto?');
      }

      const parsed = MP.parseWorkbook(wb);
      parsed.sourceName = name;
      lastParse = parsed;

      renderResult(parsed);
      dlBtn.disabled = false;
      let msg = '✅ Procesado: <b>' + parsed.equipos.toLocaleString('es-CL') +
        '</b> equipos · <b>' + parsed.events.length.toLocaleString('es-CL') + '</b> eventos.';
      if (parsed.warnings.length) msg += '<br>⚠️ ' + parsed.warnings.map(esc).join('<br>⚠️ ');
      setStatus(msg, parsed.warnings.length ? 'warn' : 'ok');
    } catch (err) {
      console.error(err);
      setStatus('❌ Error al procesar: ' + esc(err.message || err), 'error');
    }
  }

  /* ---------- Render de resumen + vista previa ---------- */
  function renderResult(parsed) {
    const ev = parsed.events;
    const stats = MP.buildStats(ev);
    const estados = [...stats.byEstado.entries()].sort((a, b) => b[1] - a[1]);

    let chips = estados.map(([k, v]) =>
      '<span class="chip">' + esc(k) + ' <b>' + v.toLocaleString('es-CL') + '</b></span>').join('');

    const cols = ['ID', 'Equipo', 'Servicio', 'Mes', 'P', 'Tipo de Programación', 'R', 'Estado', 'Causal'];
    let head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let rows = ev.slice(0, 20).map(e =>
      '<tr>' + [e.id, e.equipo, e.servicio, e.mes, e.programa, e.tipoPrograma,
                e.resultado, e.estado, e.causal]
        .map(x => '<td>' + esc(x) + '</td>').join('') + '</tr>').join('');

    resultBox.innerHTML =
      '<div class="chips">' + chips + '</div>' +
      '<div class="muted" style="margin:10px 0 6px">Vista previa (primeros 20 de ' +
        ev.length.toLocaleString('es-CL') + ' eventos):</div>' +
      '<div class="tablewrap"><table class="prev">' + head + rows + '</table></div>';
    resultBox.style.display = 'block';
  }

  /* ---------- Descarga del Excel ---------- */
  async function download() {
    if (!lastParse) return;
    dlBtn.disabled = true;
    const original = dlBtn.textContent;
    dlBtn.textContent = 'Generando Excel…';
    try {
      const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, lastParse.events, { equipos: lastParse.equipos });
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Eventos_MP_2026_' + stamp() + '.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
    } catch (err) {
      console.error(err);
      setStatus('❌ Error al generar el Excel: ' + esc(err.message || err), 'error');
    } finally {
      dlBtn.textContent = original;
      dlBtn.disabled = false;
    }
  }

  /* ---------- Eventos de UI ---------- */
  pickBtn.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
    if (f) { fileNameEl.textContent = f.name; handleFile(f); }
  });
  dlBtn.addEventListener('click', download);

  ['dragenter', 'dragover'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) { fileNameEl.textContent = f.name; handleFile(f); }
  });
})();
