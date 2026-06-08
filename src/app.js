/* ============================================================================
 *  Interfaz de usuario — Generador de Eventos MP 2026
 *  Usa los globales: ExcelJS, MP (core), MPOUT (output)
 *  Dos funciones:
 *    1) Cargar archivo y generar el Excel masivo (una fila por evento).
 *    2) Registrar mantenciones (buscar equipo, validar mes, completar y exportar).
 * ==========================================================================*/
(function () {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const state = { equipos: [], events: [], registros: [], correctivos: [], searchResults: [], selEq: null, selMonth: null, selDetalleEq: null, sgActive: -1, cEq: null };

  // ---- Utilidades ---------------------------------------------------------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
  const pad2 = n => ('0' + n).slice(-2);
  function stamp() { const d = new Date(); return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()); }
  function fmtDate(d) { return d instanceof Date ? pad2(d.getDate()) + '-' + pad2(d.getMonth() + 1) + '-' + d.getFullYear() : ''; }
  function showInline(el, kind, msg) {
    el.className = 'inline-msg ' + (kind || '') + (msg ? ' show' : '');
    el.innerHTML = msg || '';
  }
  function setStatus(msg, kind) {
    const s = $('#status');
    s.className = 'status ' + (kind || '');
    s.innerHTML = msg;
    s.style.display = msg ? 'block' : 'none';
  }

  // ---- Persistencia de registros + clave estable -------------------------
  const LS_KEY = 'mp_registros_2026';
  function isoDate(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  // Identidad estable del evento: Inventario > Serie > ID, más el mes.
  function keyOf(e) {
    const inv = (e.inv != null ? String(e.inv) : '').trim();
    const ser = (e.serie != null ? String(e.serie) : '').trim();
    const base = inv ? 'I:' + inv.toLowerCase() : ser ? 'S:' + ser.toLowerCase() : 'ID:' + e.id;
    return base + '|' + e.nMes;
  }
  function saveRegistros() {
    try {
      const data = state.registros.map(r => {
        const o = Object.assign({}, r);
        o.fechaEjecucion = (r.fechaEjecucion instanceof Date) ? isoDate(r.fechaEjecucion) : '';
        return o;
      });
      localStorage.setItem(LS_KEY, JSON.stringify(data));
    } catch (e) { /* localStorage no disponible (p. ej. algunos navegadores en file://) */ }
  }
  function loadRegistros() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return [];
      return JSON.parse(raw).map(o => {
        const r = Object.assign({}, o);
        if (typeof r.fechaEjecucion === 'string' && r.fechaEjecucion) {
          const p = r.fechaEjecucion.split('-'); r.fechaEjecucion = new Date(+p[0], +p[1] - 1, +p[2]);
        } else r.fechaEjecucion = '';
        return r;
      });
    } catch (e) { return []; }
  }

  // ---- Persistencia de eventos correctivos -------------------------------
  const LS_CORR = 'mp_correctivos_2026';
  function saveCorrectivos() {
    try {
      const data = state.correctivos.map(c => {
        const o = Object.assign({}, c);
        o.fecha = (c.fecha instanceof Date) ? isoDate(c.fecha) : (c.fecha || '');
        return o;
      });
      localStorage.setItem(LS_CORR, JSON.stringify(data));
    } catch (e) { /* localStorage no disponible */ }
  }
  function loadCorrectivos() {
    try {
      const raw = localStorage.getItem(LS_CORR);
      if (!raw) return [];
      return JSON.parse(raw).map(o => {
        const c = Object.assign({}, o);
        if (typeof c.fecha === 'string' && c.fecha) { const p = c.fecha.split('-'); c.fecha = new Date(+p[0], +p[1] - 1, +p[2]); }
        else c.fecha = '';
        return c;
      });
    } catch (e) { return []; }
  }

  // ---- Integración con Google Sheets (Apps Script) -----------------------
  // Exporta las MISMAS hojas que el Excel (Eventos, Catalogos, Resumen),
  // reusando el mismo constructor; reemplaza el contenido en la planilla.
  const LS_SHEETS_URL = 'mp_sheets_url';
  const LS_SHEETS_AUTO = 'mp_sheets_auto';
  function sheetsUrl() { try { return (localStorage.getItem(LS_SHEETS_URL) || '').trim(); } catch (e) { return ''; } }
  function sheetsAuto() { try { return localStorage.getItem(LS_SHEETS_AUTO) === '1'; } catch (e) { return false; } }

  // Valor de celda ExcelJS -> primitivo para enviar (fecha -> dd-mm-aaaa)
  function cellVal(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v);
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(t => t.text).join('');
      if ('result' in v) return v.result == null ? '' : v.result;
      if ('text' in v) return v.text;
      return '';
    }
    return v;
  }
  function sheetToRows(ws) {
    const cols = ws.columnCount, rows = ws.rowCount, out = [];
    for (let r = 1; r <= rows; r++) {
      const row = ws.getRow(r), arr = [];
      for (let c = 1; c <= cols; c++) arr.push(cellVal(row.getCell(c).value));
      out.push(arr);
    }
    return out;
  }
  // Desplegables y coloreado para la hoja Eventos (los aplica el Apps Script)
  function eventosExtras() {
    return {
      validations: [
        { header: 'Resultado (R)', values: MP.RESULTADO_OPCIONES },
        { header: 'Estado Final del Equipo', values: MP.ESTADO_FINAL_OPCIONES },
        { header: 'Ejecutor', values: MP.EJECUTORES }
      ],
      colors: [
        // Estado Final colorea su propia celda (mayor prioridad que la fila)
        { header: 'Estado Final del Equipo', rules: [
          { value: 'Operativo', color: '#C6EFCE' },
          { value: 'No operativo', color: '#FFC7CE' },
          { value: 'En servicio técnico', color: '#FFE0B2' }
        ] },
        // Estado colorea TODA la fila (formato condicional por fórmula)
        { header: 'Estado', wholeRow: true, rules: [
          { value: 'Realizada', color: '#C6EFCE' },
          { value: 'Realizada (Año Anterior)', color: '#C6EFCE' },
          { value: 'Reprogramada', color: '#FCE4A6' },
          { value: 'Puesta en Marcha', color: '#DDEBF7' },
          { value: 'Fuera de Servicio', color: '#FFC7CE' },
          { value: 'No Realizada', color: '#FFC7CE' },
          { value: 'No Ubicable', color: '#E4DFEC' },
          { value: 'Baja', color: '#D9D9D9' },
          { value: 'Pendiente', color: '#FFF2CC' },
          { value: 'Pendiente (Reprogramada)', color: '#FFF2CC' },
          { value: 'Pendiente (Año Anterior)', color: '#FFF2CC' }
        ] }
      ]
    };
  }
  // Desplegables y coloreado para la hoja Correctivos
  function correctivosExtras() {
    return {
      validations: [
        { header: 'Tipo de Evento', values: MP.CORRECTIVO_TIPOS },
        { header: 'Ejecutor', values: MP.EJECUTORES },
        { header: 'Estado Final del Equipo', values: MP.ESTADO_FINAL_OPCIONES }
      ],
      colors: [
        { header: 'Estado Final del Equipo', rules: [
          { value: 'Operativo', color: '#C6EFCE' },
          { value: 'No operativo', color: '#FFC7CE' },
          { value: 'En servicio técnico', color: '#FFE0B2' }
        ] }
      ]
    };
  }
  // Construye el mismo libro del Excel y lo pasa a filas por hoja
  function buildSheetsPayload() {
    const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, consolidatedEvents(), { equipos: state.equipos.length, correctivos: state.correctivos });
    return {
      sheets: wb.worksheets.map(ws => {
        const o = { name: ws.name, rows: sheetToRows(ws) };
        if (ws.name === 'Eventos') Object.assign(o, eventosExtras());
        if (ws.name === 'Correctivos') Object.assign(o, correctivosExtras());
        return o;
      })
    };
  }

  // POST al Apps Script. text/plain evita el preflight CORS; si no se puede leer
  // la respuesta, reintenta en modo no-cors (envío sin confirmación).
  async function postSheets(payloadObj) {
    const url = sheetsUrl();
    if (!url) throw new Error('Falta la URL de la app web.');
    const body = JSON.stringify(payloadObj);
    try {
      const res = await fetch(url, {
        method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: body, redirect: 'follow'
      });
      const j = JSON.parse(await res.text());
      if (!j.ok) throw new Error(j.error || 'Error en el script');
      return j;
    } catch (e) {
      await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: body });
      return { ok: true, unconfirmed: true };
    }
  }

  async function exportToSheets(silent) {
    if (!sheetsUrl()) { if (!silent) showInline($('#sheetsStatus'), 'err', 'Primero pega y guarda la URL de la app web.'); return; }
    if (!state.equipos.length) { if (!silent) showInline($('#sheetsStatus'), 'err', 'Carga primero el archivo de programación.'); return; }
    const btn = $('#sheetsSendAll'); const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Enviando…';
    try {
      const payload = buildSheetsPayload();
      const j = await postSheets(payload);
      const names = payload.sheets.map(s => s.name).join(', ');
      showInline($('#sheetsStatus'), 'ok', j.unconfirmed
        ? '✓ Enviado a Google Sheets (sin confirmación; revisa la planilla). Hojas: ' + names + '.'
        : '✓ Hojas actualizadas: ' + (j.sheets ? j.sheets.join(', ') : names) + '.');
    } catch (e) {
      showInline($('#sheetsStatus'), 'err', '❌ ' + (e.message || e));
    } finally { btn.textContent = orig; btn.disabled = false; }
  }
  function autoExportSheets() { if (sheetsAuto() && sheetsUrl() && state.equipos.length) exportToSheets(true); }

  async function testSheets() {
    const url = sheetsUrl();
    if (!url) { showInline($('#sheetsStatus'), 'err', 'Primero pega y guarda la URL de la app web.'); return; }
    const btn = $('#sheetsTest'); const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Probando…';
    try {
      const res = await fetch(url, { method: 'GET', redirect: 'follow' });
      const j = JSON.parse(await res.text());
      showInline($('#sheetsStatus'), 'ok', '✓ Conexión correcta. Filas en "Eventos": ' + (j.count != null ? j.count : '—') + '.');
    } catch (e) {
      showInline($('#sheetsStatus'), 'warn', '⚠️ No se pudo confirmar la conexión (puede ser CORS). Aun así el envío suele funcionar; pulsa Enviar y revisa la planilla.');
    } finally { btn.textContent = orig; btn.disabled = false; }
  }

  // ---- Carga y parseo del archivo ----------------------------------------
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
      setStatus('❌ No se cargó la librería ExcelJS.', 'error');
      return;
    }
    $('#result').style.display = 'none';
    $('#download').disabled = true;
    setStatus('⏳ Procesando <b>' + esc(name) + '</b>… (puede tardar unos segundos)', 'info');
    await new Promise(r => setTimeout(r, 50));

    try {
      const buf = await readFile(file);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      if (!wb.getWorksheet('PMP_2026')) {
        throw new Error('La hoja "PMP_2026" no existe en el archivo. ¿Es la Programación MP 2026 correcta?');
      }
      const parsed = MP.parseWorkbook(wb);
      state.equipos = parsed.equipos;
      state.events = parsed.events;
      // Los registros previos se CONSERVAN (persistencia) y se comparan con el archivo.

      renderPreview();
      renderRegistry();
      renderCorrectivos();
      renderDiscrepancias();
      $('#download').disabled = false;
      $('#registrar').style.display = 'block';   // habilitar registro de mantenciones
      resetSearch();

      let msg = '✅ Procesado: <b>' + parsed.equipos.length.toLocaleString('es-CL') +
        '</b> equipos · <b>' + parsed.events.length.toLocaleString('es-CL') + '</b> eventos.';
      if (parsed.warnings.length) msg += '<br>⚠️ ' + parsed.warnings.map(esc).join('<br>⚠️ ');
      setStatus(msg, parsed.warnings.length ? 'warn' : 'ok');
    } catch (err) {
      console.error(err);
      setStatus('❌ Error al procesar: ' + esc(err.message || err), 'error');
    }
  }

  // Eventos del archivo + las mantenciones registradas superpuestas en la fila
  // del mismo equipo y mes (la última registrada prevalece). Se calcula al
  // momento de descargar/previsualizar, por lo que eliminar un registro se refleja.
  function consolidatedEvents() {
    if (!state.registros.length) return state.events;
    const out = state.events.map(e => Object.assign({}, e));
    const idxByKey = new Map();
    out.forEach((e, i) => idxByKey.set(keyOf(e), i));
    for (const reg of state.registros) {
      const key = keyOf(reg);
      const merged = Object.assign({}, reg);
      if (idxByKey.has(key)) {
        const i = idxByKey.get(key);
        if (!merged.observacion) merged.observacion = out[i].observacion; // conservar observación original
        out[i] = merged;
      } else {
        idxByKey.set(key, out.length);
        out.push(merged);
      }
    }
    out.sort((a, b) => (MP.toNum(a.id) - MP.toNum(b.id)) || (a.nMes - b.nMes));
    return out;
  }

  function renderPreview() {
    const ev = consolidatedEvents();
    const stats = MP.buildStats(ev);
    const estados = [...stats.byEstado.entries()].sort((a, b) => b[1] - a[1]);
    const chips = estados.map(([k, v]) =>
      '<span class="chip">' + esc(k) + ' <b>' + v.toLocaleString('es-CL') + '</b></span>').join('');
    const cols = ['ID', 'Equipo', 'Servicio', 'Mes', 'P', 'Tipo de Programación', 'R', 'Estado'];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    const rows = ev.slice(0, 20).map(e =>
      '<tr>' + [e.id, e.equipo, e.servicio, e.mes, e.programa, e.tipoPrograma, e.resultado, e.estado]
        .map(x => '<td>' + esc(x) + '</td>').join('') + '</tr>').join('');
    $('#result').innerHTML =
      '<div class="chips">' + chips + '</div>' +
      '<div class="muted" style="margin:10px 0 6px">Vista previa (primeros 20 de ' +
        ev.length.toLocaleString('es-CL') + ' eventos):</div>' +
      '<div class="tablewrap"><table class="prev">' + head + rows + '</table></div>';
    $('#result').style.display = 'block';
  }

  async function downloadWorkbook(events, fname, btn) {
    const original = btn.textContent;
    btn.disabled = true; btn.textContent = 'Generando Excel…';
    try {
      const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, events, { equipos: state.equipos.length, correctivos: state.correctivos });
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fname;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
    } catch (err) {
      console.error(err);
      setStatus('❌ Error al generar el Excel: ' + esc(err.message || err), 'error');
    } finally {
      btn.textContent = original; btn.disabled = false;
    }
  }

  // ---- Búsqueda con sugerencias en vivo + historial ----------------------
  function resetSearch() {
    $('#searchInput').value = '';
    hideSuggest();
    $('#equipoDetalle').style.display = 'none';
    $('#equipoDetalle').innerHTML = '';
    showInline($('#searchMsg'), '', '');
    state.selDetalleEq = null;
  }
  function hideSuggest() { const s = $('#suggest'); s.classList.remove('show'); s.innerHTML = ''; state.sgActive = -1; }

  // Sugerencias mientras se escribe (serie o inventario)
  function renderSuggest() {
    const q = $('#searchInput').value;
    if (!q.trim()) { hideSuggest(); return; }
    const results = MP.findEquipos(state.equipos, q).slice(0, 8);
    state.searchResults = results; state.sgActive = -1;
    const s = $('#suggest');
    if (!results.length) {
      s.innerHTML = '<div class="sg-empty">Sin coincidencias para "' + esc(q) + '".</div>';
      s.classList.add('show'); return;
    }
    s.innerHTML = results.map((e, i) =>
      '<div class="sg-item" data-idx="' + i + '"><b>' + esc(e.equipo) + '</b> · ' + esc(e.marca) + ' ' + esc(e.modelo) +
        '<div class="meta">Serie: ' + esc(e.serie || '—') + ' · Inv: ' + esc(e.inv || '—') + ' · ' + esc(e.servicio) + '</div></div>'
    ).join('');
    s.classList.add('show');
    s.querySelectorAll('.sg-item').forEach(it =>
      it.addEventListener('mousedown', ev => { ev.preventDefault(); selectEquipo(state.searchResults[+it.dataset.idx]); }));
  }
  function moveActive(d) {
    const items = $('#suggest').querySelectorAll('.sg-item');
    if (!items.length) return;
    state.sgActive = (state.sgActive + d + items.length) % items.length;
    items.forEach((it, i) => it.classList.toggle('active', i === state.sgActive));
  }

  // Al elegir un equipo: mostrar su historial + botón Registrar
  function selectEquipo(eq) {
    if (!eq) return;
    state.selDetalleEq = eq;
    $('#searchInput').value = (eq.serie && String(eq.serie).trim()) ? eq.serie : eq.inv;
    hideSuggest();
    showInline($('#searchMsg'), '', '');
    renderDetalle(eq);
    $('#equipoDetalle').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function renderDetalle(eq) {
    const hist = consolidatedEvents().filter(e => e.id === eq.id).sort((a, b) => a.nMes - b.nMes);
    const pm = MP.programmedMonths(eq);
    const val = v => (v !== null && v !== undefined && String(v).trim() !== '') ? esc(v) : '—';
    const fields = [
      ['Familia', eq.familia], ['ID', eq.id], ['N° Carpeta', eq.carpeta], ['N° Inventario', eq.inv],
      ['Equipo', eq.equipo], ['Servicio', eq.servicio], ['Unidad', eq.unidad], ['Ubicación', eq.ubicacion],
      ['Procedencia', eq.procedencia], ['Marca', eq.marca], ['Modelo', eq.modelo], ['N° Serie', eq.serie],
      ['Año Instalación', eq.anio], ['Vida Útil Residual', eq.vur], ['Clasificación', eq.clasif],
      ['ENU / Baja', eq.enubaja], ['Frecuencia MP', eq.frecuencia]
    ];
    const grid = '<div class="ficha">' + fields.map(f =>
      '<div class="fi"><span class="k">' + esc(f[0]) + '</span><span class="v">' + val(f[1]) + '</span></div>').join('') + '</div>';
    const obs = (eq.observacion && String(eq.observacion).trim())
      ? '<div class="fi-obs"><span class="k">Observación</span><span class="vv">' + esc(eq.observacion) + '</span></div>' : '';
    const head = '<div class="section-title" style="font-size:14px">📋 Datos del equipo</div>' + grid + obs;
    let body;
    if (!hist.length) {
      body = '<div class="muted">Este equipo no tiene meses programados ni resultados en 2026.</div>';
    } else {
      const cols = ['Mes', 'Programa', 'Tipo', 'Resultado', 'Fecha', 'Estado', 'Ejecutor', 'Estado Final'];
      const h = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
      const rows = hist.map(e =>
        '<tr><td>' + esc(e.mes) + '</td><td>' + esc(e.programa) + '</td><td>' + esc(e.tipoPrograma) + '</td><td>' + esc(e.resultado) +
        '</td><td>' + esc(fmtDate(e.fechaEjecucion)) + '</td><td>' + esc(e.estado) + '</td><td>' + esc(e.ejecutor) + '</td><td>' + esc(e.estadoFinal) + '</td></tr>'
      ).join('');
      body = '<div class="muted" style="margin:8px 0 4px">Historial 2026 (preventivo) — meses programados: ' +
        (pm.length ? pm.map(i => MP.MONTHS_FULL[i]).join(', ') : '(ninguno)') + '</div>' +
        '<div class="tablewrap"><table class="prev">' + h + rows + '</table></div>';
    }
    // Historial de eventos correctivos del equipo
    const corr = state.correctivos.filter(c => c.id === eq.id);
    let corrBody = '';
    if (corr.length) {
      const cc = ['Tipo de Evento', 'Fecha', 'Folio', 'N° Envío', 'Empresa', 'Ejecutor', 'Descripción', 'Estado Final'];
      const ch = '<tr>' + cc.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
      const cr = corr.map(c =>
        '<tr><td>' + esc(c.tipoEvento) + '</td><td>' + esc(fmtDate(c.fecha)) + '</td><td>' + esc(c.folioSolicitud || c.folioGuia) +
        '</td><td>' + esc(c.nEnvio) + '</td><td>' + esc(c.empresa) + '</td><td>' + esc(c.ejecutor) + '</td><td>' + esc(c.descripcion) +
        '</td><td>' + esc(c.estadoFinal) + '</td></tr>'
      ).join('');
      corrBody = '<div class="muted" style="margin:10px 0 4px">Eventos correctivos registrados</div>' +
        '<div class="tablewrap"><table class="prev">' + ch + cr + '</table></div>';
    }
    $('#equipoDetalle').innerHTML = head + body + corrBody +
      '<div class="actions" style="margin-top:10px">' +
        '<button id="detRegistrar" class="btn btn-accent btn-sm" type="button">🔧 Registrar mantención preventiva</button>' +
        '<button id="detCorrectivo" class="btn btn-primary btn-sm" type="button">🛠️ Registrar evento correctivo</button>' +
      '</div>';
    $('#equipoDetalle').style.display = 'block';
    $('#detRegistrar').addEventListener('click', () => openModal(eq));
    $('#detCorrectivo').addEventListener('click', () => openCModal(eq));
  }

  // Botón Buscar / Enter: selecciona la primera coincidencia
  function doSearch() {
    if (!$('#searchInput').value.trim()) { showInline($('#searchMsg'), 'err', 'Escriba un N° de Serie o N° de Inventario.'); return; }
    const results = MP.findEquipos(state.equipos, $('#searchInput').value);
    if (!results.length) { showInline($('#searchMsg'), 'err', 'Sin coincidencias.'); return; }
    selectEquipo(results[0]);
  }

  // ---- Modal de registro --------------------------------------------------
  function fillSelect(sel, items, fmt) {
    sel.innerHTML = '<option value="">— Seleccione —</option>' +
      items.map(v => '<option value="' + esc(v) + '">' + esc(fmt ? fmt(v) : v) + '</option>').join('');
  }
  function initSelects() {
    fillSelect($('#mResultado'), MP.RESULTADO_OPCIONES, code =>
      /^C[1-8]$/.test(code) ? code + ' — ' + MP.CAUSAL[code] : code + ' — ' + MP.decodeResultado(code));
    fillSelect($('#mEjecutor'), MP.EJECUTORES);
    fillSelect($('#mEstadoFinal'), MP.ESTADO_FINAL_OPCIONES);
    fillSelect($('#cTipo'), MP.CORRECTIVO_TIPOS);
  }

  function openModal(eq) {
    state.selEq = eq; state.selMonth = null;
    const pm = MP.programmedMonths(eq);
    $('#modalEq').innerHTML =
      '<b>' + esc(eq.equipo) + '</b> (ID ' + esc(eq.id) + ')<br>' +
      'Serie: <b>' + esc(eq.serie || '—') + '</b> · Inventario: <b>' + esc(eq.inv || '—') + '</b><br>' +
      esc(eq.servicio) + ' · ' + esc(eq.unidad) +
      '<br><span class="muted">Meses programados: ' +
        (pm.length ? pm.map(i => MP.MONTHS_FULL[i]).join(', ') : '(ninguno)') + '</span>';
    $('#mFecha').value = '';
    $('#mObs').value = '';
    $('#mResultado').value = ''; $('#mEjecutor').value = ''; $('#mEstadoFinal').value = '';
    updateEstadoFinal();
    $('#mTipo').style.display = 'none';
    $('#mFields').disabled = true;
    $('#mSave').disabled = true;
    showInline($('#mFechaMsg'), '', '');
    showInline($('#modalMsg'), '', '');
    $('#modal').classList.add('show');
    setTimeout(() => $('#mFecha').focus(), 50);
  }
  function closeModal() { $('#modal').classList.remove('show'); state.selEq = null; state.selMonth = null; }

  function onDateChange() {
    $('#mTipo').style.display = 'none';
    $('#mFields').disabled = true;
    state.selMonth = null;
    const v = $('#mFecha').value;
    if (!v) { showInline($('#mFechaMsg'), '', ''); updateSave(); return; }
    const mi = parseInt(v.split('-')[1], 10) - 1;
    const eq = state.selEq;
    const pm = MP.programmedMonths(eq);
    if (!pm.includes(mi)) {
      const meses = pm.length ? pm.map(i => MP.MONTHS_FULL[i]).join(', ') : '(ninguno)';
      showInline($('#mFechaMsg'), 'err',
        '✖ <b>' + MP.MONTHS_FULL[mi] + '</b> no tiene mantención programada para este equipo. No es posible avanzar.<br>Meses programados: ' + meses + '.');
      updateSave(); return;
    }
    state.selMonth = mi;
    const code = eq.prog[mi];
    showInline($('#mFechaMsg'), 'ok', '✔ Mes válido: <b>' + MP.MONTHS_FULL[mi] + '</b>.');
    $('#mTipo').style.display = 'block';
    $('#mTipo').innerHTML = 'Tipo de programa: <b>' + esc(code) + '</b> — ' + esc(MP.PROG[code] || code);
    $('#mFields').disabled = false;
    updateEstadoFinal();
    updateSave();
  }

  // El "Estado final del equipo" solo se solicita cuando el Resultado es "Si";
  // El Estado Final siempre se puede elegir (Operativo / No operativo / En servicio técnico).
  function updateEstadoFinal() {
    $('#mEstadoFinal').disabled = false;
    $('#mEstadoFinalLabel').innerHTML = 'Estado final del equipo <span class="req">*</span>';
  }

  function updateSave() {
    $('#mSave').disabled = !(state.selMonth !== null &&
      $('#mResultado').value && $('#mEjecutor').value && $('#mEstadoFinal').value);
  }

  function saveRegistro() {
    const eq = state.selEq, mi = state.selMonth;
    if (!eq || mi === null) return;
    if (!$('#mResultado').value || !$('#mEjecutor').value || !$('#mEstadoFinal').value) {
      showInline($('#modalMsg'), 'err', 'Complete los campos obligatorios (*).'); return;
    }
    const p = $('#mFecha').value.split('-');
    const fecha = new Date(+p[0], +p[1] - 1, +p[2]);
    const reg = MP.buildRegistro(eq, mi, {
      fecha,
      resultado: $('#mResultado').value,
      observacion: $('#mObs').value.trim(),
      ejecutor: $('#mEjecutor').value,
      estadoFinal: $('#mEstadoFinal').value
    });
    state.registros.push(reg);
    saveRegistros();   // persistir en el navegador
    closeModal();
    renderRegistry();
    renderPreview();   // reflejar el registro en la vista previa y en el Excel masivo
    renderDiscrepancias();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);   // refrescar historial visible
    setStatus('✅ Mantención registrada: <b>' + esc(eq.equipo) + '</b> — ' + esc(reg.mes) +
      ' (' + esc(fmtDate(fecha)) + '). Quedará incluida al descargar el Excel.', 'ok');
    autoExportSheets();   // actualizar Google Sheets si está activado
  }

  // ---- Comparación: registrado vs. archivo cargado -----------------------
  function compareRegistros() {
    const map = new Map();
    state.events.forEach(e => { const k = keyOf(e); if (!map.has(k)) map.set(k, e); });
    const diffs = []; let iguales = 0;
    state.registros.forEach(r => {
      const ev = map.get(keyOf(r));
      const regRes = r.resultado || '';
      if (!ev) { diffs.push({ r: r, tipo: 'nf' }); return; }
      const fileRes = ev.resultado || '';
      if (fileRes === regRes) { iguales++; return; }
      diffs.push({ r: r, tipo: fileRes ? 'dif' : 'vacio', archivo: fileRes });
    });
    return { diffs, iguales };
  }

  function renderDiscrepancias() {
    const card = $('#discrepanciasCard'), box = $('#discrepancias');
    if (!state.registros.length || !state.events.length) { card.style.display = 'none'; return; }
    const { diffs, iguales } = compareRegistros();
    card.style.display = 'block';
    if (!diffs.length) {
      box.innerHTML = '<div class="inline-msg ok show">✔ Tus ' + state.registros.length +
        ' mantención(es) registradas coinciden con el archivo cargado.</div>';
      return;
    }
    const rows = diffs.map(d => {
      let txt, cls;
      if (d.tipo === 'nf') { txt = 'No encontrado en el archivo'; cls = 'warn'; }
      else if (d.tipo === 'vacio') { txt = '<i>(vacío en el archivo)</i>'; cls = 'warn'; }
      else { txt = '<b>' + esc(d.archivo) + '</b>'; cls = 'err'; }
      return '<tr class="' + cls + '"><td>' + esc(d.r.id) + '</td><td>' + esc(d.r.equipo) +
        '</td><td>' + esc(d.r.serie || d.r.inv) + '</td><td>' + esc(d.r.mes) +
        '</td><td><b>' + esc(d.r.resultado) + '</b></td><td>' + txt + '</td></tr>';
    }).join('');
    box.innerHTML =
      '<div class="disc-head">⚠️ ' + diffs.length + ' diferencia(s) entre lo registrado y el archivo' +
        (iguales ? ' · ' + iguales + ' coincidente(s)' : '') + '</div>' +
      '<div class="muted" style="margin-bottom:8px">Al descargar, <b>prevalece lo registrado</b> en el programa. Revise si debe corregir el registro o el archivo.</div>' +
      '<div class="tablewrap"><table class="prev"><tr><th>ID</th><th>Equipo</th><th>Serie/Inv</th>' +
        '<th>Mes</th><th>Registrado</th><th>En el archivo</th></tr>' + rows + '</table></div>';
  }

  // ---- Tabla de registros -------------------------------------------------
  function renderRegistry() {
    const wrap = $('#registryWrap');
    if (!state.registros.length) { wrap.style.display = 'none'; return; }
    wrap.style.display = 'block';
    const term = ($('#registryFilter') ? $('#registryFilter').value : '').trim().toLowerCase();
    const cols = ['#', 'ID', 'Equipo', 'Serie / Inv', 'Mes', 'Fecha', 'Programa', 'Resultado', 'Ejecutor', 'Estado Final', ''];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let shown = 0;
    const rows = state.registros.map((e, i) => {
      const hay = [e.id, e.equipo, e.serie, e.inv, e.mes, fmtDate(e.fechaEjecucion), e.programa, e.resultado, e.ejecutor, e.estadoFinal, e.estado].join(' ').toLowerCase();
      if (term && hay.indexOf(term) === -1) return '';
      shown++;
      return '<tr>' +
        '<td>' + (i + 1) + '</td>' +
        '<td>' + esc(e.id) + '</td>' +
        '<td>' + esc(e.equipo) + '</td>' +
        '<td>' + esc(e.serie || e.inv) + '</td>' +
        '<td>' + esc(e.mes) + '</td>' +
        '<td>' + esc(fmtDate(e.fechaEjecucion)) + '</td>' +
        '<td>' + esc(e.programa) + '</td>' +
        '<td>' + esc(e.resultado) + '</td>' +
        '<td>' + esc(e.ejecutor) + '</td>' +
        '<td>' + esc(e.estadoFinal) + '</td>' +
        '<td><button class="btn btn-del" title="Eliminar" data-del="' + i + '">✕</button></td>' +
      '</tr>';
    }).join('');
    $('#registryCount').textContent = term ? (shown + ' / ' + state.registros.length) : state.registros.length;
    const t = $('#registryTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="11" class="nomatch">Sin coincidencias para el filtro.</td></tr>');
    t.querySelectorAll('button[data-del]').forEach(b =>
      b.addEventListener('click', () => {
        state.registros.splice(+b.dataset.del, 1);
        saveRegistros(); renderRegistry(); renderPreview(); renderDiscrepancias();
        if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
      }));
  }

  // ---- Eventos correctivos: tabla -----------------------------------------
  function renderCorrectivos() {
    const wrap = $('#correctivosWrap');
    if (!state.correctivos.length) { wrap.style.display = 'none'; return; }
    wrap.style.display = 'block';
    const term = ($('#correctivosFilter') ? $('#correctivosFilter').value : '').trim().toLowerCase();
    const cols = ['#', 'ID', 'Equipo', 'Tipo de Evento', 'Fecha', 'Folio', 'N° Envío', 'Empresa', 'Ejecutor', 'Estado Final', ''];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let shown = 0;
    const rows = state.correctivos.map((c, i) => {
      const hay = [c.id, c.equipo, c.serie, c.inv, c.tipoEvento, fmtDate(c.fecha), c.folioSolicitud, c.folioGuia, c.nEnvio, c.empresa, c.ejecutor, c.descripcion, c.estadoFinal].join(' ').toLowerCase();
      if (term && hay.indexOf(term) === -1) return '';
      shown++;
      return '<tr><td>' + (i + 1) + '</td><td>' + esc(c.id) + '</td><td>' + esc(c.equipo) + '</td><td>' + esc(c.tipoEvento) +
        '</td><td>' + esc(fmtDate(c.fecha)) + '</td><td>' + esc(c.folioSolicitud || c.folioGuia) + '</td><td>' + esc(c.nEnvio) +
        '</td><td>' + esc(c.empresa) + '</td><td>' + esc(c.ejecutor) + '</td><td>' + esc(c.estadoFinal) +
        '</td><td><button class="btn btn-del" title="Eliminar" data-del="' + i + '">✕</button></td></tr>';
    }).join('');
    $('#correctivosCount').textContent = term ? (shown + ' / ' + state.correctivos.length) : state.correctivos.length;
    const t = $('#correctivosTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="11" class="nomatch">Sin coincidencias para el filtro.</td></tr>');
    t.querySelectorAll('button[data-del]').forEach(b =>
      b.addEventListener('click', () => {
        state.correctivos.splice(+b.dataset.del, 1);
        saveCorrectivos(); renderCorrectivos();
        if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
      }));
  }

  // ---- Modal de evento correctivo -----------------------------------------
  const CORR_OPCIONALES = new Set(['descripcion']);   // campos no obligatorios
  function cField(k, label, inner) {
    const tag = CORR_OPCIONALES.has(k) ? '<span class="muted">(opcional)</span>' : '<span class="req">*</span>';
    return '<div class="field"><label for="cf_' + k + '">' + esc(label) + ' ' + tag + '</label>' + inner + '</div>';
  }
  function cSelectHTML(id, items) {
    return '<select id="' + id + '"><option value="">— Seleccione —</option>' +
      items.map(v => '<option value="' + esc(v) + '">' + esc(v) + '</option>').join('') + '</select>';
  }
  function renderCFields() {
    const tipo = $('#cTipo').value;
    const cont = $('#cFields');
    if (!tipo) { cont.innerHTML = ''; updateCSave(); return; }
    const keys = MP.CORRECTIVO_CAMPOS[tipo] || [];
    cont.innerHTML = keys.map(k => {
      const label = MP.CORRECTIVO_LABELS[k] || k;
      if (k === 'ejecutor') return cField(k, label, cSelectHTML('cf_' + k, MP.EJECUTORES));
      if (k === 'estadoFinal') return cField(k, label, cSelectHTML('cf_' + k, MP.ESTADO_FINAL_OPCIONES));
      if (k === 'descripcion') return cField(k, label, '<textarea id="cf_' + k + '"></textarea>');
      if (k === 'fecha') return cField(k, label, '<input id="cf_' + k + '" type="date" min="2026-01-01" max="2026-12-31">');
      return cField(k, label, '<input id="cf_' + k + '" type="text" autocomplete="off">');
    }).join('');
    cont.querySelectorAll('input,select,textarea').forEach(el => {
      el.addEventListener('input', updateCSave); el.addEventListener('change', updateCSave);
    });
    updateCSave();
  }
  function updateCSave() {
    const tipo = $('#cTipo').value;
    let ok = !!tipo && !!state.cEq;
    if (ok) (MP.CORRECTIVO_CAMPOS[tipo] || []).forEach(k => {
      if (CORR_OPCIONALES.has(k)) return;
      const el = $('#cf_' + k); if (!el || !String(el.value).trim()) ok = false;
    });
    $('#cSave').disabled = !ok;
  }
  function openCModal(eq) {
    state.cEq = eq;
    $('#cmodalEq').innerHTML = '<b>' + esc(eq.equipo) + '</b> (ID ' + esc(eq.id) + ') · Serie: <b>' +
      esc(eq.serie || '—') + '</b> · Inv: <b>' + esc(eq.inv || '—') + '</b>';
    $('#cTipo').value = '';
    $('#cFields').innerHTML = '';
    showInline($('#cmodalMsg'), '', '');
    $('#cSave').disabled = true;
    $('#cmodal').classList.add('show');
  }
  function closeCModal() { $('#cmodal').classList.remove('show'); state.cEq = null; }
  function saveCorrectivo() {
    const eq = state.cEq, tipo = $('#cTipo').value;
    if (!eq || !tipo) return;
    const data = { tipoEvento: tipo };
    const keys = MP.CORRECTIVO_CAMPOS[tipo] || [];
    for (const k of keys) {
      const el = $('#cf_' + k); let v = el ? el.value : '';
      if (k === 'fecha' && v) { const p = v.split('-'); v = new Date(+p[0], +p[1] - 1, +p[2]); }
      else v = String(v).trim();
      if (!v && !CORR_OPCIONALES.has(k)) { showInline($('#cmodalMsg'), 'err', 'Complete los campos obligatorios (*).'); return; }
      data[k] = v;
    }
    state.correctivos.push(MP.buildCorrectivo(eq, data));
    saveCorrectivos();
    closeCModal();
    renderCorrectivos();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
    setStatus('✅ Evento correctivo registrado: <b>' + esc(eq.equipo) + '</b> — ' + esc(tipo) + '.', 'ok');
    autoExportSheets();
  }

  // ---- Conexión de eventos de UI -----------------------------------------
  $('#pick').addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';   // permite volver a subir el MISMO archivo (corregido) y que se dispare el cambio
    if (f) { $('#fileName').textContent = f.name; handleFile(f); }
  });
  $('#download').addEventListener('click', () =>
    downloadWorkbook(consolidatedEvents(), 'Eventos_MP_2026_' + stamp() + '.xlsx', $('#download')));
  $('#downloadReg').addEventListener('click', () => {
    if (!state.registros.length) return;
    downloadWorkbook(state.registros, 'Registro_Mantenciones_' + stamp() + '.xlsx', $('#downloadReg'));
  });
  $('#clearReg').addEventListener('click', () => {
    if (!state.registros.length) return;
    if (!confirm('¿Eliminar las ' + state.registros.length + ' mantenciones registradas? Esta acción no se puede deshacer.')) return;
    state.registros = [];
    saveRegistros(); renderRegistry(); renderPreview(); renderDiscrepancias();
    setStatus('Registros eliminados.', 'info');
  });
  $('#registryFilter').addEventListener('input', renderRegistry);
  $('#correctivosFilter').addEventListener('input', renderCorrectivos);

  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) { $('#fileName').textContent = f.name; handleFile(f); }
  });

  $('#searchBtn').addEventListener('click', doSearch);
  let sgTimer = null;
  $('#searchInput').addEventListener('input', () => { clearTimeout(sgTimer); sgTimer = setTimeout(renderSuggest, 120); });
  $('#searchInput').addEventListener('focus', () => { if ($('#searchInput').value.trim()) renderSuggest(); });
  $('#searchInput').addEventListener('keydown', e => {
    const open = $('#suggest').classList.contains('show');
    if (e.key === 'ArrowDown') { e.preventDefault(); if (!open) renderSuggest(); else moveActive(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); moveActive(-1); }
    else if (e.key === 'Enter') {
      e.preventDefault();
      if (open && state.sgActive >= 0 && state.searchResults[state.sgActive]) selectEquipo(state.searchResults[state.sgActive]);
      else doSearch();
    } else if (e.key === 'Escape') { hideSuggest(); }
  });
  document.addEventListener('click', e => {
    if (!$('#suggest').contains(e.target) && e.target !== $('#searchInput')) hideSuggest();
  });

  $('#mFecha').addEventListener('change', onDateChange);
  $('#mFecha').addEventListener('input', onDateChange);
  $('#mResultado').addEventListener('change', () => { updateEstadoFinal(); updateSave(); });
  ['#mEjecutor', '#mEstadoFinal'].forEach(s => $(s).addEventListener('change', updateSave));
  $('#mSave').addEventListener('click', saveRegistro);
  $('#mCancel').addEventListener('click', closeModal);
  $('#modal').addEventListener('click', e => { if (e.target === $('#modal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#modal').classList.contains('show')) closeModal(); });

  // Evento correctivo
  $('#cTipo').addEventListener('change', renderCFields);
  $('#cSave').addEventListener('click', saveCorrectivo);
  $('#cCancel').addEventListener('click', closeCModal);
  $('#cmodal').addEventListener('click', e => { if (e.target === $('#cmodal')) closeCModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#cmodal').classList.contains('show')) closeCModal(); });
  $('#clearCorr').addEventListener('click', () => {
    if (!state.correctivos.length) return;
    if (!confirm('¿Eliminar los ' + state.correctivos.length + ' eventos correctivos registrados?')) return;
    state.correctivos = []; saveCorrectivos(); renderCorrectivos();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
  });

  // Google Sheets
  $('#sheetsSave').addEventListener('click', () => {
    const u = $('#sheetsUrl').value.trim();
    try { localStorage.setItem(LS_SHEETS_URL, u); } catch (e) {}
    showInline($('#sheetsStatus'), u ? 'ok' : 'err', u ? '✓ URL guardada.' : 'URL vacía.');
  });
  $('#sheetsTest').addEventListener('click', testSheets);
  $('#sheetsSendAll').addEventListener('click', () => exportToSheets(false));
  $('#sheetsAuto').addEventListener('change', e => {
    try { localStorage.setItem(LS_SHEETS_AUTO, e.target.checked ? '1' : '0'); } catch (err) {}
    showInline($('#sheetsStatus'), 'info', e.target.checked
      ? 'Guardado automático activado.' : 'Guardado automático desactivado.');
  });
  $('#copyScript').addEventListener('click', async () => {
    const txt = $('#gsScript').textContent;
    try { await navigator.clipboard.writeText(txt); showInline($('#sheetsStatus'), 'ok', '✓ Script copiado al portapapeles.'); }
    catch (e) {
      const r = document.createRange(); r.selectNodeContents($('#gsScript'));
      const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
      showInline($('#sheetsStatus'), 'info', 'Selecciona y copia el script (Ctrl+C).');
    }
  });

  initSelects();
  // Restaurar configuración de Google Sheets
  $('#sheetsUrl').value = sheetsUrl();
  $('#sheetsAuto').checked = sheetsAuto();

  // Recuperar registros guardados en este navegador (persistencia entre sesiones)
  state.registros = loadRegistros();
  state.correctivos = loadCorrectivos();
  renderCorrectivos();
  if (state.registros.length || state.correctivos.length) {
    setStatus('ℹ️ Tienes <b>' + state.registros.length + '</b> mantención(es) y <b>' + state.correctivos.length +
      '</b> evento(s) correctivo(s) guardados en este navegador. Cargue el archivo para incluirlos al descargar.', 'info');
  }
})();
