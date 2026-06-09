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
  const state = { equipos: [], events: [], registros: [], correctivos: [], pendientes: [], archivos: [], searchResults: [], selEq: null, selMonth: null, selDetalleEq: null, sgActive: -1, cEq: null, pEq: null, gPend: null, invFilter: 'todos', resSel: null };

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
  let statusTimer = null;
  function setStatus(msg, kind) {
    const s = $('#status');
    if (statusTimer) { clearTimeout(statusTimer); statusTimer = null; }
    s.className = 'status ' + (kind || '');
    s.innerHTML = msg;
    s.style.display = msg ? 'block' : 'none';
    if (msg && kind === 'ok') statusTimer = setTimeout(() => { s.style.display = 'none'; }, 6000);
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

  // ---- Persistencia de pendientes ----------------------------------------
  const LS_PEND = 'mp_pendientes_2026';
  function savePendientes() {
    try {
      const data = state.pendientes.map(p => {
        const o = Object.assign({}, p);
        o.fechaCompromiso = (p.fechaCompromiso instanceof Date) ? isoDate(p.fechaCompromiso) : (p.fechaCompromiso || '');
        o.actualizaciones = (p.actualizaciones || []).map(a => ({ fecha: (a.fecha instanceof Date) ? isoDate(a.fecha) : (a.fecha || ''), texto: a.texto || '' }));
        o.tareas = (p.tareas || []).map(t => ({ texto: t.texto || '', hecha: !!t.hecha }));
        return o;
      });
      localStorage.setItem(LS_PEND, JSON.stringify(data));
    } catch (e) { /* localStorage no disponible */ }
  }
  function reviveDate(s) { if (typeof s === 'string' && s) { const x = s.split('-'); return new Date(+x[0], +x[1] - 1, +x[2]); } return ''; }
  function loadPendientes() {
    try {
      const raw = localStorage.getItem(LS_PEND);
      if (!raw) return [];
      return JSON.parse(raw).map(o => {
        const p = Object.assign({}, o);
        p.fechaCompromiso = reviveDate(p.fechaCompromiso);
        p.estado = p.estado || 'Abierto';
        p.actualizaciones = (p.actualizaciones || []).map(a => ({ fecha: reviveDate(a.fecha), texto: a.texto || '' }));
        p.tareas = (p.tareas || []).map(t => ({ texto: t.texto || '', hecha: !!t.hecha }));
        return p;
      });
    } catch (e) { return []; }
  }

  // ---- Persistencia de archivos (enlaces a Drive) ------------------------
  const LS_ARCH = 'mp_archivos_2026';
  function saveArchivos() { try { localStorage.setItem(LS_ARCH, JSON.stringify(state.archivos)); } catch (e) { } }
  function loadArchivos() { try { const r = localStorage.getItem(LS_ARCH); return r ? JSON.parse(r) : []; } catch (e) { return []; } }

  // ---- Integración con Google Sheets (Apps Script) -----------------------
  // Exporta las MISMAS hojas que el Excel (Eventos, Catalogos, Resumen),
  // reusando el mismo constructor; reemplaza el contenido en la planilla.
  const LS_SHEETS_URL = 'mp_sheets_url';
  const LS_SHEETS_AUTO = 'mp_sheets_auto';
  function sheetsUrl() { try { return (localStorage.getItem(LS_SHEETS_URL) || '').trim(); } catch (e) { return ''; } }
  function sheetsAuto() { try { return localStorage.getItem(LS_SHEETS_AUTO) === '1'; } catch (e) { return false; } }
  // Si la app se sirve DESDE Apps Script, usamos google.script.run (sin URL ni CORS)
  const GAS = (typeof google !== 'undefined' && google.script && google.script.run);
  function gasCall(fn, arg) {
    return new Promise((resolve, reject) => {
      const r = google.script.run.withSuccessHandler(resolve).withFailureHandler(reject);
      (arg === undefined) ? r[fn]() : r[fn](arg);
    });
  }
  function sheetsReady() { return GAS || !!sheetsUrl(); }

  // Valor de celda ExcelJS -> primitivo para enviar (fecha -> dd-mm-aaaa)
  function cellVal(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return fmtDate(v);
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(t => t.text).join('');
      if ('hyperlink' in v) return v.hyperlink;   // celda con enlace -> URL (clicable en Sheets)
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
  // Desplegables para la hoja Pendientes (responsables = ejecutores)
  function pendientesExtras() {
    return {
      validations: [
        { header: 'Tipo', values: MP.PENDIENTE_TIPOS },
        { header: 'Responsable Administrativo', values: MP.EJECUTORES },
        { header: 'Responsable de Ejecución', values: MP.EJECUTORES },
        { header: 'Estado', values: MP.ESTADO_PENDIENTE_OPCIONES }
      ],
      colors: [
        { header: 'Estado', rules: [
          { value: 'Abierto', color: '#FFF2CC' },
          { value: 'En progreso', color: '#DDEBF7' },
          { value: 'Resuelto', color: '#C6EFCE' }
        ] }
      ]
    };
  }
  // Construye el mismo libro del Excel y lo pasa a filas por hoja
  function buildSheetsPayload() {
    const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, consolidatedEvents(), { equipos: state.equipos.length, correctivos: state.correctivos, pendientes: decoratedPendientes(), archivos: state.archivos });
    return {
      // La hoja "Archivos" la gestiona el propio Apps Script al subir cada
      // archivo; no la sobrescribimos en el envío masivo para no perder enlaces.
      sheets: wb.worksheets.filter(ws => ws.name !== 'Archivos').map(ws => {
        const o = { name: ws.name, rows: sheetToRows(ws) };
        if (ws.name === 'Eventos') Object.assign(o, eventosExtras());
        if (ws.name === 'Correctivos') Object.assign(o, correctivosExtras());
        if (ws.name === 'Pendientes') Object.assign(o, pendientesExtras());
        return o;
      }),
      data: serializeState(),   // snapshot para poder LEER al reabrir
      equipos: state.equipos.length ? equiposToRows() : null   // inventario, para auto-cargar al abrir
    };
  }

  // ---- Inventario de equipos: ida y vuelta con Google Sheets (_equipos) ---
  // Permite que al abrir el link se restaure el inventario sin volver a subir el Excel.
  const EQ_COLS = ['id', 'familia', 'carpeta', 'inv', 'equipo', 'servicio', 'unidad', 'ubicacion', 'procedencia', 'marca', 'modelo', 'serie', 'anio', 'vur', 'clasif', 'enubaja', 'observacion', 'frecuencia'];
  function equiposToRows() {
    const rows = [EQ_COLS.concat(['prog', 'res'])];
    state.equipos.forEach(eq => {
      const r = EQ_COLS.map(k => eq[k] == null ? '' : String(eq[k]));
      r.push((eq.prog || []).join('|')); r.push((eq.res || []).join('|'));
      rows.push(r);
    });
    return rows;
  }
  function equiposFromRows(grid) {
    if (!grid || grid.length < 2) return [];
    const h = grid[0].map(x => String(x == null ? '' : x).trim());
    const idx = {}; h.forEach((k, i) => idx[k] = i);
    const get = (r, k) => { const i = idx[k]; const v = (i == null ? '' : r[i]); return v == null ? '' : String(v).trim(); };
    const split12 = s => { let a = (s == null ? '' : String(s)).split('|'); if (a.length === 1 && a[0] === '') a = []; while (a.length < 12) a.push(''); return a; };
    return grid.slice(1).map(r => {
      const eq = {}; EQ_COLS.forEach(k => eq[k] = get(r, k));
      eq.prog = split12(get(r, 'prog')); eq.res = split12(get(r, 'res'));
      return eq;
    }).filter(eq => eq.id || eq.equipo);
  }
  // Reconstruye state.events a partir de los equipos (igual que parseWorkbook)
  function rebuildEventsFromEquipos() {
    const events = [];
    state.equipos.forEach(eq => {
      for (let i = 0; i < 12; i++) {
        const p = eq.prog[i] || '', r = eq.res[i] || '';
        if (MP.isEmpty(p) && MP.isEmpty(r)) continue;
        events.push(MP.makeEvent(eq, i, p, r));
      }
    });
    events.sort((a, b) => (MP.toNum(a.id) - MP.toNum(b.id)) || (a.nMes - b.nMes));
    state.events = events;
  }
  // Carga un inventario restaurado y deja la app lista (preview, registro, vistas)
  function applyEquipos(equipos) {
    state.equipos = equipos;
    rebuildEventsFromEquipos();
    renderPreview(); renderRegistry(); renderCorrectivos(); renderPendientes(); renderDiscrepancias();
    $('#download').disabled = false;
    $('#registrar').style.display = 'block';
    $('#registrarHint').style.display = 'none';
  }

  // Snapshot serializable de lo registrado (fechas -> ISO)
  function serializeState() {
    const sd = d => (d instanceof Date) ? isoDate(d) : (d || '');
    return {
      registros: state.registros.map(r => Object.assign({}, r, { fechaEjecucion: sd(r.fechaEjecucion) })),
      correctivos: state.correctivos.map(c => Object.assign({}, c, { fecha: sd(c.fecha) })),
      pendientes: state.pendientes.map(p => Object.assign({}, p, {
        fechaCompromiso: sd(p.fechaCompromiso),
        actualizaciones: (p.actualizaciones || []).map(a => ({ fecha: sd(a.fecha), texto: a.texto || '' })),
        tareas: (p.tareas || []).map(t => ({ texto: t.texto || '', hecha: !!t.hecha }))
      }))
    };
  }
  // Revive un snapshot y reemplaza el estado local
  function restoreState(obj) {
    if (!obj) return false;
    let any = false;
    if (Array.isArray(obj.registros)) {
      state.registros = obj.registros.map(o => Object.assign({}, o, { fechaEjecucion: reviveDate(o.fechaEjecucion) })); any = true;
    }
    if (Array.isArray(obj.correctivos)) {
      state.correctivos = obj.correctivos.map(o => Object.assign({}, o, { fecha: reviveDate(o.fecha) })); any = true;
    }
    if (Array.isArray(obj.pendientes)) {
      state.pendientes = obj.pendientes.map(o => Object.assign({}, o, {
        fechaCompromiso: reviveDate(o.fechaCompromiso),
        estado: o.estado || 'Abierto',
        actualizaciones: (o.actualizaciones || []).map(a => ({ fecha: reviveDate(a.fecha), texto: a.texto || '' })),
        tareas: (o.tareas || []).map(t => ({ texto: t.texto || '', hecha: !!t.hecha }))
      })); any = true;
    }
    return any;
  }

  // Reconstrucción desde las hojas visibles (cuando no hay snapshot _datos)
  function asDate(v) {
    if (v instanceof Date) return v;
    if (typeof v === 'string') {
      let m = v.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
      m = v.match(/^(\d{2})-(\d{2})-(\d{4})/); if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    }
    return '';
  }
  function rowsToObjs(grid) {
    if (!grid || grid.length < 2) return [];
    const h = grid[0].map(x => String(x == null ? '' : x).trim());
    return grid.slice(1).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
  }
  function sid(v) { return v == null ? '' : String(typeof v === 'number' ? Math.round(v) : v).trim(); }
  function reconstructFromTablas(tablas) {
    const corr = rowsToObjs(tablas.Correctivos).map(o => ({
      id: sid(o['ID']), inv: sid(o['N° Inventario']), serie: sid(o['N° Serie']), equipo: o['Equipo'] || '',
      servicio: o['Servicio'] || '', unidad: o['Unidad'] || '', tipoEvento: o['Tipo de Evento'] || '', fecha: asDate(o['Fecha']),
      folioSolicitud: sid(o['Folio Solicitud']), nEnvio: sid(o['N° Envío']), folioGuia: sid(o['Folio Guía Despacho']),
      empresa: o['Empresa'] || '', ejecutor: o['Ejecutor'] || '', descripcion: o['Descripción'] || '', estadoFinal: o['Estado Final del Equipo'] || ''
    })).filter(c => c.id || c.equipo);
    const pend = rowsToObjs(tablas.Pendientes).map(o => ({
      id: sid(o['ID']), inv: sid(o['N° Inventario']), serie: sid(o['N° Serie']), equipo: o['Equipo'] || '',
      servicio: o['Servicio'] || '', unidad: o['Unidad'] || '', tipo: o['Tipo'] || '', fechaCompromiso: asDate(o['Fecha de Compromiso']),
      respAdministrativo: o['Responsable Administrativo'] || '', respEjecucion: o['Responsable de Ejecución'] || '',
      observacion: o['Observación'] || '', estado: o['Estado'] || 'Abierto', tareas: [], actualizaciones: []
    })).filter(p => p.id || p.equipo);
    // Tareas y bitácora por ID (solo si hay un único pendiente por equipo, para no mezclar)
    const cnt = {}; pend.forEach(p => cnt[p.id] = (cnt[p.id] || 0) + 1);
    const tById = {}, bById = {};
    rowsToObjs(tablas.Tareas).forEach(o => { const id = sid(o['ID']); (tById[id] = tById[id] || []).push({ texto: o['Tarea'] || '', hecha: /^s[ií]$/i.test(String(o['Hecha']).trim()) }); });
    rowsToObjs(tablas.Bitacora).forEach(o => { const id = sid(o['ID']); (bById[id] = bById[id] || []).push({ fecha: asDate(o['Fecha']), texto: o['Actualización'] || '' }); });
    pend.forEach(p => { if (cnt[p.id] === 1) { p.tareas = tById[p.id] || []; p.actualizaciones = bById[p.id] || []; } });
    return { correctivos: corr, pendientes: pend };
  }
  // Archivos (enlaces a Drive) leídos de la hoja "Archivos"
  function archivosFromTabla(grid) {
    return rowsToObjs(grid).map(o => ({
      id: sid(o['ID']), inv: sid(o['N° Inventario']), serie: sid(o['N° Serie']),
      equipo: o['Equipo'] || '', servicio: o['Servicio'] || '', categoria: o['Categoría'] || '',
      descripcion: o['Descripción'] || '',
      nombre: o['Nombre del archivo'] || '', enlace: o['Enlace'] || '', fecha: o['Fecha de carga'] || ''
    })).filter(a => a.enlace);
  }

  // POST al Apps Script. text/plain evita el preflight CORS; si no se puede leer
  // la respuesta, reintenta en modo no-cors (envío sin confirmación).
  async function postSheets(payloadObj) {
    if (GAS) {
      const j = await gasCall('appPush', payloadObj);
      if (!j || !j.ok) throw new Error((j && j.error) || 'Error en el script');
      return j;
    }
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

  // Lee un archivo como base64 (sin el prefijo data:)
  function readFileB64(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => { const s = String(fr.result); const i = s.indexOf(','); resolve(i >= 0 ? s.slice(i + 1) : s); };
      fr.onerror = () => reject(new Error('No se pudo leer el archivo.'));
      fr.readAsDataURL(file);
    });
  }
  // Sube un archivo a Drive (vía Apps Script) y registra el enlace en la hoja "Archivos"
  async function uploadArchivoFile(eq, file, categoria, descripcion) {
    const dataBase64 = await readFileB64(file);
    const payload = {
      action: 'upload', equipoId: eq.id, equipo: eq.equipo, inv: eq.inv, serie: eq.serie,
      servicio: eq.servicio, categoria: categoria || '', descripcion: descripcion || '', nombre: file.name, mime: file.type || 'application/octet-stream', dataBase64
    };
    let j;
    if (GAS) { j = await gasCall('appUpload', payload); }
    else {
      const url = sheetsUrl();
      if (!url) throw new Error('Conecta Google Sheets (Apps Script) para subir archivos.');
      const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload), redirect: 'follow' });
      j = JSON.parse(await res.text());
    }
    if (!j || !j.ok) throw new Error((j && j.error) || 'No se pudo subir el archivo.');
    state.archivos.push({ id: eq.id, inv: eq.inv, serie: eq.serie, equipo: eq.equipo, servicio: eq.servicio, categoria: categoria || '', descripcion: descripcion || '', nombre: j.name || file.name, enlace: j.url, fecha: j.fecha || '' });
    saveArchivos();
    return j;
  }
  // Adjunta un archivo a un evento ya registrado (preventivo/correctivo/pendiente):
  // lo sube a Drive con categoría = tipo de evento y guarda el enlace en el registro.
  async function attachToEvent(kind, record, file, categoria, descripcion) {
    if (!file) return;
    if (!sheetsReady()) { setStatus('⚠️ Para adjuntar archivos conéctate primero en <b>☁️ Google Sheets</b>. El evento se guardó igual.', 'warn'); return; }
    setStatus('📎 Subiendo adjunto a Drive…', 'info');
    try {
      const j = await uploadArchivoFile(record, file, categoria, descripcion);
      record.archivoUrl = j.url; record.archivoNombre = j.name || file.name;
      (kind === 'reg' ? saveRegistros : kind === 'corr' ? saveCorrectivos : savePendientes)();
      renderRegistry(); renderCorrectivos(); renderPendientes();
      if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
      autoExportSheets();
      setStatus('✅ Adjunto subido: <b>' + esc(j.name || file.name) + '</b> (' + esc(categoria) + ').', 'ok');
    } catch (e) {
      setStatus('⚠️ El evento se guardó, pero el adjunto no se pudo subir: ' + esc(e.message || e), 'warn');
    }
  }

  async function exportToSheets(silent) {
    if (!sheetsReady()) { if (!silent) showInline($('#sheetsStatus'), 'err', 'Primero pega y guarda la URL de la app web.'); return; }
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
  function autoExportSheets() { if (sheetsAuto() && sheetsReady() && state.equipos.length) exportToSheets(true); }

  async function testSheets() {
    if (!sheetsReady()) { showInline($('#sheetsStatus'), 'err', 'Primero pega y guarda la URL de la app web.'); return; }
    const btn = $('#sheetsTest'); const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Probando…';
    try {
      let j;
      if (GAS) { j = await gasCall('appPull'); }
      else { const res = await fetch(apiUrl(), { method: 'GET', redirect: 'follow' }); j = JSON.parse(await res.text()); }
      showInline($('#sheetsStatus'), 'ok', '✓ Conexión correcta. Filas en "Eventos": ' + (j && j.count != null ? j.count : '—') + '.');
    } catch (e) {
      showInline($('#sheetsStatus'), 'warn', '⚠️ No se pudo confirmar la conexión (puede ser CORS). Aun así el envío suele funcionar; pulsa Enviar y revisa la planilla.');
    } finally { btn.textContent = orig; btn.disabled = false; }
  }

  // URL de lectura de la app web (añade ?api). extra='equipos' para el inventario.
  function apiUrl(extra) {
    const u = sheetsUrl(); const q = extra ? 'api=' + extra : 'api=1';
    return u + (u.indexOf('?') >= 0 ? '&' : '?') + q;
  }
  // Trae el inventario por separado (puede ser grande). Devuelve la grilla o null.
  async function pullEquiposGrid() {
    try {
      if (GAS) { const je = await gasCall('appPullEquipos'); return je && je.equipos ? je.equipos : null; }
      const res = await fetch(apiUrl('equipos'), { method: 'GET', redirect: 'follow' });
      const je = JSON.parse(await res.text()); return je && je.equipos ? je.equipos : null;
    } catch (e) { return null; }
  }
  // Lee de Google Sheets el snapshot de datos y lo carga en la app
  async function pullFromSheets(silent) {
    if (!sheetsReady()) { if (!silent) showInline($('#sheetsStatus'), 'err', 'Primero pega y guarda la URL de la app web.'); return; }
    const btn = $('#sheetsPull'); const orig = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = 'Trayendo…'; }
    try {
      let j;
      if (GAS) { j = await gasCall('appPull'); }
      else { const res = await fetch(apiUrl(), { method: 'GET', redirect: 'follow' }); j = JSON.parse(await res.text()); }
      if (!j) { if (!silent) showInline($('#sheetsStatus'), 'warn', '⚠️ La planilla no devolvió datos. Reimplementa la última versión del script (Implementar → Nueva versión) e inténtalo otra vez.'); return; }
      // Archivos: siempre desde la hoja "Archivos" (la gestiona Apps Script)
      if (j.tablas && j.tablas.Archivos) { state.archivos = archivosFromTabla(j.tablas.Archivos); saveArchivos(); }
      // Inventario de equipos: si no hay uno cargado, traerlo (por separado; best-effort)
      let equiposCargados = 0;
      if (!state.equipos.length) {
        const grid = j.equipos || await pullEquiposGrid();
        if (grid) { const eqs = equiposFromRows(grid); if (eqs.length) { applyEquipos(eqs); equiposCargados = eqs.length; } }
      }
      const eqMsg = equiposCargados ? equiposCargados + ' equipos, ' : '';
      let obj = null;
      if (j.data) { try { obj = (typeof j.data === 'string') ? JSON.parse(j.data) : j.data; } catch (_) { obj = null; } }
      if (obj && restoreState(obj)) {
        saveRegistros(); saveCorrectivos(); savePendientes();
        renderRegistry(); renderCorrectivos(); renderPendientes();
        if (state.events.length) { renderPreview(); renderDiscrepancias(); }
        if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
        revealRegistrarIfData(); touch();
        showInline($('#sheetsStatus'), 'ok', '✓ Datos traídos de Google Sheets: ' + eqMsg + state.registros.length +
          ' mant., ' + state.correctivos.length + ' corr., ' + state.pendientes.length + ' pend.' +
          (state.archivos.length ? ', ' + state.archivos.length + ' archivo(s).' : ''));
      } else if (j.tablas) {
        // Sin snapshot _datos: reconstruir correctivos y pendientes desde las hojas visibles
        const rec = reconstructFromTablas(j.tablas);
        if (rec.correctivos.length || rec.pendientes.length || state.archivos.length) {
          state.correctivos = rec.correctivos; state.pendientes = rec.pendientes;
          saveCorrectivos(); savePendientes();
          renderCorrectivos(); renderPendientes();
          if (state.events.length) { renderPreview(); renderDiscrepancias(); }
          if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
          revealRegistrarIfData(); touch();
          showInline($('#sheetsStatus'), 'ok', '✓ Traídos desde las hojas: ' + eqMsg + rec.correctivos.length +
            ' correctivo(s), ' + rec.pendientes.length + ' pendiente(s)' +
            (state.archivos.length ? ' y ' + state.archivos.length + ' archivo(s)' : '') +
            '. (Las mantenciones preventivas requieren el snapshot _datos: vuelve a Enviar con esta versión.)');
        } else if (equiposCargados) {
          showInline($('#sheetsStatus'), 'ok', '✓ Inventario traído de Google Sheets: ' + equiposCargados + ' equipos.');
        } else if (!silent) {
          showInline($('#sheetsStatus'), 'info', 'La planilla no tiene datos de la app para traer.');
        }
      } else if (equiposCargados) {
        showInline($('#sheetsStatus'), 'ok', '✓ Inventario traído de Google Sheets: ' + equiposCargados + ' equipos.');
      } else if (!silent) {
        showInline($('#sheetsStatus'), 'info', 'La planilla aún no tiene datos guardados por la app.');
      }
    } catch (e) {
      if (!silent) showInline($('#sheetsStatus'), 'warn', '⚠️ No se pudo leer de Google Sheets (puede ser CORS): ' + (e.message || e));
    } finally { if (btn) { btn.textContent = orig; btn.disabled = false; } }
  }
  // Muestra la tarjeta de registro si hay datos (registros/correctivos) aunque no se haya cargado el archivo
  function revealRegistrarIfData() {
    if (state.registros.length || state.correctivos.length) {
      $('#registrar').style.display = 'block';
      $('#registrarHint').style.display = 'none';
    }
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
      renderPendientes();
      renderDiscrepancias();
      $('#download').disabled = false;
      $('#registrar').style.display = 'block';   // habilitar registro de mantenciones
      $('#registrarHint').style.display = 'none';
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

  function estadoStyle(st) {
    if (st.startsWith('Realizada')) return 'background:#C6EFCE;color:#136b30';
    if (st === 'Reprogramada') return 'background:#FCE4A6;color:#8a6d00';
    if (st === 'Puesta en Marcha') return 'background:#DDEBF7;color:#0a4a6e';
    if (st === 'Fuera de Servicio' || st === 'No Realizada') return 'background:#FFC7CE;color:#b3261e';
    if (st === 'No Ubicable') return 'background:#E4DFEC;color:#5b4b8a';
    if (st === 'Baja') return 'background:#D9D9D9;color:#444';
    if (st.startsWith('Pendiente')) return 'background:#FFF2CC;color:#7a5b00';
    return 'background:#eef3f5;color:#333';
  }
  function renderPreview() {
    const ev = consolidatedEvents();
    let real = 0, reprog = 0, pend = 0;
    for (const e of ev) {
      if (e.estado.startsWith('Realizada')) real++;
      else if (e.estado === 'Reprogramada') reprog++;
      else if (e.estado.startsWith('Pendiente')) pend++;
    }
    const cards = [
      ['Equipos', state.equipos.length], ['Eventos', ev.length], ['Realizadas', real],
      ['Reprogramadas', reprog], ['Pendientes', pend], ['Correctivos', state.correctivos.length]
    ].map(s => '<div class="stat"><div class="n">' + s[1].toLocaleString('es-CL') + '</div><div class="l">' + esc(s[0]) + '</div></div>').join('');
    const cols = ['ID', 'Equipo', 'Servicio', 'Mes', 'P', 'Tipo de Programación', 'R', 'Estado'];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    const rows = ev.slice(0, 20).map(e =>
      '<tr><td>' + esc(e.id) + '</td><td>' + esc(e.equipo) + '</td><td>' + esc(e.servicio) + '</td><td>' + esc(e.mes) +
      '</td><td>' + esc(e.programa) + '</td><td>' + esc(e.tipoPrograma) + '</td><td>' + esc(e.resultado) +
      '</td><td><span class="pill" style="' + estadoStyle(e.estado) + '">' + esc(e.estado) + '</span></td></tr>').join('');
    $('#result').innerHTML =
      '<div class="stats">' + cards + '</div>' +
      '<div class="muted" style="margin:4px 0 6px">Vista previa (primeros 20 de ' +
        ev.length.toLocaleString('es-CL') + ' eventos):</div>' +
      '<div class="tablewrap"><table class="prev">' + head + rows + '</table></div>';
    $('#result').style.display = 'block';
    touch();
  }

  async function downloadWorkbook(events, fname, btn) {
    const original = btn.textContent;
    btn.disabled = true; btn.textContent = 'Generando Excel…';
    try {
      const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, events, { equipos: state.equipos.length, correctivos: state.correctivos, pendientes: decoratedPendientes(), archivos: state.archivos });
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
    try { $('#equipoDetalle').scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } catch (e) {}
  }

  // Estilo de la pastilla de Estado Final del equipo
  function efStyle(v) {
    if (v === 'Operativo') return 'background:#C6EFCE;color:#136b30';
    if (v === 'No operativo') return 'background:#FFC7CE;color:#b3261e';
    if (v === 'En servicio técnico') return 'background:#FFE0B2;color:#8a5a00';
    return 'background:#eef3f5;color:#333';
  }
  // Estado final del equipo según el ÚLTIMO evento registrado (preventivo o
  // correctivo) por fecha. Empate de fecha: gana el registrado más tarde.
  function estadoFinalActual(eqId) {
    let best = null;
    const consider = (fecha, ef, tipo) => {
      if (!(fecha instanceof Date) || !ef) return;
      if (!best || fecha.getTime() >= best.fecha.getTime()) best = { fecha, ef, tipo };
    };
    state.registros.forEach(r => { if (r.id === eqId) consider(r.fechaEjecucion, r.estadoFinal, 'mantención preventiva (' + r.mes + ')'); });
    state.correctivos.forEach(c => { if (c.id === eqId) consider(c.fecha, c.estadoFinal, 'evento correctivo: ' + c.tipoEvento); });
    return best;
  }

  // ===== Estado actual del equipo · panel lateral · inventario · resumen ===
  function equipoMap() { const m = new Map(); state.equipos.forEach(eq => m.set(eq.id, eq)); return m; }
  function eventsByEquipo() {
    const m = new Map();
    consolidatedEvents().forEach(e => { let a = m.get(e.id); if (!a) { a = []; m.set(e.id, a); } a.push(e); });
    return m;
  }
  // Categoría de estado del equipo según su último evento (con estado final),
  // o "Baja" si algún evento lo marca como dado de baja.
  function categoriaFrom(eq, evs) {
    const baja = (evs || []).some(e => e.estado === 'Baja' || /^\s*baja\s*$/i.test(String(e.resultado || ''))) || /baja/i.test(String(eq.enubaja || ''));
    if (baja) return 'Baja';
    const ef = estadoFinalActual(eq.id);
    return ef ? ef.ef : 'Sin estado';
  }
  // Fecha de la última actividad registrada del equipo (mant./correctivo/pendiente)
  function ultimaActualizacion(eqId) {
    let best = null;
    const consider = d => { if (d instanceof Date && (!best || d.getTime() > best.getTime())) best = d; };
    state.registros.forEach(r => { if (r.id === eqId) consider(r.fechaEjecucion); });
    state.correctivos.forEach(c => { if (c.id === eqId) consider(c.fecha); });
    state.pendientes.forEach(p => { if (p.id === eqId) { consider(p.fechaCompromiso); (p.actualizaciones || []).forEach(a => consider(a.fecha)); } });
    return best;
  }
  function pendAbiertosDe(eqId) { return state.pendientes.filter(p => p.id === eqId && p.estado !== 'Resuelto').length; }
  function estCatClass(cat) {
    return cat === 'Operativo' ? 'est-op' : cat === 'No operativo' ? 'est-no'
      : cat === 'En servicio técnico' ? 'est-st' : cat === 'Baja' ? 'est-baja' : 'est-na';
  }

  // ---- Navegación entre vistas (panel lateral) ---------------------------
  function panelActive(id) { const p = $('#tab-' + id); return !!(p && p.classList.contains('active')); }
  function setActivePanel(panelId) { document.querySelectorAll('.tabpanel').forEach(p => p.classList.toggle('active', p.id === panelId)); }
  function goTab(id) {
    document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === id));
    document.querySelectorAll('.invcat').forEach(b => b.classList.remove('active'));
    setActivePanel('tab-' + id);
    if (id === 'resumen') renderResumen();
    if (id === 'inventario') renderInventario();
    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch (e) { }
  }
  function goInventario(cat) {
    state.invFilter = cat || 'todos';
    document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.invcat').forEach(b => b.classList.toggle('active', b.dataset.cat === state.invFilter));
    setActivePanel('tab-inventario');
    renderInventario();
    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch (e) { }
  }
  function verDetalle(eq) { if (!eq) return; goTab('registrar'); selectEquipo(eq); }

  // ---- Contadores del panel lateral --------------------------------------
  function setCount(sel, v) { const el = $(sel); if (el) el.textContent = (v === '' || v == null) ? '' : String(v); }
  function renderSidebarCounts() {
    const nP = state.pendientes.filter(p => p.estado !== 'Resuelto').length;
    setCount('#navPendCount', nP || '');
    if (!state.equipos.length) { ['#catTodos', '#catOperativo', '#catNoOp', '#catST', '#catBaja', '#catPend'].forEach(s => setCount(s, '')); return; }
    const map = eventsByEquipo();
    const pendIds = new Set(state.pendientes.filter(p => p.estado !== 'Resuelto').map(p => p.id));
    let op = 0, no = 0, st = 0, baja = 0, cp = 0;
    state.equipos.forEach(eq => {
      const cat = categoriaFrom(eq, map.get(eq.id));
      if (cat === 'Operativo') op++; else if (cat === 'No operativo') no++;
      else if (cat === 'En servicio técnico') st++; else if (cat === 'Baja') baja++;
      if (pendIds.has(eq.id)) cp++;
    });
    setCount('#catTodos', state.equipos.length); setCount('#catOperativo', op); setCount('#catNoOp', no);
    setCount('#catST', st); setCount('#catBaja', baja); setCount('#catPend', cp);
  }
  // Re-render de las vistas derivadas tras cualquier cambio de datos
  function touch() {
    renderSidebarCounts();
    if (panelActive('inventario')) renderInventario();
    if (panelActive('resumen')) renderResumen();
  }

  // ---- Vista Inventario ---------------------------------------------------
  const INV_CAT_LABEL = { todos: '📋 Inventario — Todos', 'Operativo': '🟢 Operativos', 'No operativo': '🔴 No operativos', 'En servicio técnico': '🛠️ En servicio técnico', 'Baja': '⚫ Baja', pendientes: '📎 Con pendientes' };
  function renderInventario() {
    const hint = $('#inventarioHint'), card = $('#inventarioCard');
    if (!state.equipos.length) { if (hint) hint.style.display = 'block'; if (card) card.style.display = 'none'; return; }
    if (hint) hint.style.display = 'none'; if (card) card.style.display = 'block';
    const cat = state.invFilter || 'todos';
    $('#inventarioTitle').textContent = INV_CAT_LABEL[cat] || ('📋 Inventario — ' + cat);
    const map = eventsByEquipo();
    const pendIds = new Set(state.pendientes.filter(p => p.estado !== 'Resuelto').map(p => p.id));
    const term = ($('#inventarioFilter') ? $('#inventarioFilter').value : '').trim().toLowerCase();
    const cols = ['Familia', 'ID', 'N° Carpeta', 'N° Inventario', 'Equipo', 'Servicio', 'Unidad', 'Ubicación', 'Procedencia', 'Marca', 'Modelo', 'N° Serie', 'Año', 'VUR', 'Clasificación', 'ENU/Baja', 'Frecuencia', 'Observación', 'Estado actual', 'Última actualización', 'Pendientes'];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let shown = 0;
    const rows = state.equipos.map((eq, i) => {
      const c = categoriaFrom(eq, map.get(eq.id));
      if (cat === 'pendientes') { if (!pendIds.has(eq.id)) return ''; }
      else if (cat !== 'todos') { if (c !== cat) return ''; }
      if (term) {
        const hay = [eq.familia, eq.id, eq.carpeta, eq.inv, eq.equipo, eq.servicio, eq.unidad, eq.ubicacion, eq.marca, eq.modelo, eq.serie, c].join(' ').toLowerCase();
        if (hay.indexOf(term) === -1) return '';
      }
      shown++;
      const ua = ultimaActualizacion(eq.id), np = pendAbiertosDe(eq.id);
      const v = x => (x != null && String(x).trim() !== '') ? esc(x) : '—';
      return '<tr class="clik" data-eq="' + i + '">' +
        '<td>' + v(eq.familia) + '</td><td>' + v(eq.id) + '</td><td>' + v(eq.carpeta) + '</td><td>' + v(eq.inv) + '</td>' +
        '<td>' + v(eq.equipo) + '</td><td>' + v(eq.servicio) + '</td><td>' + v(eq.unidad) + '</td><td>' + v(eq.ubicacion) + '</td>' +
        '<td>' + v(eq.procedencia) + '</td><td>' + v(eq.marca) + '</td><td>' + v(eq.modelo) + '</td><td>' + v(eq.serie) + '</td>' +
        '<td>' + v(eq.anio) + '</td><td>' + v(eq.vur) + '</td><td>' + v(eq.clasif) + '</td><td>' + v(eq.enubaja) + '</td>' +
        '<td>' + v(eq.frecuencia) + '</td><td>' + v(eq.observacion) + '</td>' +
        '<td><span class="estpill ' + estCatClass(c) + '">' + esc(c) + '</span></td>' +
        '<td>' + (ua ? esc(fmtDate(ua)) : '—') + '</td>' +
        '<td>' + (np ? '<b style="color:#7a5b00">Sí (' + np + ')</b>' : '—') + '</td>' +
        '</tr>';
    }).join('');
    $('#inventarioCount').textContent = shown + (shown !== state.equipos.length ? ' / ' + state.equipos.length : '');
    const t = $('#inventarioTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="' + cols.length + '" class="nomatch">Sin equipos en esta vista.</td></tr>');
    t.querySelectorAll('tr.clik').forEach(tr => tr.addEventListener('click', () => verDetalle(state.equipos[+tr.dataset.eq])));
  }

  // ---- Vista Resumen mensual (con desglose clicable) ---------------------
  function monthlyAgg() {
    const ev = consolidatedEvents();
    const m = MP.MONTHS_FULL.map(() => ({ prog: [], real: [], pend: [] }));
    ev.forEach(e => {
      const b = m[e.nMes - 1]; if (!b) return;
      if (e.programa) b.prog.push(e);
      if (e.estado && e.estado.indexOf('Realizada') === 0) b.real.push(e);
      else if (e.estado && e.estado.indexOf('Pendiente') === 0) b.pend.push(e);
    });
    return m;
  }
  function renderResumen() {
    const hint = $('#resumenHint'), card = $('#resumenCard');
    if (!state.events.length) { if (hint) hint.style.display = 'block'; if (card) card.style.display = 'none'; return; }
    if (hint) hint.style.display = 'none'; if (card) card.style.display = 'block';
    const m = monthlyAgg();
    const cell = (mi, metric, arr) => arr.length
      ? '<button class="rcell' + (state.resSel && state.resSel.mi === mi && state.resSel.metric === metric ? ' active' : '') + '" data-mi="' + mi + '" data-metric="' + metric + '">' + arr.length + '</button>'
      : '<span class="rcell zero">0</span>';
    const tot = { prog: 0, real: 0, pend: 0 };
    const body = m.map((x, mi) => {
      tot.prog += x.prog.length; tot.real += x.real.length; tot.pend += x.pend.length;
      return '<tr><td class="mescell">' + esc(MP.MONTHS_FULL[mi]) + '</td>' +
        '<td>' + cell(mi, 'prog', x.prog) + '</td><td>' + cell(mi, 'real', x.real) + '</td><td>' + cell(mi, 'pend', x.pend) + '</td></tr>';
    }).join('');
    const head = '<tr><th>Mes</th><th>Programadas</th><th>Realizadas</th><th>Pendientes</th></tr>';
    const totRow = '<tr><td class="mescell">Total</td><td><b>' + tot.prog + '</b></td><td><b>' + tot.real + '</b></td><td><b>' + tot.pend + '</b></td></tr>';
    $('#resumenTable').innerHTML = head + body + totRow;
    $('#resumenTable').querySelectorAll('.rcell[data-mi]').forEach(b =>
      b.addEventListener('click', () => { state.resSel = { mi: +b.dataset.mi, metric: b.dataset.metric }; renderResumen(); }));
    renderResumenDrill();
  }
  function renderResumenDrill() {
    const box = $('#resumenDrill'); if (!box) return;
    const sel = state.resSel;
    if (!sel) { box.innerHTML = '<div class="muted">Haz clic en un número de la tabla para ver la lista de equipos de ese grupo.</div>'; return; }
    const m = monthlyAgg()[sel.mi];
    const arr = (m ? m[sel.metric] : []) || [];
    const label = sel.metric === 'prog' ? 'Programadas' : sel.metric === 'real' ? 'Realizadas' : 'Pendientes';
    const emap = equipoMap();
    const head = '<tr><th>ID</th><th>Equipo</th><th>Servicio</th><th>Programa</th><th>Resultado</th><th>Fecha</th><th>Estado</th></tr>';
    const rows = arr.map((e, i) => '<tr class="clik" data-ei="' + i + '"><td>' + esc(e.id) + '</td><td>' + esc(e.equipo) +
      '</td><td>' + esc(e.servicio) + '</td><td>' + esc(e.programa) + '</td><td>' + esc(e.resultado) +
      '</td><td>' + esc(fmtDate(e.fechaEjecucion)) + '</td><td><span class="pill" style="' + estadoStyle(e.estado) + '">' + esc(e.estado) + '</span></td></tr>').join('');
    box.innerHTML = '<div class="drillhead">' + esc(MP.MONTHS_FULL[sel.mi]) + ' · ' + label + ' <span class="badge">' + arr.length + '</span>' +
      '<button class="btn btn-ghost btn-sm" id="resClear" type="button" style="margin-left:auto">✕ Cerrar</button></div>' +
      (arr.length ? '<div class="tablewrap"><table class="prev">' + head + rows + '</table></div>' : '<div class="muted">Sin equipos.</div>');
    const rc = $('#resClear'); if (rc) rc.addEventListener('click', () => { state.resSel = null; renderResumen(); });
    box.querySelectorAll('tr.clik').forEach(tr => tr.addEventListener('click', () => {
      const e = arr[+tr.dataset.ei]; const eq = emap.get(e.id) || state.equipos.find(q => q.id === e.id); if (eq) verDetalle(eq);
    }));
  }

  function renderDetalle(eq) {
    const hist = consolidatedEvents().filter(e => e.id === eq.id).sort((a, b) => a.nMes - b.nMes);
    const pm = MP.programmedMonths(eq);
    const ef = estadoFinalActual(eq.id);
    const efBanner = ef
      ? '<div class="efbanner" style="' + efStyle(ef.ef) + '">Estado final del equipo: <b>' + esc(ef.ef) +
        '</b> <span class="efmeta">— según el último evento (' + esc(ef.tipo) + ' del ' + esc(fmtDate(ef.fecha)) + ')</span></div>'
      : '<div class="efbanner efnone">Estado final del equipo: <b>—</b> <span class="efmeta">— aún no hay eventos registrados con estado para este equipo</span></div>';
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
    // Pendientes del equipo
    const pend = state.pendientes.filter(p => p.id === eq.id);
    let pendBody = '';
    if (pend.length) {
      const pc = ['Tipo', 'Fecha compromiso', 'Resp. administrativo', 'Resp. ejecución', 'Observación'];
      const ph = '<tr>' + pc.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
      const pr = pend.map(p =>
        '<tr><td>' + esc(p.tipo || '—') + '</td><td>' + esc(fmtDate(p.fechaCompromiso)) + '</td><td>' + esc(p.respAdministrativo) + '</td><td>' +
        esc(p.respEjecucion) + '</td><td>' + esc(p.observacion) + '</td></tr>'
      ).join('');
      pendBody = '<div class="muted" style="margin:10px 0 4px">Pendientes registrados</div>' +
        '<div class="tablewrap"><table class="prev">' + ph + pr + '</table></div>';
    }
    $('#equipoDetalle').innerHTML = efBanner + head + body + corrBody + pendBody +
      '<div class="actions" style="margin-top:10px">' +
        '<button id="detRegistrar" class="btn btn-accent btn-sm" type="button">🔧 Registrar mantención preventiva</button>' +
        '<button id="detCorrectivo" class="btn btn-primary btn-sm" type="button">🛠️ Registrar evento correctivo</button>' +
        '<button id="detPendiente" class="btn btn-ghost btn-sm" type="button">📌 Registrar pendiente</button>' +
      '</div>';
    $('#equipoDetalle').style.display = 'block';
    $('#detRegistrar').addEventListener('click', () => openModal(eq));
    $('#detCorrectivo').addEventListener('click', () => openCModal(eq));
    $('#detPendiente').addEventListener('click', () => openPModal(eq));
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
    fillSelect($('#pTipo'), MP.PENDIENTE_TIPOS);
    fillSelect($('#pRespAdmin'), MP.EJECUTORES);
    fillSelect($('#pRespEjec'), MP.EJECUTORES);
    fillSelect($('#gEstado'), MP.ESTADO_PENDIENTE_OPCIONES);
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
    $('#mFile').value = ''; $('#mFileDesc').value = ''; $('#mAttachField').style.display = sheetsReady() ? 'block' : 'none';
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
    const attFile = $('#mFile') ? $('#mFile').files[0] : null;   // capturar antes de cerrar
    const attDesc = $('#mFileDesc') ? $('#mFileDesc').value.trim() : '';
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
    if (attFile) attachToEvent('reg', reg, attFile, 'Mantención preventiva', attDesc || (reg.mes + ' · ' + reg.resultado + ' · ' + fmtDate(fecha)));
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
  // Celda con enlace al archivo adjunto del evento (📎) o "—"
  function adjCell(o) {
    return o && o.archivoUrl
      ? '<a href="' + esc(o.archivoUrl) + '" target="_blank" rel="noopener" title="' + esc(o.archivoNombre || 'Ver adjunto') + '">📎</a>'
      : '<span class="muted">—</span>';
  }
  function renderRegistry() {
    const wrap = $('#registryWrap');
    if (!state.registros.length) { wrap.style.display = 'none'; return; }
    wrap.style.display = 'block';
    const term = ($('#registryFilter') ? $('#registryFilter').value : '').trim().toLowerCase();
    const cols = ['#', 'ID', 'Equipo', 'Serie / Inv', 'Mes', 'Fecha', 'Programa', 'Resultado', 'Ejecutor', 'Estado Final', 'Adjunto', ''];
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
        '<td>' + adjCell(e) + '</td>' +
        '<td><button class="btn btn-del" title="Eliminar" data-del="' + i + '">✕</button></td>' +
      '</tr>';
    }).join('');
    $('#registryCount').textContent = term ? (shown + ' / ' + state.registros.length) : state.registros.length;
    const t = $('#registryTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="12" class="nomatch">Sin coincidencias para el filtro.</td></tr>');
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
    if (!state.correctivos.length) { wrap.style.display = 'none'; touch(); return; }
    wrap.style.display = 'block';
    const term = ($('#correctivosFilter') ? $('#correctivosFilter').value : '').trim().toLowerCase();
    const cols = ['#', 'ID', 'Equipo', 'Tipo de Evento', 'Fecha', 'Folio', 'N° Envío', 'Empresa', 'Ejecutor', 'Estado Final', 'Adjunto', ''];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let shown = 0;
    const rows = state.correctivos.map((c, i) => {
      const hay = [c.id, c.equipo, c.serie, c.inv, c.tipoEvento, fmtDate(c.fecha), c.folioSolicitud, c.folioGuia, c.nEnvio, c.empresa, c.ejecutor, c.descripcion, c.estadoFinal].join(' ').toLowerCase();
      if (term && hay.indexOf(term) === -1) return '';
      shown++;
      return '<tr><td>' + (i + 1) + '</td><td>' + esc(c.id) + '</td><td>' + esc(c.equipo) + '</td><td>' + esc(c.tipoEvento) +
        '</td><td>' + esc(fmtDate(c.fecha)) + '</td><td>' + esc(c.folioSolicitud || c.folioGuia) + '</td><td>' + esc(c.nEnvio) +
        '</td><td>' + esc(c.empresa) + '</td><td>' + esc(c.ejecutor) + '</td><td>' + esc(c.estadoFinal) +
        '</td><td>' + adjCell(c) + '</td><td><button class="btn btn-del" title="Eliminar" data-del="' + i + '">✕</button></td></tr>';
    }).join('');
    $('#correctivosCount').textContent = term ? (shown + ' / ' + state.correctivos.length) : state.correctivos.length;
    const t = $('#correctivosTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="12" class="nomatch">Sin coincidencias para el filtro.</td></tr>');
    t.querySelectorAll('button[data-del]').forEach(b =>
      b.addEventListener('click', () => {
        state.correctivos.splice(+b.dataset.del, 1);
        saveCorrectivos(); renderCorrectivos();
        if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
      }));
    touch();
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
    $('#cFile').value = ''; $('#cFileDesc').value = ''; $('#cAttachField').style.display = sheetsReady() ? 'block' : 'none';
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
    const corr = MP.buildCorrectivo(eq, data);
    const attFile = $('#cFile') ? $('#cFile').files[0] : null;
    const attDesc = $('#cFileDesc') ? $('#cFileDesc').value.trim() : '';
    state.correctivos.push(corr);
    saveCorrectivos();
    closeCModal();
    renderCorrectivos();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
    setStatus('✅ Evento correctivo registrado: <b>' + esc(eq.equipo) + '</b> — ' + esc(tipo) + '.', 'ok');
    autoExportSheets();
    if (attFile) {
      const auto = tipo + (corr.fecha ? ' · ' + fmtDate(corr.fecha) : '') +
        (corr.folioSolicitud ? ' · Folio ' + corr.folioSolicitud : corr.folioGuia ? ' · Guía ' + corr.folioGuia : '') +
        (corr.empresa ? ' · ' + corr.empresa : '');
      attachToEvent('corr', corr, attFile, 'Correctivo: ' + tipo, attDesc || auto);
    }
  }

  // ---- Pendientes: tabla --------------------------------------------------
  // Clasificación Eisenhower de un pendiente
  function clasifPendiente(p) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let urgente = false;
    if (p.fechaCompromiso instanceof Date) {
      const d = Math.round((p.fechaCompromiso.getTime() - today.getTime()) / 86400000);
      urgente = d <= 7;   // vencido o dentro de 7 días
    }
    const ef = estadoFinalActual(p.id);
    const importante = !!(ef && (ef.ef === 'No operativo' || ef.ef === 'En servicio técnico'));
    const cuadrante = urgente && importante ? 'Hacer ya' : (!urgente && importante ? 'Planificar' : (urgente ? 'Delegar' : 'Posponer'));
    return { urgente, importante, cuadrante };
  }
  // Pendientes con su clasificación, para exportar
  function tareasResumen(p) { const t = p.tareas || []; return t.length ? (t.filter(x => x.hecha).length + '/' + t.length) : ''; }
  function ultimaActTexto(p) {
    const a = p.actualizaciones || []; if (!a.length) return '';
    const last = a[a.length - 1];
    return (last.fecha ? fmtDate(last.fecha) + ': ' : '') + (last.texto || '');
  }
  function decoratedPendientes() {
    return state.pendientes.map(p => {
      const c = clasifPendiente(p);
      return Object.assign({}, p, {
        estado: p.estado || 'Abierto',
        tareasResumen: tareasResumen(p),
        ultimaAct: ultimaActTexto(p),
        cuadrante: c.cuadrante,
        urgencia: c.urgente ? 'Urgente' : 'No urgente',
        importancia: c.importante ? 'Importante' : 'No importante'
      });
    });
  }
  function estadoPendStyle(s) {
    if (s === 'Resuelto') return 'background:#C6EFCE;color:#136b30';
    if (s === 'En progreso') return 'background:#DDEBF7;color:#0a4a6e';
    return 'background:#FFF2CC;color:#7a5b00'; // Abierto
  }
  function dueInfo(f) {
    if (!(f instanceof Date)) return '';
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const d = Math.round((f.getTime() - today.getTime()) / 86400000);
    if (d < 0) return '<span class="venc">vencido hace ' + (-d) + ' día(s)</span>';
    if (d === 0) return '<span class="venc">vence hoy</span>';
    return 'en ' + d + ' día(s)';
  }
  // Prioridad para Ivy Lee / Sapo: importante(2) + urgente(1), luego fecha asc.
  function pendPrioritized() {
    const list = state.pendientes.map((p, i) => ({ p, i })).filter(x => x.p.estado !== 'Resuelto');
    const score = x => { const c = clasifPendiente(x.p); return (c.importante ? 2 : 0) + (c.urgente ? 1 : 0); };
    const ft = f => (f instanceof Date) ? f.getTime() : Infinity;
    list.sort((a, b) => score(b) - score(a) || ft(a.p.fechaCompromiso) - ft(b.p.fechaCompromiso));
    return list;
  }
  function pItemHTML(p, i, num) {
    return '<div class="pitem" data-pi="' + i + '" title="Clic para gestionar"><div class="pnum">' + num + '</div><div class="pbody">' +
      '<div class="pe">' + esc(p.equipo) + ' <span style="color:var(--muted);font-weight:600">(ID ' + esc(p.id) + ')</span></div>' +
      '<div class="pf">📅 ' + esc(fmtDate(p.fechaCompromiso)) + ' · ' + dueInfo(p.fechaCompromiso) + '</div>' +
      '<div class="pm">Ejec.: ' + esc(p.respEjecucion || '—') + ' · ' + esc(p.observacion || '') + '</div>' +
      '<div style="margin-top:3px"><span class="pill" style="' + estadoPendStyle(p.estado) + '">' + esc(p.estado || 'Abierto') + '</span></div>' +
      '</div></div>';
  }
  function wireOpenG(cont) {
    cont.querySelectorAll('[data-pi]').forEach(el =>
      el.addEventListener('click', () => openGModal(state.pendientes[+el.dataset.pi])));
  }
  function renderIvy() {
    const list = pendPrioritized().slice(0, 6);
    const cont = $('#ivy');
    cont.innerHTML = list.length ? list.map(({ p, i }, n) => pItemHTML(p, i, n + 1)).join('')
      : '<div class="muted">No hay pendientes activos. 🎉</div>';
    wireOpenG(cont);
  }
  function renderSapo() {
    const list = pendPrioritized();
    const cont = $('#sapo');
    if (!list.length) { cont.innerHTML = '<div class="muted">No hay pendientes activos. 🎉</div>'; return; }
    const sapo = list[0], rest = list.slice(1);
    cont.innerHTML =
      '<div class="sapocard" data-pi="' + sapo.i + '">🐸 <b>El sapo de hoy</b> — hazlo primero<div style="margin-top:6px">' +
        '<b style="color:var(--teal-d)">' + esc(sapo.p.equipo) + '</b> (ID ' + esc(sapo.p.id) + ') · 📅 ' + esc(fmtDate(sapo.p.fechaCompromiso)) + ' · ' + dueInfo(sapo.p.fechaCompromiso) +
        '<div class="muted" style="margin-top:2px">Ejec.: ' + esc(sapo.p.respEjecucion || '—') + (sapo.p.observacion ? ' · ' + esc(sapo.p.observacion) : '') + '</div></div></div>' +
      (rest.length ? '<div class="muted" style="margin:10px 0 4px">Luego, en orden:</div>' + rest.map(({ p, i }, n) => pItemHTML(p, i, n + 2)).join('') : '');
    wireOpenG(cont);
  }
  function renderPendManage() {
    const has = state.pendientes.length > 0;
    $('#pendManageCard').style.display = has ? 'block' : 'none';
    if (!has) return;
    renderEisenhower(); renderIvy(); renderSapo();
  }
  // Tablero Eisenhower (4 cuadrantes). Los pendientes "Resuelto" no aparecen.
  function renderEisenhower() {
    if (!state.pendientes.length) return;
    const quads = { 'Hacer ya': [], 'Planificar': [], 'Delegar': [], 'Posponer': [] };
    let resueltos = 0;
    state.pendientes.forEach((p, i) => {
      if (p.estado === 'Resuelto') { resueltos++; return; }
      quads[clasifPendiente(p).cuadrante].push({ p, i });
    });
    const ft = f => (f instanceof Date) ? f.getTime() : Infinity;
    Object.keys(quads).forEach(k => quads[k].sort((a, b) => ft(a.p.fechaCompromiso) - ft(b.p.fechaCompromiso)));
    const meta = {
      'Hacer ya': ['q1', '🔴 Hacer ya', 'Urgente + Importante'],
      'Planificar': ['q2', '🔵 Planificar', 'Importante, no urgente'],
      'Delegar': ['q3', '🟡 Delegar', 'Urgente, no importante'],
      'Posponer': ['q4', '⚪ Posponer', 'Ni urgente ni importante']
    };
    $('#eisenhower').innerHTML = Object.keys(meta).map(k => {
      const m = meta[k], items = quads[k];
      const body = items.length ? items.map(({ p, i }) => {
        const tr = tareasResumen(p);
        return '<div class="pcard" data-pi="' + i + '" title="Clic para gestionar">' +
          '<div class="pe">' + esc(p.equipo) + ' <span style="color:var(--muted);font-weight:600">(ID ' + esc(p.id) + ')</span></div>' +
          '<div class="pf">📅 ' + esc(fmtDate(p.fechaCompromiso)) + ' · ' + dueInfo(p.fechaCompromiso) + '</div>' +
          '<div class="pm">Ejec.: ' + esc(p.respEjecucion || '—') + ' · Adm.: ' + esc(p.respAdministrativo || '—') + '</div>' +
          (p.observacion ? '<div class="pm">' + esc(p.observacion) + '</div>' : '') +
          '<div style="margin-top:4px"><span class="pill" style="' + estadoPendStyle(p.estado) + '">' + esc(p.estado || 'Abierto') + '</span>' +
          (tr ? ' <span class="pill" style="background:#eef3f5;color:#333">✓ ' + tr + '</span>' : '') + '</div>' +
          '</div>';
      }).join('') : '<div class="empty">— sin pendientes —</div>';
      return '<div class="quad ' + m[0] + '"><h4>' + m[1] + ' <span class="qcount">' + items.length + '</span></h4>' +
        '<div class="qsub muted">' + m[2] + '</div>' + body + '</div>';
    }).join('') + (resueltos ? '<div class="muted" style="grid-column:1/-1;margin-top:4px">✔ ' + resueltos + ' pendiente(s) resuelto(s) (ocultos del tablero; visibles en la tabla).</div>' : '');
    $('#eisenhower').querySelectorAll('.pcard[data-pi]').forEach(el =>
      el.addEventListener('click', () => openGModal(state.pendientes[+el.dataset.pi])));
  }

  function renderPendientes() {
    renderPendManage();
    const has = state.pendientes.length > 0;
    $('#pendTableCard').style.display = has ? 'block' : 'none';
    $('#pendNone').style.display = has ? 'none' : 'block';
    if (!has) { touch(); return; }
    const term = ($('#pendientesFilter') ? $('#pendientesFilter').value : '').trim().toLowerCase();
    const cols = ['#', 'ID', 'Equipo', 'Tipo', 'Fecha compromiso', 'Estado', 'Tareas', 'Resp. ejecución', 'Observación', 'Adjunto', '', ''];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    let shown = 0;
    const rows = state.pendientes.map((p, i) => {
      const hay = [p.id, p.equipo, p.serie, p.inv, p.tipo, fmtDate(p.fechaCompromiso), p.estado, p.respAdministrativo, p.respEjecucion, p.observacion].join(' ').toLowerCase();
      if (term && hay.indexOf(term) === -1) return '';
      shown++;
      const tr = tareasResumen(p);
      return '<tr><td>' + (i + 1) + '</td><td>' + esc(p.id) + '</td><td>' + esc(p.equipo) + '</td><td>' + esc(p.tipo || '—') + '</td><td>' + esc(fmtDate(p.fechaCompromiso)) +
        '</td><td><span class="pill" style="' + estadoPendStyle(p.estado) + '">' + esc(p.estado || 'Abierto') + '</span></td><td>' + esc(tr || '—') +
        '</td><td>' + esc(p.respEjecucion) + '</td><td>' + esc(p.observacion) + '</td><td>' + adjCell(p) +
        '</td><td><button class="btn btn-ghost btn-sm" title="Gestionar" data-gest="' + i + '">⚙️ Gestionar</button></td>' +
        '<td><button class="btn btn-del" title="Eliminar" data-del="' + i + '">✕</button></td></tr>';
    }).join('');
    $('#pendientesCount').textContent = term ? (shown + ' / ' + state.pendientes.length) : state.pendientes.length;
    const t = $('#pendientesTable');
    t.innerHTML = head + rows + (shown ? '' : '<tr><td colspan="12" class="nomatch">Sin coincidencias para el filtro.</td></tr>');
    t.querySelectorAll('button[data-gest]').forEach(b =>
      b.addEventListener('click', () => openGModal(state.pendientes[+b.dataset.gest])));
    t.querySelectorAll('button[data-del]').forEach(b =>
      b.addEventListener('click', () => {
        state.pendientes.splice(+b.dataset.del, 1);
        savePendientes(); renderPendientes();
        if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
      }));
    touch();
  }

  // ---- Modal de pendiente -------------------------------------------------
  function updatePSave() {
    $('#pSave').disabled = !(state.pEq && $('#pTipo').value && $('#pFecha').value && $('#pRespAdmin').value && $('#pRespEjec').value && $('#pObs').value.trim());
  }
  function openPModal(eq) {
    state.pEq = eq;
    $('#pmodalEq').innerHTML = '<b>' + esc(eq.equipo) + '</b> (ID ' + esc(eq.id) + ') · Serie: <b>' +
      esc(eq.serie || '—') + '</b> · Inv: <b>' + esc(eq.inv || '—') + '</b>';
    $('#pTipo').value = ''; $('#pFecha').value = ''; $('#pRespAdmin').value = ''; $('#pRespEjec').value = ''; $('#pObs').value = '';
    $('#pFile').value = ''; $('#pFileDesc').value = ''; $('#pAttachField').style.display = sheetsReady() ? 'block' : 'none';
    showInline($('#pmodalMsg'), '', '');
    $('#pSave').disabled = true;
    $('#pmodal').classList.add('show');
  }
  // Fija la fecha de compromiso a N días hábiles (o N días corridos) desde hoy.
  function addBusinessDays(base, n) { const r = new Date(base); let added = 0; while (added < n) { r.setDate(r.getDate() + 1); const wd = r.getDay(); if (wd !== 0 && wd !== 6) added++; } return r; }
  function setPFecha(d) { $('#pFecha').value = isoDate(d); updatePSave(); }
  function closePModal() { $('#pmodal').classList.remove('show'); state.pEq = null; }
  function savePendiente() {
    const eq = state.pEq;
    if (!eq) return;
    if (!$('#pTipo').value || !$('#pFecha').value || !$('#pRespAdmin').value || !$('#pRespEjec').value || !$('#pObs').value.trim()) {
      showInline($('#pmodalMsg'), 'err', 'Complete todos los campos.'); return;
    }
    const x = $('#pFecha').value.split('-');
    const pTipo = $('#pTipo').value;
    const pend = MP.buildPendiente(eq, {
      tipo: pTipo,
      fechaCompromiso: new Date(+x[0], +x[1] - 1, +x[2]),
      respAdministrativo: $('#pRespAdmin').value,
      respEjecucion: $('#pRespEjec').value,
      observacion: $('#pObs').value.trim()
    });
    const attFile = $('#pFile') ? $('#pFile').files[0] : null;
    const attDesc = $('#pFileDesc') ? $('#pFileDesc').value.trim() : '';
    state.pendientes.push(pend);
    savePendientes();
    closePModal();
    renderPendientes();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
    setStatus('✅ Pendiente registrado: <b>' + esc(eq.equipo) + '</b>.', 'ok');
    autoExportSheets();
    if (attFile) attachToEvent('pend', pend, attFile, 'Pendiente: ' + pTipo, attDesc || (pTipo + ' · ' + fmtDate(pend.fechaCompromiso)));
  }

  // ---- Gestión de un pendiente (estado, bitácora, tareas) ----------------
  function gPersist() { savePendientes(); renderPendientes(); if (state.selDetalleEq) renderDetalle(state.selDetalleEq); autoExportSheets(); }
  function openGModal(p) {
    state.gPend = p;
    renderGModal();
    $('#gmodal').classList.add('show');
  }
  function closeGModal() { $('#gmodal').classList.remove('show'); state.gPend = null; }
  function renderGModal() {
    const p = state.gPend; if (!p) return;
    $('#gmodalEq').innerHTML = '<b>' + esc(p.equipo) + '</b> (ID ' + esc(p.id) + ') · Inv: <b>' + esc(p.inv || '—') +
      '</b>' + (p.tipo ? ' · Tipo: <b>' + esc(p.tipo) + '</b>' : '') +
      '<br>📅 Compromiso: <b>' + esc(fmtDate(p.fechaCompromiso)) + '</b> · Ejec.: ' + esc(p.respEjecucion || '—') +
      ' · Adm.: ' + esc(p.respAdministrativo || '—') + (p.observacion ? '<br>' + esc(p.observacion) : '');
    // Estado
    $('#gEstado').value = p.estado || 'Abierto';
    // Tareas
    const tareas = p.tareas || [];
    const hechas = tareas.filter(t => t.hecha).length;
    $('#gTareaResumen').textContent = tareas.length ? ('(' + hechas + '/' + tareas.length + ' hechas)') : '';
    $('#gTareaLista').innerHTML = tareas.length ? tareas.map((t, i) =>
      '<div class="titem"><label><input type="checkbox" data-tg="' + i + '"' + (t.hecha ? ' checked' : '') + '> <span' +
      (t.hecha ? ' style="text-decoration:line-through;color:var(--muted)"' : '') + '>' + esc(t.texto) + '</span></label>' +
      '<button class="btn btn-del" title="Eliminar tarea" data-td="' + i + '">✕</button></div>').join('')
      : '<div class="muted" style="font-size:12.5px">Sin tareas.</div>';
    // Bitácora (más reciente primero)
    const acts = (p.actualizaciones || []).slice().reverse();
    $('#gActLista').innerHTML = acts.length ? acts.map(a =>
      '<div class="aitem"><span class="af">' + esc(fmtDate(a.fecha)) + '</span> ' + esc(a.texto) + '</div>').join('')
      : '<div class="muted" style="font-size:12.5px">Sin actualizaciones.</div>';
    // wire
    $('#gTareaLista').querySelectorAll('input[data-tg]').forEach(c =>
      c.addEventListener('change', () => { p.tareas[+c.dataset.tg].hecha = c.checked; gPersist(); renderGModal(); }));
    $('#gTareaLista').querySelectorAll('button[data-td]').forEach(b =>
      b.addEventListener('click', () => { p.tareas.splice(+b.dataset.td, 1); gPersist(); renderGModal(); }));
  }
  function gAddTarea() {
    const p = state.gPend; const txt = $('#gTareaTexto').value.trim();
    if (!p || !txt) return;
    (p.tareas = p.tareas || []).push({ texto: txt, hecha: false });
    $('#gTareaTexto').value = '';
    gPersist(); renderGModal();
  }
  function gAddAct() {
    const p = state.gPend; const txt = $('#gActTexto').value.trim();
    if (!p || !txt) return;
    const fv = $('#gActFecha').value;
    const fecha = fv ? (function () { const x = fv.split('-'); return new Date(+x[0], +x[1] - 1, +x[2]); })() : new Date();
    (p.actualizaciones = p.actualizaciones || []).push({ fecha, texto: txt });
    $('#gActTexto').value = '';
    gPersist(); renderGModal();
  }

  // ---- Utilidades PDF: unir / separar ------------------------------------
  // Las librerías (pdf-lib, JSZip) se cargan desde CDN solo al usar la vista.
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('No se pudo cargar ' + src));
      document.head.appendChild(s);
    });
  }
  async function ensureLib(globalName, urls) {
    if (window[globalName]) return window[globalName];
    for (const u of urls) { try { await loadScript(u); if (window[globalName]) return window[globalName]; } catch (e) { } }
    throw new Error('No se pudo cargar la librería (¿hay conexión a internet?).');
  }
  const ensurePDFLib = () => ensureLib('PDFLib', ['https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js', 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js']);
  const ensureJSZip = () => ensureLib('JSZip', ['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js']);
  function fmtSize(n) { return n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(1) + ' KB' : (n / 1048576).toFixed(1) + ' MB'; }
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  // --- Unir PDF ---
  const pdfMerge = [];
  function renderMergeList() {
    const cont = $('#mergeList');
    cont.innerHTML = pdfMerge.map((f, i) =>
      '<div class="pdfitem"><span class="ord">' + (i + 1) + '</span>' +
      '<span class="nm">' + esc(f.name) + '</span><span class="sz">' + fmtSize(f.size) + '</span>' +
      '<span class="pbtns">' +
      '<button data-up="' + i + '" title="Subir"' + (i === 0 ? ' disabled' : '') + '>▲</button>' +
      '<button data-down="' + i + '" title="Bajar"' + (i === pdfMerge.length - 1 ? ' disabled' : '') + '>▼</button>' +
      '<button data-rm="' + i + '" title="Quitar">✕</button></span></div>').join('');
    $('#mergeBtn').disabled = pdfMerge.length < 2;
    cont.querySelectorAll('button[data-up]').forEach(b => b.addEventListener('click', () => { const i = +b.dataset.up; [pdfMerge[i - 1], pdfMerge[i]] = [pdfMerge[i], pdfMerge[i - 1]]; renderMergeList(); }));
    cont.querySelectorAll('button[data-down]').forEach(b => b.addEventListener('click', () => { const i = +b.dataset.down; [pdfMerge[i + 1], pdfMerge[i]] = [pdfMerge[i], pdfMerge[i + 1]]; renderMergeList(); }));
    cont.querySelectorAll('button[data-rm]').forEach(b => b.addEventListener('click', () => { pdfMerge.splice(+b.dataset.rm, 1); renderMergeList(); }));
  }
  function addMergeFiles(files) {
    for (const f of files) if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') pdfMerge.push(f);
    renderMergeList();
  }
  async function doMerge() {
    if (pdfMerge.length < 2) return;
    const btn = $('#mergeBtn'); const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Uniendo…';
    showInline($('#mergeMsg'), 'info', 'Cargando librería y uniendo PDF…');
    try {
      const PDFLib = await ensurePDFLib();
      const out = await PDFLib.PDFDocument.create();
      let totalPaginas = 0;
      for (const f of pdfMerge) {
        const buf = await readFile(f);
        const src = await PDFLib.PDFDocument.load(buf, { ignoreEncryption: true });
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p => out.addPage(p)); totalPaginas += pages.length;
      }
      const bytes = await out.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'PDF_unido_' + stamp() + '.pdf');
      showInline($('#mergeMsg'), 'ok', '✓ Unidos ' + pdfMerge.length + ' PDF (' + totalPaginas + ' páginas). Descarga iniciada.');
    } catch (e) {
      showInline($('#mergeMsg'), 'err', '❌ ' + (e.message || e));
    } finally { btn.textContent = orig; btn.disabled = pdfMerge.length < 2; }
  }

  // --- Separar PDF ---
  let pdfSplitFile = null, pdfSplitPages = 0;
  async function onSplitFile(f) {
    pdfSplitFile = f || null; pdfSplitPages = 0;
    $('#splitBtn').disabled = true;
    if (!f) { $('#splitInfo').textContent = ''; return; }
    $('#splitInfo').textContent = 'Leyendo…';
    try {
      const PDFLib = await ensurePDFLib();
      const doc = await PDFLib.PDFDocument.load(await readFile(f), { ignoreEncryption: true });
      pdfSplitPages = doc.getPageCount();
      $('#splitInfo').innerHTML = '<b>' + esc(f.name) + '</b> · ' + pdfSplitPages + ' página(s) · ' + fmtSize(f.size);
      $('#splitBtn').disabled = false;
    } catch (e) { $('#splitInfo').innerHTML = '<span style="color:var(--err)">No se pudo leer el PDF: ' + esc(e.message || e) + '</span>'; }
  }
  // "1-3, 4, 5-8" -> [{label, idx:[0-based...]}], validando contra max páginas
  function parseRanges(str, max) {
    const out = [];
    String(str || '').split(',').forEach(tok => {
      tok = tok.trim(); if (!tok) return;
      const m = tok.match(/^(\d+)\s*-\s*(\d+)$/);
      if (m) {
        const a = +m[1], b = +m[2];
        if (a < 1 || b < 1 || a > max || b > max || a > b) throw new Error('Rango inválido: "' + tok + '" (el PDF tiene ' + max + ' páginas).');
        const idx = []; for (let p = a; p <= b; p++) idx.push(p - 1); out.push({ label: a + '-' + b, idx });
      } else if (/^\d+$/.test(tok)) {
        const a = +tok; if (a < 1 || a > max) throw new Error('Página fuera de rango: "' + tok + '" (1–' + max + ').');
        out.push({ label: String(a), idx: [a - 1] });
      } else throw new Error('No entiendo "' + tok + '". Usa por ejemplo 1-3, 4, 5-8.');
    });
    if (!out.length) throw new Error('Indica al menos un rango (ej: 1-3, 4, 5-8).');
    return out;
  }
  async function doSplit() {
    if (!pdfSplitFile) return;
    const btn = $('#splitBtn'); const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Separando…';
    showInline($('#splitMsg'), 'info', 'Cargando librerías y separando…');
    try {
      const PDFLib = await ensurePDFLib();
      const buf = await readFile(pdfSplitFile);
      const src = await PDFLib.PDFDocument.load(buf, { ignoreEncryption: true });
      const n = src.getPageCount();
      let partes;
      if ($('#splitMode').value === 'ranges') {
        partes = parseRanges($('#splitRanges').value, n);
      } else {
        partes = []; for (let i = 0; i < n; i++) partes.push({ label: 'pagina_' + (i + 1), idx: [i] });
      }
      const JSZip = await ensureJSZip();
      const zip = new JSZip();
      const base = (pdfSplitFile.name || 'documento').replace(/\.pdf$/i, '');
      for (let k = 0; k < partes.length; k++) {
        const nd = await PDFLib.PDFDocument.create();
        const pages = await nd.copyPages(src, partes[k].idx);
        pages.forEach(p => nd.addPage(p));
        const bytes = await nd.save();
        zip.file(base + '_' + partes[k].label.replace(/[^0-9a-zA-Z_-]+/g, '-') + '.pdf', bytes);
      }
      const blob = await zip.generateAsync({ type: 'blob' });
      downloadBlob(blob, base + '_separado_' + stamp() + '.zip');
      showInline($('#splitMsg'), 'ok', '✓ Generados ' + partes.length + ' PDF en un .zip. Descarga iniciada.');
    } catch (e) {
      showInline($('#splitMsg'), 'err', '❌ ' + (e.message || e));
    } finally { btn.textContent = orig; btn.disabled = !pdfSplitFile; }
  }

  // ---- Conexión de eventos de UI -----------------------------------------
  // Navegación del panel lateral (vistas + inventario por estado)
  document.querySelectorAll('.tab').forEach(btn => btn.addEventListener('click', () => goTab(btn.dataset.tab)));
  document.querySelectorAll('.invcat').forEach(btn => btn.addEventListener('click', () => goInventario(btn.dataset.cat)));
  const LS_SIDE = 'mp_sidebar_collapsed';
  function applySidebar() { try { $('#layout').classList.toggle('collapsed', localStorage.getItem(LS_SIDE) === '1'); } catch (e) { } }
  $('#sideToggle').addEventListener('click', () => {
    const c = !$('#layout').classList.contains('collapsed');
    $('#layout').classList.toggle('collapsed', c);
    try { localStorage.setItem(LS_SIDE, c ? '1' : '0'); } catch (e) { }
  });
  applySidebar();

  // Selector de vista de pendientes (triple foco)
  document.querySelectorAll('.pview').forEach(btn => btn.addEventListener('click', () => {
    const v = btn.dataset.pview;
    document.querySelectorAll('.pview').forEach(b => b.classList.toggle('active', b === btn));
    document.querySelectorAll('.pview-panel').forEach(p => p.classList.toggle('active', p.id === 'view-' + v));
  }));

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
  $('#inventarioFilter').addEventListener('input', renderInventario);

  // Utilidades PDF
  $('#mergeFiles').addEventListener('change', e => { addMergeFiles(e.target.files); e.target.value = ''; });
  $('#mergeBtn').addEventListener('click', doMerge);
  $('#mergeClear').addEventListener('click', () => { pdfMerge.length = 0; renderMergeList(); showInline($('#mergeMsg'), '', ''); });
  $('#splitFile').addEventListener('change', e => { const f = e.target.files && e.target.files[0]; showInline($('#splitMsg'), '', ''); onSplitFile(f); });
  $('#splitMode').addEventListener('change', () => { $('#splitRangesField').style.display = $('#splitMode').value === 'ranges' ? 'block' : 'none'; });
  $('#splitBtn').addEventListener('click', doSplit);

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
  $('#mClose').addEventListener('click', closeModal);
  $('#status').addEventListener('click', () => { $('#status').style.display = 'none'; });
  $('#modal').addEventListener('click', e => { if (e.target === $('#modal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#modal').classList.contains('show')) closeModal(); });

  // Evento correctivo
  $('#cTipo').addEventListener('change', renderCFields);
  $('#cSave').addEventListener('click', saveCorrectivo);
  $('#cCancel').addEventListener('click', closeCModal);
  $('#cClose').addEventListener('click', closeCModal);
  $('#cmodal').addEventListener('click', e => { if (e.target === $('#cmodal')) closeCModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#cmodal').classList.contains('show')) closeCModal(); });
  $('#clearCorr').addEventListener('click', () => {
    if (!state.correctivos.length) return;
    if (!confirm('¿Eliminar los ' + state.correctivos.length + ' eventos correctivos registrados?')) return;
    state.correctivos = []; saveCorrectivos(); renderCorrectivos();
    if (state.selDetalleEq) renderDetalle(state.selDetalleEq);
  });

  // Pendientes
  ['#pTipo', '#pFecha', '#pRespAdmin', '#pRespEjec', '#pObs'].forEach(s => {
    $(s).addEventListener('input', updatePSave); $(s).addEventListener('change', updatePSave);
  });
  $('#pQuick3').addEventListener('click', () => { const t = new Date(); t.setHours(0, 0, 0, 0); setPFecha(addBusinessDays(t, 3)); });
  $('#pQuick7').addEventListener('click', () => { const t = new Date(); t.setHours(0, 0, 0, 0); t.setDate(t.getDate() + 7); setPFecha(t); });
  $('#pSave').addEventListener('click', savePendiente);
  $('#pCancel').addEventListener('click', closePModal);
  $('#pClose').addEventListener('click', closePModal);
  $('#pmodal').addEventListener('click', e => { if (e.target === $('#pmodal')) closePModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#pmodal').classList.contains('show')) closePModal(); });
  $('#pendientesFilter').addEventListener('input', renderPendientes);
  // Gestión de pendientes
  $('#gEstado').addEventListener('change', () => { if (state.gPend) { state.gPend.estado = $('#gEstado').value; gPersist(); renderGModal(); } });
  $('#gTareaAdd').addEventListener('click', gAddTarea);
  $('#gTareaTexto').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); gAddTarea(); } });
  $('#gActAdd').addEventListener('click', gAddAct);
  $('#gClose').addEventListener('click', closeGModal);
  $('#gCerrar').addEventListener('click', closeGModal);
  $('#gmodal').addEventListener('click', e => { if (e.target === $('#gmodal')) closeGModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#gmodal').classList.contains('show')) closeGModal(); });
  $('#clearPend').addEventListener('click', () => {
    if (!state.pendientes.length) return;
    if (!confirm('¿Eliminar los ' + state.pendientes.length + ' pendientes registrados?')) return;
    state.pendientes = []; savePendientes(); renderPendientes();
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
  $('#sheetsPull').addEventListener('click', () => pullFromSheets(false));
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
  if (GAS) {
    // Servida desde Apps Script: conexión directa, sin URL.
    $('#sheetsUrl').value = ''; $('#sheetsUrl').placeholder = 'Conectado a esta planilla (Apps Script) — no necesitas URL';
    $('#sheetsUrl').disabled = true;
    $('#sheetsSave').style.display = 'none';
    showInline($('#sheetsStatus'), 'ok', '✓ Conectado a esta planilla (Apps Script). Usa “Enviar” y “Traer” directamente.');
  }

  // Recuperar registros guardados en este navegador (persistencia entre sesiones)
  state.registros = loadRegistros();
  state.correctivos = loadCorrectivos();
  state.pendientes = loadPendientes();
  state.archivos = loadArchivos();
  renderCorrectivos();
  renderPendientes();
  renderSidebarCounts();
  revealRegistrarIfData();
  if (state.registros.length || state.correctivos.length || state.pendientes.length) {
    setStatus('ℹ️ Tienes <b>' + state.registros.length + '</b> mantención(es), <b>' + state.correctivos.length +
      '</b> correctivo(s) y <b>' + state.pendientes.length + '</b> pendiente(s) guardados. ' +
      'Cargue el archivo para incluirlos al descargar.', 'info');
  }
  // Traer automáticamente lo guardado en Google Sheets (si hay URL configurada)
  if (sheetsReady()) pullFromSheets(true);
})();
