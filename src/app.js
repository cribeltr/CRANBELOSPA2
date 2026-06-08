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
  const state = { equipos: [], events: [], registros: [], searchResults: [], selEq: null, selMonth: null };

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
      const wb = MPOUT.buildOutputWorkbook(ExcelJS, MP, events, { equipos: state.equipos.length });
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

  // ---- Búsqueda de equipos -----------------------------------------------
  function resetSearch() {
    $('#searchInput').value = '';
    $('#searchResults').innerHTML = '';
    showInline($('#searchMsg'), '', '');
  }

  function doSearch() {
    const q = $('#searchInput').value;
    const results = MP.findEquipos(state.equipos, q);
    state.searchResults = results;
    const box = $('#searchResults');
    if (!q.trim()) { showInline($('#searchMsg'), 'err', 'Ingrese un N° de Serie o N° de Inventario.'); box.innerHTML = ''; return; }
    if (!results.length) { showInline($('#searchMsg'), 'err', 'Sin coincidencias para "' + esc(q) + '".'); box.innerHTML = ''; return; }
    showInline($('#searchMsg'), 'ok', results.length + ' equipo(s) encontrado(s).');
    box.innerHTML = results.map((e, i) =>
      '<div class="resitem">' +
        '<div class="info"><b>' + esc(e.equipo) + '</b> · ' + esc(e.marca) + ' ' + esc(e.modelo) +
          '<div class="meta">Serie: ' + esc(e.serie || '—') + ' · Inv: ' + esc(e.inv || '—') +
          ' · ' + esc(e.servicio) + ' · ' + esc(e.unidad) + '</div></div>' +
        '<button class="btn btn-primary btn-sm" data-idx="' + i + '">Registrar mantención</button>' +
      '</div>').join('');
    box.querySelectorAll('button[data-idx]').forEach(b =>
      b.addEventListener('click', () => openModal(state.searchResults[+b.dataset.idx])));
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
  // en cualquier otro caso queda en blanco y deshabilitado (no obligatorio).
  function updateEstadoFinal() {
    const isSi = $('#mResultado').value === 'Si';
    const sel = $('#mEstadoFinal');
    sel.disabled = !isSi;
    if (!isSi) sel.value = '';
    $('#mEstadoFinalLabel').innerHTML = isSi
      ? 'Estado final del equipo <span class="req">*</span>'
      : 'Estado final del equipo <span class="muted">(solo si el resultado es “Si”)</span>';
  }

  function updateSave() {
    const isSi = $('#mResultado').value === 'Si';
    const efOk = !isSi || !!$('#mEstadoFinal').value;     // obligatorio solo cuando Resultado = Si
    $('#mSave').disabled = !(state.selMonth !== null &&
      $('#mResultado').value && $('#mEjecutor').value && efOk);
  }

  function saveRegistro() {
    const eq = state.selEq, mi = state.selMonth;
    if (!eq || mi === null) return;
    const isSi = $('#mResultado').value === 'Si';
    if (!$('#mResultado').value || !$('#mEjecutor').value || (isSi && !$('#mEstadoFinal').value)) {
      showInline($('#modalMsg'), 'err', 'Complete los campos obligatorios (*).'); return;
    }
    const p = $('#mFecha').value.split('-');
    const fecha = new Date(+p[0], +p[1] - 1, +p[2]);
    const reg = MP.buildRegistro(eq, mi, {
      fecha,
      resultado: $('#mResultado').value,
      observacion: $('#mObs').value.trim(),
      ejecutor: $('#mEjecutor').value,
      estadoFinal: isSi ? $('#mEstadoFinal').value : ''   // en blanco si el resultado no es "Si"
    });
    state.registros.push(reg);
    saveRegistros();   // persistir en el navegador
    closeModal();
    renderRegistry();
    renderPreview();   // reflejar el registro en la vista previa y en el Excel masivo
    renderDiscrepancias();
    setStatus('✅ Mantención registrada: <b>' + esc(eq.equipo) + '</b> — ' + esc(reg.mes) +
      ' (' + esc(fmtDate(fecha)) + '). Quedará incluida al descargar el Excel.', 'ok');
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
    $('#registryCount').textContent = state.registros.length;
    const cols = ['#', 'ID', 'Equipo', 'Serie / Inv', 'Mes', 'Fecha', 'Programa', 'Resultado', 'Ejecutor', 'Estado Final', ''];
    const head = '<tr>' + cols.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
    const rows = state.registros.map((e, i) =>
      '<tr>' +
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
      '</tr>').join('');
    const t = $('#registryTable');
    t.innerHTML = head + rows;
    t.querySelectorAll('button[data-del]').forEach(b =>
      b.addEventListener('click', () => {
        state.registros.splice(+b.dataset.del, 1);
        saveRegistros(); renderRegistry(); renderPreview(); renderDiscrepancias();
      }));
  }

  // ---- Conexión de eventos de UI -----------------------------------------
  $('#pick').addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
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
  $('#searchInput').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });

  $('#mFecha').addEventListener('change', onDateChange);
  $('#mFecha').addEventListener('input', onDateChange);
  $('#mResultado').addEventListener('change', () => { updateEstadoFinal(); updateSave(); });
  ['#mEjecutor', '#mEstadoFinal'].forEach(s => $(s).addEventListener('change', updateSave));
  $('#mSave').addEventListener('click', saveRegistro);
  $('#mCancel').addEventListener('click', closeModal);
  $('#modal').addEventListener('click', e => { if (e.target === $('#modal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#modal').classList.contains('show')) closeModal(); });

  initSelects();

  // Recuperar registros guardados en este navegador (persistencia entre sesiones)
  state.registros = loadRegistros();
  if (state.registros.length) {
    setStatus('ℹ️ Tienes <b>' + state.registros.length + '</b> mantención(es) registradas guardadas en este navegador. ' +
      'Cargue el archivo para verlas, compararlas e incluirlas al descargar.', 'info');
  }
})();
