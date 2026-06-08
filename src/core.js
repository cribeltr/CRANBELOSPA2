/* ============================================================================
 *  Núcleo de procesamiento — Programación MP 2026
 *  Funciona igual en Node (require('exceljs')) y en el navegador (ExcelJS UMD).
 *  Expone: window.MP / module.exports = { parseWorkbook, buildOutputWorkbook,
 *           EJECUTORES, CAUSAL, ... }
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MP = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---- Constantes de estructura ------------------------------------------
  const HEADER_ROW = 7;
  const FIRST_DATA_ROW = 8;

  const MONTHS = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const MONTHS_FULL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  // Columnas de datos del equipo (1-based). Se excluye S(19)=Responsable MP.
  // Q(17)=Observación se incluye a pedido del usuario.
  const COL = {
    fam: 1, id: 2, carpeta: 3, inv: 4, equipo: 5, servicio: 6, unidad: 7,
    ubicacion: 8, procedencia: 9, marca: 10, modelo: 11, serie: 12,
    anio: 13, vur: 14, clasif: 15, enubaja: 16, observacion: 17, frecuencia: 18
  };
  // PMP_2026: meses en columnas únicas T..AE (20..31)
  const pmpMonthCol = i => 20 + i;
  // Registro_MP-2026: cada mes con dos subcolumnas P (programa) y R (resultado)
  const regPCol = i => 20 + i * 2;   // T, V, X, ...
  const regRCol = i => 21 + i * 2;   // U, W, Y, ...

  const PMP_SHEET = 'PMP_2026';
  const REG_SHEET = 'Registro_MP-2026';

  // ---- Diccionarios de decodificación ------------------------------------
  const PROG = {
    'X':  'Mantención Preventiva Programada',
    'R':  'Mantención Preventiva Reprogramada',
    'RA': 'Mantención Preventiva Reprogramada de Año Anterior',
    'PM': 'Puesta en Marcha'
  };

  const CAUSAL = {
    C1: 'Imposibilidad de desocupar el equipo del paciente por indicación clínica',
    C2: 'Equipo en servicio técnico',
    C3: 'Equipo no operativo, a la espera de repuestos o accesorios',
    C4: 'Equipo en préstamo a otro hospital o institución',
    C5: 'No disponibilidad de horas hombre del funcionario SEC por alta carga laboral',
    C6: 'No disponibilidad de horas hombre del servicio técnico externo',
    C7: 'Ausencia justificada del funcionario SEC superior a 15 días',
    C8: 'Contingencia hospitalaria'
  };
  const CAUSAL_30 = new Set(['C1','C5','C6','C7','C8']);   // reprogramar <= 30 días
  const CAUSAL_REINTEGRO = new Set(['C2','C3','C4']);      // sin nueva fecha

  const RESULT_TXT = {
    'SI': 'Mantención Preventiva Realizada',
    'SI-RA': 'Mantención de Año Anterior Realizada',
    'FS': 'Fuera de Servicio',
    'NO': 'No Realizada',
    'NU': 'No Ubicable',
    'BAJA': 'Equipo Dado de Baja'
  };

  const EJECUTORES = [
    'Carlos Bahamondes Seguel',
    'Cristián Beltrán Oviedo',
    'Cristina Rozas Urrutia',
    'Daniel Díaz Neira',
    'Ignacio Berner Bergara',
    'Macarena Toledo',
    'Marco Ulloa',
    'Matías Soazo Garrido',
    'Ricardo Matus Aroca',
    'Tito Millapán Riquelme',
    'Personal Externo'
  ];

  // Opciones del desplegable "Estado Final del Equipo" (se llena manualmente).
  const ESTADO_FINAL_OPCIONES = ['Operativo', 'No operativo', 'En servicio técnico'];

  // Opciones del desplegable "Resultado (R)".
  const RESULTADO_OPCIONES = ['Si', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'Si-RA', 'FS', 'No', 'NU', 'Baja'];

  // ---- Helpers de lectura de celdas (ExcelJS) ----------------------------
  // Devuelve el valor MOSTRADO de la celda como texto (resuelve fórmulas con
  // su resultado en caché y conserva ceros a la izquierda guardados como texto).
  function cellText(cell) {
    if (!cell) return '';
    let v = cell.value;
    if (v === null || v === undefined) {
      const t = cell.text;
      return (t === null || t === undefined) ? '' : String(t).trim();
    }
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(p => p.text).join('').trim();
      if ('result' in v) v = v.result;
      else if ('text' in v) v = v.text;
      else if ('hyperlink' in v) v = v.text || v.hyperlink;
      else { const t = cell.text; return (t == null) ? '' : String(t).trim(); }
    }
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }
  // "vacío" = celda sin contenido útil ("" o "0")
  function isEmpty(s) { return s === '' || s === '0'; }
  function toNum(s) { const n = Number(s); return isNaN(n) ? null : n; }

  function isDataRow(ws, r) {
    const id = cellText(ws.getRow(r).getCell(COL.id));
    const eq = cellText(ws.getRow(r).getCell(COL.equipo));
    return toNum(id) !== null && !isEmpty(eq);
  }

  function readEquipo(ws, r) {
    const g = c => cellText(ws.getRow(r).getCell(c));
    return {
      familia: g(COL.fam), id: g(COL.id), carpeta: g(COL.carpeta), inv: g(COL.inv),
      equipo: g(COL.equipo), servicio: g(COL.servicio), unidad: g(COL.unidad),
      ubicacion: g(COL.ubicacion), procedencia: g(COL.procedencia), marca: g(COL.marca),
      modelo: g(COL.modelo), serie: g(COL.serie), anio: g(COL.anio), vur: g(COL.vur),
      clasif: g(COL.clasif), enubaja: g(COL.enubaja),
      observacion: g(COL.observacion), frecuencia: g(COL.frecuencia)
    };
  }

  // ---- Decodificadores ----------------------------------------------------
  function decodeResultado(rRaw) {
    if (isEmpty(rRaw)) return '';
    const u = rRaw.toUpperCase();
    if (/^C[1-8]$/.test(u)) return 'Mantención Preventiva Reprogramada';
    return RESULT_TXT[u] || rRaw;
  }
  function causalCode(rRaw) {
    const u = (rRaw || '').toUpperCase();
    return /^C[1-8]$/.test(u) ? u : '';
  }
  function reglaReprog(cCode) {
    if (!cCode) return '';
    if (CAUSAL_30.has(cCode)) return 'Reprogramar dentro de los 30 días (mes origen mantiene X + causal; mes destino lleva R)';
    if (CAUSAL_REINTEGRO.has(cCode)) return 'Sin nueva fecha; registrar en el mes real de ejecución al reintegrarse el equipo';
    return '';
  }
  function estado(pRaw, rRaw) {
    const u = (rRaw || '').toUpperCase();
    if (!isEmpty(rRaw)) {
      if (u === 'SI') return 'Realizada';
      if (u === 'SI-RA') return 'Realizada (Año Anterior)';
      if (/^C[1-8]$/.test(u)) return 'Reprogramada';
      if (u === 'FS') return 'Fuera de Servicio';
      if (u === 'NO') return 'No Realizada';
      if (u === 'NU') return 'No Ubicable';
      if (u === 'BAJA') return 'Baja';
      return rRaw;
    }
    const p = (pRaw || '').toUpperCase();
    if (p === 'PM') return 'Puesta en Marcha';
    if (p === 'R')  return 'Pendiente (Reprogramada)';
    if (p === 'RA') return 'Pendiente (Año Anterior)';
    if (p === 'X')  return 'Pendiente';
    return 'Pendiente';
  }

  // ---- Construcción de un evento (fila) ----------------------------------
  // Centraliza el armado del objeto-evento para reutilizarlo tanto en la
  // generación masiva como en el registro manual de mantenciones.
  function makeEvent(eq, i, pVal, rVal, extra) {
    extra = extra || {};
    pVal = pVal || ''; rVal = rVal || '';
    const cCode = causalCode(rVal);
    return {
      familia: eq.familia, id: eq.id, carpeta: eq.carpeta, inv: eq.inv,
      equipo: eq.equipo, servicio: eq.servicio, unidad: eq.unidad,
      ubicacion: eq.ubicacion, procedencia: eq.procedencia, marca: eq.marca,
      modelo: eq.modelo, serie: eq.serie, anio: eq.anio, vur: eq.vur,
      clasif: eq.clasif, enubaja: eq.enubaja,
      observacion: ('observacion' in extra) ? extra.observacion : eq.observacion,
      frecuencia: eq.frecuencia,
      mes: MONTHS_FULL[i], nMes: i + 1,
      programa: pVal,
      tipoPrograma: PROG[pVal.toUpperCase()] || (pVal || ''),
      resultado: rVal,
      detalleResultado: decodeResultado(rVal),
      fechaEjecucion: ('fechaEjecucion' in extra) ? extra.fechaEjecucion : '',
      causal: cCode,
      causalDesc: cCode ? CAUSAL[cCode] : '',
      regla: reglaReprog(cCode),
      estado: estado(pVal, rVal),
      estadoFinal: ('estadoFinal' in extra) ? extra.estadoFinal : '',
      ejecutor: ('ejecutor' in extra) ? extra.ejecutor : ''
    };
  }

  // ---- Parseo principal: workbook -> equipos + eventos -------------------
  function parseWorkbook(wb) {
    const pmp = wb.getWorksheet(PMP_SHEET);
    const reg = wb.getWorksheet(REG_SHEET);
    if (!pmp) throw new Error('No se encontró la hoja "' + PMP_SHEET + '".');

    const warnings = [];
    if (!reg) warnings.push('No se encontró la hoja "' + REG_SHEET + '"; no se incluirán resultados de ejecución.');

    // Mapa de resultados (R) del Registro, por ID
    const regResByID = new Map();
    if (reg) {
      const lastR = reg.rowCount;
      for (let r = FIRST_DATA_ROW; r <= lastR; r++) {
        if (!isDataRow(reg, r)) continue;
        const id = cellText(reg.getRow(r).getCell(COL.id));
        const arr = [];
        for (let i = 0; i < 12; i++) {
          const v = cellText(reg.getRow(r).getCell(regRCol(i)));
          arr.push(isEmpty(v) ? '' : v);
        }
        regResByID.set(id, arr);
      }
    }

    const equipos = [];
    const events = [];
    const last = pmp.rowCount;
    for (let r = FIRST_DATA_ROW; r <= last; r++) {
      if (!isDataRow(pmp, r)) continue;
      const eq = readEquipo(pmp, r);
      const res = regResByID.get(eq.id) || [];
      const prog = [];
      for (let i = 0; i < 12; i++) {
        const v = cellText(pmp.getRow(r).getCell(pmpMonthCol(i)));
        prog.push(isEmpty(v) ? '' : v);
      }
      eq.prog = prog;                                   // 12 valores de programación (X/R/RA/PM)
      eq.res = [];
      for (let i = 0; i < 12; i++) eq.res.push(res[i] || '');
      equipos.push(eq);

      for (let i = 0; i < 12; i++) {
        const pVal = eq.prog[i], rVal = eq.res[i];
        if (isEmpty(pVal) && isEmpty(rVal)) continue;   // sin evento ese mes
        events.push(makeEvent(eq, i, pVal, rVal));
      }
    }

    // Ordenar por ID y luego por mes
    events.sort((a, b) => (toNum(a.id) - toNum(b.id)) || (a.nMes - b.nMes));

    return { events, equipos, warnings };
  }

  // ---- Helpers para el registro manual de mantenciones -------------------
  // Busca equipos por N° de Serie o N° de Inventario (exacto y luego parcial).
  function findEquipos(equipos, query) {
    const q = String(query == null ? '' : query).trim().toLowerCase();
    if (!q) return [];
    const norm = s => String(s == null ? '' : s).trim().toLowerCase();
    const exact = equipos.filter(e => norm(e.serie) === q || norm(e.inv) === q);
    if (exact.length) return exact;
    return equipos.filter(e =>
      norm(e.serie).includes(q) || norm(e.inv).includes(q) || norm(e.id) === q);
  }

  // Meses (0-11) en que el equipo tiene mantención programada (P no vacío).
  function programmedMonths(eq) {
    const out = [];
    if (eq && eq.prog) for (let i = 0; i < 12; i++) if (!isEmpty(eq.prog[i])) out.push(i);
    return out;
  }

  // Arma un evento a partir del formulario de registro manual.
  function buildRegistro(eq, monthIndex, data) {
    data = data || {};
    const pVal = (eq.prog && eq.prog[monthIndex]) ? eq.prog[monthIndex] : '';
    return makeEvent(eq, monthIndex, pVal, data.resultado || '', {
      fechaEjecucion: data.fecha || '',
      observacion: data.observacion || '',
      ejecutor: data.ejecutor || '',
      estadoFinal: data.estadoFinal || ''
    });
  }

  // ---- Estadísticas para hoja Resumen ------------------------------------
  function buildStats(events) {
    const byEstado = new Map();
    const byMes = MONTHS_FULL.map(() => ({ prog: 0, real: 0, reprog: 0, pend: 0, otras: 0 }));
    for (const e of events) {
      byEstado.set(e.estado, (byEstado.get(e.estado) || 0) + 1);
      const m = byMes[e.nMes - 1];
      if (e.programa) m.prog++;
      const st = e.estado;
      if (st.startsWith('Realizada')) m.real++;
      else if (st === 'Reprogramada') m.reprog++;
      else if (st.startsWith('Pendiente') || st === 'Puesta en Marcha') m.pend++;
      else m.otras++;
    }
    return { byEstado, byMes };
  }

  return {
    HEADER_ROW, FIRST_DATA_ROW, MONTHS, MONTHS_FULL, COL,
    PROG, CAUSAL, CAUSAL_30, CAUSAL_REINTEGRO, RESULT_TXT,
    EJECUTORES, ESTADO_FINAL_OPCIONES, RESULTADO_OPCIONES,
    cellText, isEmpty, toNum, isDataRow, readEquipo,
    decodeResultado, causalCode, reglaReprog, estado,
    parseWorkbook, buildStats, makeEvent,
    findEquipos, programmedMonths, buildRegistro
  };
});
