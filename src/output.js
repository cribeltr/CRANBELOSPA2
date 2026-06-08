/* ============================================================================
 *  Constructor del Excel de salida (una fila por evento) — ExcelJS.
 *  buildOutputWorkbook(ExcelJS, MP, events, meta) -> ExcelJS.Workbook
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MPOUT = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const COLOR = {
    header: 'FF15616D', headerText: 'FFFFFFFF',
    realizada: 'FFC6EFCE', reprogramada: 'FFFCE4A6', pendiente: 'FFFFF2CC',
    pm: 'FFDDEBF7', fueraServicio: 'FFFFC7CE', noReal: 'FFFFC7CE',
    nu: 'FFE4DFEC', baja: 'FFD9D9D9', band: 'FFF2F7F8',
    catTitle: 'FF2A6F77'
  };
  function estadoFill(st) {
    if (st.startsWith('Realizada')) return COLOR.realizada;
    if (st === 'Reprogramada') return COLOR.reprogramada;
    if (st === 'Puesta en Marcha') return COLOR.pm;
    if (st === 'Fuera de Servicio') return COLOR.fueraServicio;
    if (st === 'No Realizada') return COLOR.noReal;
    if (st === 'No Ubicable') return COLOR.nu;
    if (st === 'Baja') return COLOR.baja;
    if (st.startsWith('Pendiente')) return COLOR.pendiente;
    return null;
  }
  const solid = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
  const thin = { style: 'thin', color: { argb: 'FFD0D7DE' } };
  const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

  // Definición de columnas de la hoja Eventos
  const COLS = [
    { key: 'familia',          header: 'Familia',                 w: 16, t: 's' },
    { key: 'id',               header: 'ID',                      w: 7,  t: 'n' },
    { key: 'carpeta',          header: 'N° Carpeta',              w: 12, t: 's' },
    { key: 'inv',              header: 'N° Inventario',           w: 14, t: 's' },
    { key: 'equipo',           header: 'Equipo',                  w: 26, t: 's' },
    { key: 'servicio',         header: 'Servicio',                w: 22, t: 's' },
    { key: 'unidad',           header: 'Unidad',                  w: 22, t: 's' },
    { key: 'ubicacion',        header: 'Ubicación',               w: 16, t: 's' },
    { key: 'procedencia',      header: 'Procedencia',             w: 13, t: 's' },
    { key: 'marca',            header: 'Marca',                   w: 18, t: 's' },
    { key: 'modelo',           header: 'Modelo',                  w: 20, t: 's' },
    { key: 'serie',            header: 'N° Serie',                w: 18, t: 's' },
    { key: 'anio',             header: 'Año Instalación',         w: 10, t: 'n' },
    { key: 'vur',              header: 'Vida Útil Residual',      w: 10, t: 'n' },
    { key: 'clasif',           header: 'Clasificación',           w: 14, t: 's' },
    { key: 'enubaja',          header: 'ENU / Baja',              w: 11, t: 's' },
    { key: 'frecuencia',       header: 'Frecuencia MP',           w: 13, t: 's' },
    { key: 'observacion',      header: 'Observación',             w: 44, t: 's' },
    { key: 'mes',              header: 'Mes',                     w: 12, t: 's' },
    { key: 'nMes',             header: 'N° Mes',                  w: 7,  t: 'n' },
    { key: 'programa',         header: 'Programa (P)',            w: 11, t: 's' },
    { key: 'tipoPrograma',     header: 'Tipo de Programación',    w: 34, t: 's' },
    { key: 'resultado',        header: 'Resultado (R)',           w: 12, t: 's' },
    { key: 'detalleResultado', header: 'Detalle del Resultado',   w: 34, t: 's' },
    { key: 'fechaEjecucion',   header: 'Fecha de Ejecución',      w: 16, t: 'd' },
    { key: 'causal',           header: 'Causal Reprog.',          w: 11, t: 's' },
    { key: 'causalDesc',       header: 'Descripción de la Causal',w: 42, t: 's' },
    { key: 'regla',            header: 'Regla de Reprogramación', w: 42, t: 's' },
    { key: 'estado',           header: 'Estado',                  w: 22, t: 's' },
    { key: 'estadoFinal',      header: 'Estado Final del Equipo', w: 20, t: 's' },
    { key: 'ejecutor',         header: 'Ejecutor',                w: 24, t: 's' }
  ];
  const TEXT_KEYS = new Set(['carpeta', 'inv', 'serie', 'observacion']); // conservar tal cual (ceros a la izq. / texto)
  const NUM_KEYS = new Set(['id', 'anio', 'vur', 'nMes']);
  const DATE_KEYS = new Set(['fechaEjecucion']);                          // formato fecha (llenado manual)
  const DATE_FMT = 'dd-mm-yyyy';

  function styleHeaderRow(row) {
    row.height = 30;
    row.eachCell(c => {
      c.font = { bold: true, color: { argb: COLOR.headerText }, size: 10, name: 'Calibri' };
      c.fill = solid(COLOR.header);
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      c.border = borderAll;
    });
  }

  function buildOutputWorkbook(ExcelJS, MP, events, meta) {
    meta = meta || {};
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Generador de Eventos MP';
    wb.created = new Date();

    /* ---------- Hoja Catálogos (también fuente del desplegable) ---------- */
    const cat = wb.addWorksheet('Catalogos');
    cat.getColumn(1).width = 26; // Ejecutores (A) -> origen del dropdown
    cat.getColumn(2).width = 6;
    cat.getColumn(3).width = 58;
    cat.getColumn(5).width = 6;
    cat.getColumn(6).width = 58;
    cat.getColumn(8).width = 14;

    // Ejecutores en A1.. (A2:A12 = lista). Encabezado en A1.
    cat.getCell('A1').value = 'Ejecutores';
    MP.EJECUTORES.forEach((n, i) => { cat.getCell(i + 2, 1).value = n; }); // (fila, col)=A2..A12

    // Tabla de códigos de Programa (B..C)
    cat.getCell('B1').value = 'Programa (P)'; cat.getCell('C1').value = 'Significado';
    let rr = 2;
    Object.keys(MP.PROG).forEach(k => {
      cat.getCell(rr, 2).value = k; cat.getCell(rr, 3).value = MP.PROG[k]; rr++;
    });
    // Tabla de Resultados (E..F)
    cat.getCell('E1').value = 'Resultado (R)'; cat.getCell('F1').value = 'Significado';
    const resList = [
      ['Si', 'Mantención Preventiva Realizada'],
      ['C1 - C8', 'Mantención Preventiva Reprogramada (ver causales)'],
      ['Si-RA', 'Mantención de Año Anterior Realizada'],
      ['FS', 'Fuera de Servicio'],
      ['No', 'No Realizada'],
      ['NU', 'No Ubicable'],
      ['Baja', 'Equipo Dado de Baja']
    ];
    resList.forEach((x, i) => { cat.getCell(i + 2, 5).value = x[0]; cat.getCell(i + 2, 6).value = x[1]; });

    // Tabla de Causales (H..J)
    cat.getCell('H1').value = 'Causal'; cat.getCell('I1').value = 'Descripción'; cat.getCell('J1').value = 'Regla de reprogramación';
    cat.getColumn(9).width = 58; cat.getColumn(10).width = 50;
    let cr = 2;
    Object.keys(MP.CAUSAL).forEach(k => {
      cat.getCell(cr, 8).value = k;
      cat.getCell(cr, 9).value = MP.CAUSAL[k];
      cat.getCell(cr, 10).value = MP.CAUSAL_30.has(k)
        ? 'Reprogramar dentro de 30 días'
        : 'Sin nueva fecha; registrar al reintegrarse';
      cr++;
    });
    // Estilo encabezados catálogo
    ['A1', 'B1', 'C1', 'E1', 'F1', 'H1', 'I1', 'J1'].forEach(a => {
      const c = cat.getCell(a);
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = solid(COLOR.catTitle);
    });

    /* ---------- Hoja Eventos -------------------------------------------- */
    const ws = wb.addWorksheet('Eventos', {
      views: [{ state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'A2' }]
    });
    ws.columns = COLS.map(c => ({ header: c.header, key: c.key, width: c.w }));
    styleHeaderRow(ws.getRow(1));

    events.forEach((e, idx) => {
      const row = ws.addRow(e);
      const band = idx % 2 === 1;
      row.eachCell({ includeEmpty: true }, (cell, col) => {
        const def = COLS[col - 1];
        cell.border = borderAll;
        cell.font = { size: 10, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', wrapText: false };
        if (band) cell.fill = solid(COLOR.band);
        if (!def) return;
        if (TEXT_KEYS.has(def.key)) { cell.numFmt = '@'; }            // texto literal
        else if (NUM_KEYS.has(def.key)) {
          const n = MP.toNum(cell.value);
          if (n !== null) cell.value = n;                              // numérico real
        } else if (DATE_KEYS.has(def.key)) {
          if (cell.value === '' || cell.value === null) cell.value = null; // celda vacía limpia
          cell.numFmt = DATE_FMT;                                      // lista para escribir la fecha
        }
      });
      // Conservar N° Serie / Inventario / Carpeta como TEXTO (ceros a la izq.)
      TEXT_KEYS.forEach(k => {
        const idxCol = COLS.findIndex(c => c.key === k) + 1;
        const cell = row.getCell(idxCol);
        if (cell.value !== null && cell.value !== undefined && cell.value !== '')
          cell.value = String(cell.value);
      });
      // Color del Estado
      const estCol = COLS.findIndex(c => c.key === 'estado') + 1;
      const f = estadoFill(e.estado);
      if (f) { row.getCell(estCol).fill = solid(f); row.getCell(estCol).font = { size: 10, bold: true, name: 'Calibri' }; }
      // Resaltar causal
      if (e.causal) {
        const cCol = COLS.findIndex(c => c.key === 'causal') + 1;
        row.getCell(cCol).fill = solid(COLOR.reprogramada);
        row.getCell(cCol).font = { size: 10, bold: true, name: 'Calibri' };
      }
    });

    // Autofiltro + congelado ya aplicado
    const lastColLetter = ws.getColumn(COLS.length).letter;
    ws.autoFilter = { from: 'A1', to: lastColLetter + '1' };

    // Validación de datos (desplegable) en la columna Ejecutor — un solo rango
    const ejeCol = ws.getColumn(COLS.findIndex(c => c.key === 'ejecutor') + 1);
    const ejeLetter = ejeCol.letter;
    const lastRow = events.length + 1;
    if (lastRow >= 2) {
      ws.dataValidations.add(ejeLetter + '2:' + ejeLetter + lastRow, {
        type: 'list', allowBlank: true,
        formulae: ['Catalogos!$A$2:$A$' + (MP.EJECUTORES.length + 1)],
        showErrorMessage: true,
        errorStyle: 'warning',
        errorTitle: 'Ejecutor no válido',
        error: 'Seleccione un ejecutor de la lista.'
      });
    }

    // Validación (desplegable) en "Estado Final del Equipo": Operativo / No operativo
    const efLetter = ws.getColumn(COLS.findIndex(c => c.key === 'estadoFinal') + 1).letter;
    if (lastRow >= 2) {
      const efRange = efLetter + '2:' + efLetter + lastRow;
      ws.dataValidations.add(efRange, {
        type: 'list', allowBlank: true,
        formulae: ['"' + MP.ESTADO_FINAL_OPCIONES.join(',') + '"'],
        showErrorMessage: true,
        errorStyle: 'warning',
        errorTitle: 'Valor no válido',
        error: 'Seleccione "Operativo" o "No operativo".'
      });
      // Coloreado automático según el valor elegido
      ws.addConditionalFormatting({
        ref: efRange,
        rules: [
          { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"Operativo"'],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLOR.realizada } },
                     font: { color: { argb: 'FF1B7F3B' }, bold: true } } },
          { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"No operativo"'],
            style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLOR.fueraServicio } },
                     font: { color: { argb: 'FFB3261E' }, bold: true } } }
        ]
      });
    }

    /* ---------- Hoja Resumen ------------------------------------------- */
    const stats = MP.buildStats(events);
    const sum = wb.addWorksheet('Resumen');
    sum.getColumn(1).width = 26; sum.getColumn(2).width = 14;
    sum.getColumn(3).width = 14; sum.getColumn(4).width = 14;
    sum.getColumn(5).width = 14; sum.getColumn(6).width = 14;
    sum.getCell('A1').value = 'Resumen de Eventos — Programación MP 2026';
    sum.getCell('A1').font = { bold: true, size: 14, color: { argb: COLOR.catTitle } };
    sum.getCell('A2').value = 'Generado:';
    sum.getCell('B2').value = new Date();
    sum.getCell('B2').numFmt = 'dd-mm-yyyy hh:mm';
    sum.getCell('A3').value = 'Equipos procesados:'; sum.getCell('B3').value = meta.equipos || 0;
    sum.getCell('A4').value = 'Total de eventos:';    sum.getCell('B4').value = events.length;

    let r = 6;
    sum.getCell('A' + r).value = 'Eventos por Estado';
    sum.getCell('A' + r).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sum.getCell('A' + r).fill = solid(COLOR.header);
    sum.getCell('B' + r).value = 'Cantidad';
    sum.getCell('B' + r).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sum.getCell('B' + r).fill = solid(COLOR.header);
    r++;
    const estados = [...stats.byEstado.entries()].sort((a, b) => b[1] - a[1]);
    estados.forEach(([k, v]) => {
      sum.getCell('A' + r).value = k;
      sum.getCell('B' + r).value = v;
      const f = estadoFill(k); if (f) sum.getCell('A' + r).fill = solid(f);
      r++;
    });

    r += 1;
    const head = ['Mes', 'Programadas', 'Realizadas', 'Reprogramadas', 'Pend./PM', 'Otras'];
    head.forEach((h, i) => {
      const c = sum.getCell(r, i + 1);
      c.value = h; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = solid(COLOR.header);
    });
    r++;
    stats.byMes.forEach((m, i) => {
      sum.getCell(r, 1).value = MP.MONTHS_FULL[i];
      sum.getCell(r, 2).value = m.prog;
      sum.getCell(r, 3).value = m.real;
      sum.getCell(r, 4).value = m.reprog;
      sum.getCell(r, 5).value = m.pend;
      sum.getCell(r, 6).value = m.otras;
      r++;
    });

    // Orden de hojas: Eventos primero
    wb.worksheets.forEach((w, i) => { w.orderNo = i; });
    return wb;
  }

  return { buildOutputWorkbook };
});
