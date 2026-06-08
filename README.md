# Generador de Eventos — Programación MP 2026

Aplicación **HTML de un solo archivo** que convierte el libro *Programación MP 2026*
(equipos médicos críticos del Hospital Hernán Henríquez Aravena) en un Excel con
**una fila por evento** de mantención preventiva.

Todo el procesamiento ocurre **en el navegador**: el archivo nunca se sube a
ningún servidor y la herramienta funciona **sin conexión a internet**.

---

## Uso rápido

1. Abra **`index.html`** con un navegador moderno (Chrome, Edge o Firefox) — basta
   con hacer doble clic en el archivo.
2. Arrastre el archivo **Programación MP 2026** (`.xlsm` o `.xlsx`) a la zona indicada,
   o pulse **Seleccionar archivo**.
3. Revise el resumen (equipos y eventos detectados) y pulse
   **⬇️ Descargar Excel (una fila por evento)**.

Se descarga `Eventos_MP_2026_AAAAMMDD.xlsx`.

---

## Registrar mantención (modo interactivo)

Tras cargar el archivo se habilita la tarjeta **🔧 Registrar mantención**:

1. **Busque** el equipo por su **N° de Serie** o **N° de Inventario**.
2. Pulse **Registrar mantención** en el resultado: se abre una **nueva vista**.
3. Ingrese la **fecha de ejecución**. Si el mes **no coincide** con un mes
   programado para ese equipo, **el sistema no deja avanzar** (e indica los meses
   programados).
4. Con un mes válido se **carga el tipo de programa** automáticamente.
5. Complete **Resultado** (lista), **Observación** (opcional) y **Ejecutor** (lista).
   El **Estado final del equipo** (lista) solo se solicita cuando el **Resultado es
   "Si"**; en cualquier otro caso queda en blanco. Pulse **Guardar registro**.

Las mantenciones registradas se acumulan en una tabla y:

- se **consolidan en el Excel masivo**: al pulsar *Descargar Excel (una fila por
  evento)*, cada registro **actualiza la fila** de su equipo y mes (resultado,
  fecha, ejecutor, estado final y observación); y
- también se pueden exportar por separado con **⬇️ Descargar registro (Excel)**
  → `Registro_Mantenciones_AAAAMMDD.xlsx` (mismo formato).

### Persistencia y comparación con el archivo

- Los registros se **guardan en el navegador** (localStorage): se conservan al
  recargar la página y **al volver a subir el archivo** (no se pierden).
- Al cargar el archivo, el sistema **compara** cada mantención registrada con el
  resultado del archivo para el mismo equipo y mes, y muestra un panel
  **🔍 Comparación: registrado vs. archivo** con las diferencias (por ejemplo,
  registraste `C2` pero el archivo trae `C3`, o está vacío, o el equipo/mes no
  aparece). Al descargar, **prevalece lo registrado** en el programa.
- El botón **🗑️ Limpiar** borra todos los registros guardados.

> Nota: la persistencia depende del navegador. Algunos navegadores no guardan
> datos al abrir el archivo con `file://`; en ese caso los registros se mantienen
> durante la sesión (incluso al re-subir el archivo) pero no tras cerrar la
> pestaña. Use **Descargar registro** para conservar una copia.

---

## Qué lee del archivo

- **Hoja `PMP_2026`** (programación): datos desde la fila 7, columnas **B a AE**,
  excluyendo solo **S** (Responsable MP); se **incluye Q** (Observación). Las
  columnas **T a AE** son los doce meses; cada mes lleva el código de programación (P).
- **Hoja `Registro_MP-2026`** (ejecución): mismos campos de equipo (también se
  ignoran **Q** y **S**). Cada mes tiene dos subcolumnas: **P** (programa) y
  **R** (resultado). Las columnas P son fórmulas que reflejan la hoja de
  programación; la herramienta lee su **resultado en caché**.

Cada equipo se identifica de forma única por su **N° de Serie** o **N° de
Inventario**. Los números de serie se conservan **tal cual** (se respetan los
ceros a la izquierda; por ejemplo `0024` no se transforma en `24`).

### Definición de "evento"

Se genera una fila por cada **equipo × mes** que tenga programación (P) y/o
resultado (R). Una mantención reprogramada produce, naturalmente, dos filas: el
mes de origen (con `X` + la causal) y el mes de destino (con `R`).

---

## Excel de salida

**Hoja `Eventos`** — una fila por evento, con encabezado fijo, autofiltro y
columnas:

| Datos del equipo | Evento |
|---|---|
| Familia, ID, N° Carpeta, N° Inventario, Equipo, Servicio, Unidad, Ubicación, Procedencia, Marca, Modelo, N° Serie, Año Instalación, Vida Útil Residual, Clasificación, ENU / Baja, Frecuencia MP, **Observación** | Mes, N° Mes, Programa (P), Tipo de Programación, Resultado (R), Detalle del Resultado, **Fecha de Ejecución**, Causal Reprog., Descripción de la Causal, Regla de Reprogramación, **Estado**, **Estado Final del Equipo**, **Ejecutor** |

- **Observación**: texto tal cual viene en la columna Q del equipo.
- **Resultado (R)**: **desplegable** con los códigos válidos (`Si`, `C1`–`C8`,
  `Si-RA`, `FS`, `No`, `NU`, `Baja`).
- **Fecha de Ejecución**: columna de **llenado manual** con formato de fecha
  (`dd-mm-aaaa`). El archivo de origen solo registra el *mes* (no el día), por lo
  que la fecha exacta se ingresa a mano cuando se ejecuta la mantención.
- **Estado** se colorea automáticamente (Realizada, Reprogramada, Pendiente,
  Fuera de Servicio, No Realizada, No Ubicable, Baja, Puesta en Marcha).
- **Estado Final del Equipo**: **desplegable** de llenado manual con las opciones
  *Operativo* / *No operativo* (se colorea verde/rojo según lo elegido).
- **Ejecutor**: **desplegable** con los 11 ejecutores disponibles
  (lista en la hoja `Catalogos`).

**Hoja `Catalogos`** — fuente del desplegable de Ejecutor y diccionario de
códigos (Programa, Resultado y Causales con su regla).

**Hoja `Resumen`** — totales de equipos y eventos, conteo por Estado y tabla
mensual (programadas / realizadas / reprogramadas / pendientes / otras).

---

## Diccionario de códigos

**Programación (P):** `X` Programada · `R` Reprogramada · `RA` Reprogramada de
Año Anterior · `PM` Puesta en Marcha.

**Resultado (R):** `Si` Realizada · `C1`–`C8` Reprogramada (causales) ·
`Si-RA` de Año Anterior Realizada · `FS` Fuera de Servicio · `No` No Realizada ·
`NU` No Ubicable · `Baja` Dado de Baja.

**Causales y reglas de reprogramación:**

| Código | Descripción | Regla |
|---|---|---|
| C1 | Imposibilidad de desocupar el equipo por indicación clínica | Reprogramar ≤ 30 días |
| C2 | Equipo en servicio técnico | Sin nueva fecha; al reintegrarse |
| C3 | Equipo no operativo (espera repuestos/accesorios) | Sin nueva fecha; al reintegrarse |
| C4 | Equipo en préstamo a otro hospital/institución | Sin nueva fecha; al reintegrarse |
| C5 | No disponibilidad HH funcionario SEC por carga laboral | Reprogramar ≤ 30 días |
| C6 | No disponibilidad HH servicio técnico externo | Reprogramar ≤ 30 días |
| C7 | Ausencia justificada funcionario SEC > 15 días | Reprogramar ≤ 30 días |
| C8 | Contingencia hospitalaria | Reprogramar ≤ 30 días |

Para C1, C5, C6, C7 y C8 el mes de origen mantiene `X` en P y la causal en R, y
el mes de destino lleva `R` en P.

---

## Estructura del proyecto (para mantenimiento)

```
index.html            Aplicación lista para usar (autocontenida).
src/template.html     Plantilla HTML + estilos + interfaz.
src/core.js           Lectura y transformación (workbook → eventos).
src/output.js         Construcción del Excel de salida (ExcelJS).
src/app.js            Lógica de la interfaz (carga, resumen, descarga).
vendor/exceljs.min.js Librería ExcelJS (incrustada para uso offline).
build.js              Ensambla index.html a partir de lo anterior.
```

Para regenerar `index.html` tras editar `src/` o actualizar la librería:

```bash
node build.js
```

`build.js` solo usa módulos nativos de Node (no requiere `npm install`).
