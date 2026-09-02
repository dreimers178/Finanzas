# Reportes mensuales

Dos reportes que se generan **cada mes** a partir de un export nuevo. Los scripts
son reutilizables: se les pasa el archivo del mes y producen el `.xlsx` con las
mismas tablas y formato de siempre. Cada archivo se abre directo en Excel/Mac
(recálculo forzado al abrir).

---

## 1. Reporte de Portal Uploads

**Script:** `generar_portal_uploads.py`
**Entrada:** export de Google Sheets/Excel con la hoja `Form Responses`
(columnas: `Invoice Number`, `Portal Name or Link`, `Submitted By`,
`Collector assigned`, `Timestamp`, `Message Timestamp`, …) o un CSV equivalente.

**Métrica base:** `Resolution Time = Message Timestamp − Timestamp`.

**Uso:**
```bash
python3 reportes/generar_portal_uploads.py <archivo_del_mes.xlsx> [AAAA-MM] [salida.xlsx]
# Ej: python3 reportes/generar_portal_uploads.py export_sep.xlsx 2026-09 Reporte_Sep_2026.xlsx
```
Si no se pasa el mes, toma el último mes con datos.

**Contenido del reporte:**
- **Resolution Time por Mes** (todos los meses): `# Uploads`, `Resueltos`,
  `Prom. horas`, `Prom. días`, `% mismo día`.
- **Distribución por tiempo de resolución** (mes destacado): <24h, 1–2d, 2–4d, 4–7d, >7d.
- **Resolution Time por Portal** (mes destacado, nombres normalizados
  Ariba/ARIBA/ariba → Ariba): `# Uploads`, `%`, `Prom. días`, `% mismo día`.
- **Atípicos >7 días** (mes destacado) en hoja aparte.
- Hojas de respaldo: `Form Responses`, `Nomenclature`, `Datos (calculo)`.

---

## 2. Reporte de Tesorio Tasks

**Script:** `generar_tesorio_tasks.py`
**Entrada:** CSV export de Tesorio (columnas: `Invoice Number`, `Assignees`,
`Title`, `Created At`, `Assigned by`, `Priority`, `Status`, `Completed At`).

**Métrica base:** `Resolution Time = Completed At − Created At` (solo tasks `DONE`).

**Estados:**
- `DONE` = completadas.
- `TO_DO` = pendientes al día de hoy.
- `WORKING` = en proceso.
Para `TO_DO`/`WORKING` se calculan **días abiertos = hoy − Created At**.

**Uso:**
```bash
python3 reportes/generar_tesorio_tasks.py <export.csv> [salida.xlsx] [hoy=AAAA-MM-DD]
# Ej: python3 reportes/generar_tesorio_tasks.py tesorio_sep.csv Reporte_Tesorio_Sep.xlsx
```
> Nota: la fecha "hoy" está fijada al día del export para reproducibilidad;
> al correr un mes nuevo, ajústala (o quita la línea que la fija en el script).

**Contenido del reporte:**
- **Estado de las tasks** (KPIs): total, completadas, pendientes, en proceso,
  prom. días, mediana, % mismo día.
- **Resolution Time por Mes** (por mes de creación): `# Creadas`, `Completadas`,
  `Prom. horas`, `Prom. días`, `Mediana días`, `% mismo día`.
- **Distribución por tiempo de resolución** (completadas).
- **Resolution Time por Prioridad** (URGENT / HIGH_PRIORITY / NORMAL).
- **Pendientes y En Proceso** en hoja aparte, con días abiertos (semáforo).
- **Tasks lentas >7 días** en hoja aparte.
- Hoja de respaldo: `Datos (calculo)`.

---

## Convenciones comunes
- Fuente Arial, encabezados azul marino, tablas con bordes finos.
- Números derivados guardados como valores; los resúmenes usan fórmulas vivas
  (`COUNTIF`/`AVERAGEIF`/`COUNTIFS`) sobre la hoja `Datos (calculo)`.
- `fullCalcOnLoad` activado para que Excel recalcule al abrir.
- Portales normalizados y medianas/`Otros` calculados sobre los datos actuales
  (anotado al pie de cada tabla).
