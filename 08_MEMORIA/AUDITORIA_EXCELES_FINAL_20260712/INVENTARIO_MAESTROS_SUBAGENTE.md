# Inventario y selección de libros maestros

Fecha de auditoría: 2026-07-12 (Europe/Madrid).

## Alcance

Se inventariaron recursivamente y en modo de solo lectura estas rutas:

- `08_MEMORIA/ENTREGABLE`.
- `08_MEMORIA/Excel_visibilidad_FP`.
- `08_MEMORIA/Excel_benchmark`.

No se recorrió el contenido de directorios prohibidos (`candidate`, `debug`, `OLD_CONTEXT_USELESS`, `OLD_NO_LEER_POR_DEFECTO`, `99_ARCHIVE` ni `_ARCHIVO_OBSOLETO_NO_USAR`). No se modificó, guardó, movió ni recalculó ningún Excel durante esta subauditoría.

La inspección combinó:

- SHA-256, tamaño y fecha de modificación.
- Lectura completa de cada entrada ZIP para detectar fallos de descompresión.
- Análisis directo de OpenXML: libro, hojas, relaciones, shared strings, fórmulas y valores almacenados, tablas, gráficos y series, nombres definidos, filtros, validaciones, celdas combinadas, freeze panes, áreas de impresión, estilos, propiedades de cálculo, conexiones y enlaces externos.
- Comparación normalizada celda a celda entre versiones, separando contenido/fórmula de cambios exclusivamente visuales.

`openpyxl` no estaba instalado en el intérprete Python disponible. Para no instalar dependencias ni arriesgar una reescritura, la inspección se realizó directamente sobre ZIP/OpenXML; esto permitió leer simultáneamente fórmula y valor cacheado sin guardar los paquetes.

## Inventario

El primer inventario, realizado antes de la creación del backup por el proceso principal, contenía cinco libros. La instantánea profunda posterior contiene siete rutas porque el backup añadió dos copias binariamente idénticas; no son candidatos nuevos.

| Ruta relativa | Bytes | Modificación | SHA-256 |
|---|---:|---|---|
| `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx` | 280.969 | 2026-06-26 13:59:54 +02:00 | `947161E4ECF609EF8CD97AD24BB604EFEDE04FD1251642ECFEAC1BF5AAEEFE09` |
| `08_MEMORIA/ENTREGABLE/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.219 | 2026-06-26 14:37:36 +02:00 | `AD953B6F760DBFF580099B8DA899929CAED3383419C97E8E69C3E6D6F9E52CF4` |
| `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_V.5.xlsx` | 282.039 | 2026-06-26 11:40:12 +02:00 | `F9CD06E7E986102028C7FE28523BEAD774ED8E724FD7464382CF2C91A45293C1` |
| `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx` | 282.145 | 2026-06-26 12:10:23 +02:00 | `C46ED2C41F559053BD1C290F74FCEF82BD4509832921294E4EB4C02DAE28B0FB` |
| `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.087 | 2026-06-26 13:01:26 +02:00 | `AEB8835BAD0B2A8D015D399DDA11A2CD475B3F320750ABD333750026456A36EC` |
| `08_MEMORIA/ENTREGABLE/BACKUP_ANTES_CIERRE_EXCELES_20260712/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx` | 280.969 | 2026-06-26 13:59:54 +02:00 | `947161E4ECF609EF8CD97AD24BB604EFEDE04FD1251642ECFEAC1BF5AAEEFE09` |
| `08_MEMORIA/ENTREGABLE/BACKUP_ANTES_CIERRE_EXCELES_20260712/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.219 | 2026-06-26 14:37:36 +02:00 | `AD953B6F760DBFF580099B8DA899929CAED3383419C97E8E69C3E6D6F9E52CF4` |

Las dos copias de backup tienen exactamente el mismo tamaño y SHA-256 que los originales de `ENTREGABLE`.

## Comparación estructural de candidatos únicos

| Libro | Hojas | Registros de celda | Fórmulas | Tablas | Gráficos | Nombres definidos | Errores almacenados | Series rotas |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| Visibilidad — `ENTREGABLE` | 31 | 9.932 | 9 | 10 | 6 | 11 | 0 | 0 |
| Visibilidad — auxiliar V.5 | 31 | 9.971 | 9 | 10 | 6 | 11 | 0 | 0 |
| Visibilidad — auxiliar FINAL | 31 | 9.971 | 9 | 10 | 6 | 11 | 0 | 0 |
| Benchmark — `ENTREGABLE` | 11 | 34.912 | 0 | 0 | 6 | 10 | 0 | 0 |
| Benchmark — auxiliar | 11 | 34.912 | 0 | 0 | 6 | 10 | 0 | 0 |

En todos los libros:

- el ZIP se leyó completo sin error;
- están presentes `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml` y las relaciones del libro;
- no se localizaron `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!`, errores Excel almacenados ni nombres definidos rotos;
- no hay gráficos sin series ni series con referencias `#REF!`;
- no hay enlaces externos, conexiones, macros ni tablas dinámicas.

## Maestro recomendado de visibilidad, detección, FP y Wazuh

**Ruta recomendada:**

`C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`

**SHA-256:** `947161E4ECF609EF8CD97AD24BB604EFEDE04FD1251642ECFEAC1BF5AAEEFE09`.

### Razones

1. Es la versión prioritaria de `ENTREGABLE` y no una copia auxiliar.
2. Conserva 31 hojas, 10 tablas, 9 fórmulas, 6 gráficos y 11 nombres definidos.
3. Integra contenido real de:
   - TEC-001 a TEC-009;
   - P1, P2, P3, P4, `CLIENT_EVENT` y Router;
   - FP-001 a FP-010;
   - Wazuh, incluidas `COMPARACION_VR_WAZUH` y `WAZUH_DETALLE`;
   - artifacts públicos, incluidos `ProcessCreation`, `ServiceCreation`, `SysmonLogForward`, `Hayabusa.Monitoring`, `ETW.Monitoring` y `TrackNetworkConnections`;
   - dashboard, resumen ejecutivo, discrepancias y fuentes.
4. La comparación normalizada contra la auxiliar FINAL detecta 127 diferencias de contenido o fórmula, concentradas en:
   - `00_Guia`: 2 celdas;
   - `01_Dashboard`: 45 celdas;
   - `03_Evaluacion_VR`: 80 celdas.
5. Esas diferencias no son pérdida accidental: el libro de `ENTREGABLE` sustituye estados antiguos como `Pendiente`, `Pendiente benchmark` y `Cerrar comparación pública antes de custom` por resultados consolidados P1-P4/`Client_Event` y conclusiones técnicas por TEC. También elimina del dashboard una lista de acciones ya superada.
6. La tabla `EvaluacionVRTable` queda consolidada en `A4:O13`; las auxiliares mantienen la estructura anterior `A4:S13`, con columnas de planificación ya obsoletas.
7. El menor número de celdas frente a las auxiliares responde a esa consolidación, no a ausencia de Wazuh, FP, custom o evidencia TEC.

Las dos variantes auxiliares de visibilidad tienen exactamente el mismo contenido y las mismas fórmulas entre sí: cero diferencias semánticas. Sus 7.872 diferencias son exclusivamente de estilo. Por tanto, la auxiliar denominada FINAL es una revisión visual de V.5, pero ambas conservan el estado semántico anterior al libro de `ENTREGABLE`.

### Hojas del maestro

`00_Guia`, `01_Dashboard`, `02_Seleccion_Tecnicas`, `03_Evaluacion_VR`, `04_Control_Publicos`, `05_Matriz_Resultados`, `06_Resumen`, `07_Catalogo_Artifacts`, `08_Benignas_FP`, `09_Benchmark_Plan`, `10_Gaps_Custom_P1P4`, `99_Listas`, `11_Alertabilidad`, `12_Plan_Alertas`, `13_Hayabusa_Resumen`, `14_Arquitectura_Custom`, `15_Publicos_Definitivo`, `16_Scripts_Validados`, `17_Incoherencias_Cerradas`, `README`, `RESUMEN_EJECUTIVO`, `TECNICAS_REAL_VR`, `VISIBILIDAD_SISTEMA`, `ALERTAS_VR`, `FP_RUNNER`, `FP_HITS`, `COMPARACION_VR_WAZUH`, `WAZUH_DETALLE`, `DISCREPANCIAS`, `FUENTES` y `GRAFICAS`.

### Gráficos detectados

- `01_Dashboard`: 1 gráfico de barras, 1 serie válida, referencia real a `H5:H7` y `J5:J7`; el XML no contiene título propio y debe revisarse en la edición final.
- `GRAFICAS`: `Eventos Sysmon/4104 por TEC` — 3 series válidas.
- `GRAFICAS`: `Alertas VR por perfil` — 1 serie válida.
- `GRAFICAS`: `FP OK/SKIPPED/HITS` — 1 serie válida.
- `GRAFICAS`: `Wazuh base vs custom` — 2 series válidas.
- `GRAFICAS`: `Reglas Wazuh 110xxx` — 2 series válidas.

No hay gráficos vacíos ni referencias rotas. El libro mantiene `calcChain.xml`; `forceFullCalc` no está activado en el paquete inicial, por lo que procede el recálculo posterior mediante Excel COM solicitado por el usuario.

## Maestro recomendado de benchmark

**Ruta recomendada:**

`C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`

**SHA-256:** `AD953B6F760DBFF580099B8DA899929CAED3383419C97E8E69C3E6D6F9E52CF4`.

### Razones

1. Es la versión prioritaria de `ENTREGABLE`.
2. La comparación contra la auxiliar arroja cero diferencias de contenido o fórmula. Las 170 diferencias detectadas son exclusivamente de estilo en `TRAZABILIDAD`, además de metadatos y posicionamiento de dibujo.
3. `RUNS_VALIDOS` ocupa `A1:V10`: una cabecera y exactamente nueve ejecuciones.
4. Las nueve filas están marcadas `VALID` y `ValidForClientBenchmark=1`:
   - `BASELINE_NO_VR`: 3 repeticiones;
   - `VR_IDLE`: 3 repeticiones;
   - `VR_TEC_RUNNER`: 3 repeticiones.
5. `SERVER_GUI` figura como exclusión metodológica, no como cuarto escenario del resumen válido.
6. Conserva 34.912 registros de celda, 11 hojas, 6 gráficos con referencias reales y 10 nombres definidos.
7. El paquete inicial ya contiene `forceFullCalc=1` y no tiene `calcChain.xml`, coherente con que no existen fórmulas almacenadas.

### Hojas del maestro

`README`, `RESUMEN_EJECUTIVO`, `RUNS_VALIDOS`, `RESUMEN_ESCENARIO`, `RAW_SAMPLES`, `PROCESS_SAMPLES`, `CALIDAD_DATOS`, `GRAFICAS`, `TRAZABILIDAD`, `LIMITACIONES` y `DISCREPANCIAS`.

### Gráficos detectados

- `CPU media por escenario` — 1 serie válida.
- `CPU maxima por escenario` — 1 serie válida.
- `RAM media por escenario` — 1 serie válida.
- `RAM maxima por escenario` — 1 serie válida.
- `Duracion media por escenario` — 1 serie válida.
- `Comparacion BASELINE vs VR_IDLE vs VR_TEC_RUNNER` — 3 series válidas.

Todos tienen título, series y rangos OpenXML válidos; no se detectan series vacías ni referencias rotas.

### Observación técnica

El benchmark no contiene fórmulas Excel: los resúmenes y métricas están materializados como valores. Esto no implica una fórmula rota, pero sí significa que `CalculateFullRebuild` no puede regenerar esos agregados desde `RAW_SAMPLES` por sí solo. La trazabilidad se apoya en las hojas de detalle, rutas fuente y SHA-256. Si el cierre exige agregados recalculables dentro de Excel, deberán introducirse fórmulas o tablas estructuradas de forma controlada sin alterar las nueve mediciones.

## Dictamen de selección

- Maestro de visibilidad/detección/FP/Wazuh: **RECOMENDADO** — libro de `ENTREGABLE`, SHA-256 `947161E4...FE09`.
- Maestro de benchmark: **RECOMENDADO** — libro de `ENTREGABLE`, SHA-256 `AD953B6F...52CF4`.
- Estado de esta subauditoría: selección estructural completada; no sustituye la validación Excel COM ni la reapertura final de los entregables definitivos.

