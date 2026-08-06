# Auditoría de candidatos del benchmark

Fecha: 2026-07-12 17:38 CEST.

Alcance: auditoría comparativa de los candidatos autorizados del benchmark. No se modificó, guardó ni copió ningún libro Excel. No se leyeron rutas históricas prohibidas ni carpetas `candidate`.

## Recomendación de maestro

Maestro recomendado:

`C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`

- Tamaño: 253.219 bytes.
- SHA-256: `AD953B6F760DBFF580099B8DA899929CAED3383419C97E8E69C3E6D6F9E52CF4`.
- Fecha: 2026-06-26 14:37:36 CEST.
- Hojas: 11.
- Gráficos: 6.
- Nombres definidos: 10, ninguno roto.

Motivo: es la versión prioritaria de `ENTREGABLE`, conserva el consolidado final y es más completa que el regenerado canónico en estructura de presentación, trazabilidad y control. Incluye `README`, `LIMITACIONES` y `DISCREPANCIAS`, además de rangos `RAW_SAMPLES` y `PROCESS_SAMPLES` enriquecidos.

La copia de `08_MEMORIA\Excel_benchmark` contiene exactamente los mismos valores y fórmulas. La comparación COM de sus 11 rangos usados produjo 0 diferencias de valor y 0 diferencias de fórmula. Su distinto hash se debe a metadatos, estilos, dibujo y XML de presentación/hoja, no a mediciones distintas.

## Inventario comparado

| Ubicación | Bytes | SHA-256 | Resultado |
|---|---:|---|---|
| `08_MEMORIA\ENTREGABLE\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.219 | `AD953B6F...E52CF4` | Maestro recomendado |
| `08_MEMORIA\ENTREGABLE\BACKUP_ANTES_CIERRE_EXCELES_20260712\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.219 | `AD953B6F...E52CF4` | Duplicado binario del maestro |
| `08_MEMORIA\Excel_benchmark\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.087 | `AEB8835B...A36EC` | Mismos valores/fórmulas; versión anterior de presentación |
| `04_EVIDENCE\Excel_benchmark_26062026\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.087 | `AEB8835B...A36EC` | Duplicado binario de la copia auxiliar |
| `04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_REGENERADOS\02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx` | 152.067 | `88551556...52B8C` | Fuente canónica de datos; 9 hojas y 4 gráficos |

El regenerado canónico conserva `RUNS_9`, `METODOLOGIA` y la fuente de datos original, pero tiene 9 hojas, 4 gráficos, 0 nombres definidos y rangos RAW/PROCESS menos ricos. Debe mantenerse como fuente trazable, no como maestro de presentación final.

## Estructura del maestro

| Hoja | Rango usado | Filtro | Inmovilización | Fórmulas | Gráficos |
|---|---|---|---|---:|---:|
| README | A1:B11 | A1:B11 | A2 | 0 | 0 |
| RESUMEN_EJECUTIVO | A1:C18 | A1:C18 | A2 | 0 | 0 |
| RUNS_VALIDOS | A1:V10 | A1:V10 | A2 | 0 | 0 |
| RESUMEN_ESCENARIO | A1:L4 | A1:L4 | A2 | 0 | 0 |
| RAW_SAMPLES | A1:AL423 | A1:AL423 | A2 | 0 | 0 |
| PROCESS_SAMPLES | A1:W771 | A1:W771 | A2 | 0 | 0 |
| CALIDAD_DATOS | A1:E20 | A1:E20 | A2 | 0 | 0 |
| GRAFICAS | A1:K54 | No aplica | A2 | 0 | 6 |
| TRAZABILIDAD | A1:H82 | A1:H82 | A2 | 0 | 0 |
| LIMITACIONES | A1:B8 | A1:B8 | A2 | 0 | 0 |
| DISCREPANCIAS | A1:E4 | A1:E4 | A2 | 0 | 0 |

No existen tablas estructuradas `ListObject`; los rangos de datos utilizan `AutoFilter`. Tampoco existen fórmulas: el consolidado contiene valores materializados. No se detectaron enlaces externos, conexiones, tablas dinámicas, VBA ni nombres definidos rotos.

Esta ausencia de fórmulas significa que `CalculateFullRebuild` no alterará los datos derivados; la validación final COM sigue siendo necesaria para fijar propiedades de cálculo, actualizar objetos y comprobar la reapertura.

## Validación de las nueve ejecuciones

- 9 filas de ejecución y 9 válidas para benchmark de cliente.
- Escenarios exactos: `BASELINE_NO_VR`, `VR_IDLE`, `VR_TEC_RUNNER`.
- Tres repeticiones exactas por escenario: 1, 2 y 3.
- Sin `RunId` duplicados.
- Sin pares escenario/repetición duplicados.
- Sin escenarios ajenos.
- 422 filas de `RAW_SAMPLES`, asociadas a 9 `RunId`.
- 770 filas de `PROCESS_SAMPLES`.
- `notepad.exe`: 0 filas.
- `RunnerStillRunningAtEnd=True`: presente en las tres ejecuciones `VR_TEC_RUNNER`; se conserva correctamente como WARN no bloqueante.

Valores consolidados verificados:

| Métrica | Valor |
|---|---:|
| CPU media `VR_TEC_RUNNER` | 0,95 % |
| RAM media `VR_IDLE` | 57,48 MB |
| RAM media `VR_TEC_RUNNER` | 57,42 MB |
| Pico CPU cliente | 54 % |
| Pico RAM cliente | 57,97 MB |
| Duración media `BASELINE_NO_VR` | 182,69 s |
| Duración media `VR_IDLE` | 181,57 s |
| Duración media `VR_TEC_RUNNER` | 315,60 s |

## Exclusión de SERVER_GUI

`SERVER_GUI` está observado, pero excluido correctamente de las métricas principales:

- 422 filas RAW tienen proceso de servidor GUI observado.
- `PROCESS_SAMPLES` contiene 422 filas de rol `SERVER_GUI`, 284 de `CLIENT_SERVICE` y 64 de `OTHER_VELOCIRAPTOR`.
- `RUNS_VALIDOS` y `RESUMEN_ESCENARIO` usan métricas `VR_Client_*`, no `VR_Total_IncludingServer_*`.
- `README` documenta `IncludeServerGuiInTotal=False` y `ServerGuiExcludedFromClientMetrics=True`.
- `LIMITACIONES` y `DISCREPANCIAS` documentan la exclusión.
- Las columnas `VR_Total_IncludingServer_*` se conservan únicamente como trazabilidad secundaria en `RAW_SAMPLES`.

Resultado: `SERVER_GUI` sigue excluido del benchmark principal sin ocultar su observación en datos crudos.

## Auditoría de gráficos

Las seis definiciones de serie son no vacías y no contienen `#REF!`.

| Gráfico | Fuente | Series | Resultado |
|---|---|---:|---|
| CPU media por escenario | `GRAFICAS!A1:B4` | 1 | Rango correcto; falta unidad visible y leyenda |
| CPU máxima por escenario | `GRAFICAS!D1:E4` | 1 | Rango correcto; falta unidad visible y leyenda |
| RAM media por escenario | `GRAFICAS!G1:H4` | 1 | Rango correcto; falta unidad visible y leyenda |
| RAM máxima por escenario | `GRAFICAS!J1:K4` | 1 | Rango correcto; falta unidad visible y leyenda |
| Duración media por escenario | `GRAFICAS!A18:B21` | 1 | Rango correcto; falta unidad visible y leyenda |
| Comparación de escenarios | `GRAFICAS!D18:G21` | 3 | No apto: mezcla %, MB y segundos en el mismo eje |

Hallazgos:

1. Todos los gráficos tienen `HasLegend=False` y 0 títulos de eje. Las unidades solo existen en los nombres de serie y por ello no se ven.
2. El gráfico 6 usa como categorías `AvgCPUPercent`, `AvgRAMMB` y `AvgDurationSeconds` en un único eje. La comparación no es cuantitativamente válida.
3. Ningún gráfico compara las repeticiones 1, 2 y 3. Falta uno de los gráficos mínimos requeridos.

## Errores de contenido visibles

### BENCH-001 — RESUMEN_EJECUTIVO!C3

- Observado por Excel COM: `Value2=46274`, texto mostrado `09-sep`, formato `dd-mmm`.
- Contenido pretendido: texto `9/9`.
- Impacto: la observación de runs válidos se presenta como fecha.

### BENCH-002 — RESUMEN_EJECUTIVO!B5

- Observado por Excel COM: `Value2=37683`, texto mostrado `03/03/2003`, formato `dd/mm/aaaa`.
- Contenido pretendido: texto `3/3`.
- Impacto: el KPI de repeticiones por escenario se presenta como fecha.

Nota de precisión: `RESUMEN_EJECUTIVO!B3` contiene el número 9 con formato General y es correcto. La celda de fecha errónea de esa fila es `C3`.

## Reparaciones mínimas necesarias

1. En `RESUMEN_EJECUTIVO`, aplicar primero formato Texto y escribir `C3='9/9'` y `B5='3/3'`.
2. Conservar los gráficos 1 a 5, activar sus leyendas y hacer visibles las unidades `%`, `MB` y `s` mediante títulos de eje o títulos explícitos.
3. Sustituir el gráfico 6 por una comparación homogénea de las tres repeticiones. Opción mínima y trazable: duración en segundos por repetición y escenario, con categorías `RUNS_VALIDOS!C2:C4` y series `RUNS_VALIDOS!F2:F4`, `F5:F7` y `F8:F10`.
4. No modificar CPU, RAM, duraciones, escenarios, repeticiones ni muestras.
5. No incorporar campañas de artifacts públicos al benchmark.
6. Tras la edición: cálculo automático, `CalculateFullRebuild`, guardado, reapertura, comprobación de 6 gráficos no vacíos, 9/9 runs, exclusión de `SERVER_GUI` y ausencia de avisos de reparación.

## Validación técnica realizada

OpenXML:

- 50 entradas ZIP leídas completamente.
- Partes esenciales de libro y estilos presentes.
- 0 entradas ilegibles.
- 0 logs de reparación internos.
- 0 fórmulas rotas.
- 0 series rotas o vacías.
- 0 nombres definidos rotos.
- 0 enlaces externos.

Limitación: esta comprobación valida legibilidad e integridad práctica del paquete y sus relaciones relevantes, no un esquema XSD completo.

Excel COM:

- Apertura correcta en modo de solo lectura.
- 11 hojas accesibles.
- 0 celdas de fórmula con error.
- 0 constantes de error.
- 6 gráficos accesibles y con series.
- 0 enlaces externos y 0 conexiones.
- Sin guardado ni recálculo en esta fase, conforme al alcance de solo auditoría.

## Estado previo a edición

`APTO CON OBSERVACIONES` como maestro de partida.

No es apto todavía como libro definitivo de cierre hasta corregir `C3`, `B5`, las leyendas/unidades y el gráfico 6, y repetir la validación OpenXML/Excel COM sobre el archivo final.

El detalle estructurado se conserva en `AUDITORIA_BENCHMARK_SUBAGENTE.json`.
