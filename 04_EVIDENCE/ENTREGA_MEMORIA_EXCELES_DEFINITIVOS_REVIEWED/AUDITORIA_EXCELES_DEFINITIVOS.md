# Auditoría técnica de Excel definitivos

Fecha: 2026-06-18  
Alcance: `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS`  
Estado global: **NO APTO**

## Resumen ejecutivo

Ningún Excel está perfecto para entrega en su estado actual.

- El libro 01 contiene cobertura y FP útiles, pero mezcla agregados SOC de campañas no identificadas, usa tres hits contaminados en sus KPI, carece de trazabilidad fila a fila y tiene un gráfico con rango incorrecto que omite TEC-009.
- El libro 02 es coherente como control de calidad, pero confirma que falta el benchmark definitivo del cliente Velociraptor. Conforme a las reglas de esta auditoría, esto es BLOCKER.
- El libro 03 es histórico, contiene dos tablas con rangos incompletos y emplea lenguaje de “definitivo” que contradice el estado vigente.
- El libro 04 es una plantilla metodológica con todas las pruebas pendientes; no es un resultado definitivo.
- La transferencia TEC-009 sí ocurrió y existe ZIP recibido con hash coincidente. El fallo está en la trazabilidad del Excel, no en la transferencia.

No se ha modificado ningún Excel ni se han creado copias corregidas: la mayoría de los cambios exige seleccionar una campaña oficial o repetir pruebas. No se genera `CHANGELOG_CORRECCIONES.md`.

## Archivos revisados

| Archivo | Hojas/datos | Estado de auditoría | Propósito |
|---|---:|---|---|
| `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx` | 8 hojas, 6 gráficos | NO APTO | Detección, FP y fiabilidad |
| `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx` | 8 hojas, 4 gráficos | NO APTO como benchmark; válido como control de calidad | Benchmark legacy invalidado |
| `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx` | 19 hojas, 1 gráfico | NO APTO en paquete final sin rotulado/corrección | Referencia histórica |
| `04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx` | 9 hojas, 2 gráficos | APTO solo como plantilla/anexo, no como resultado | Plan metodológico público |
| `fuentes_analisis/*.xlsx` | 3 libros duplicados | Redundantes | Copias originales |
| `fuentes_benchmark/*.csv` | 3 CSV | Coherentes con libro 02 | Fuentes normalizadas |
| `fuentes_fp/*.csv` | 3 CSV | Incluyen contaminación documentada | Fuentes FP |
| `MANIFEST_ENTREGA_MEMORIA.csv` | 30 filas | Desactualizado | Inventario del paquete |

No hay `.xls` ni `.xlsm`.

## Hallazgos por severidad

### BLOCKER (7)

1. **02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx** - RESUMEN_EJECUTIVO / SCENARIO_SUMMARY / CALIDAD_DATOS - RESUMEN_EJECUTIVO!A2:D9; SCENARIO_SUMMARY!A2:R4
   - Problema: No existe benchmark definitivo del agente cliente. El libro solo documenta datos legacy invalidados: CLIENT_SERVICE=0, ServiceStatus=Stopped en VR_IDLE/VR_TEC_RUNNER y SchemaVersion=1.0.
   - Evidencia: 00_CONTEXT/DECISIONS.md D035; 05_LOGS/BENCHMARKS/validado/scenario_quality.csv; el propio libro declara NO_VALIDO_PARA_CONCLUSIONES_DEFINITIVAS_DE_RENDIMIENTO.
   - Corrección: Repetir BASELINE_NO_VR, VR_IDLE y VR_TEC_RUNNER con el script corregido; exigir CLIENT_SERVICE en idle/runner y regenerar el Excel desde los nuevos logs.
   - Automático seguro: false
2. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - TEC_ATTACK_VALIDATION - E4:G12
   - Problema: Los Hits SOC y el perfil máximo mezclan o no identifican campañas. No coinciden con el JSONL pre-FP disponible: Excel=27/106/31/18/26/17/27/22/26; JSONL=21/106/15/18/11/17/4/2/6. TEC-008 figura P1_CRITICAL en Excel, pero en ese JSONL su máximo es P2_EVENT.
   - Evidencia: Carpeta_Compartida_TFM/inbox/soc_alerts_PRE_FP_20260618_144023.jsonl, SHA256 935CC7DA04A0B8BB0122BD96FBCDCD2258BF0BF3618DF032E045F944689D72CD, 200 filas.
   - Corrección: Elegir un único RunId/ventana ofensiva como fuente oficial, documentarla y recalcular todas las filas por TEC y perfil directamente desde ese JSONL.
   - Automático seguro: false
3. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - TEC_ATTACK_VALIDATION / TRAZABILIDAD - G4:G12; A4:E14
   - Problema: La columna Detección confirmada=Sí no tiene trazabilidad fila a fila hasta alertas CLIENT_EVENT. Combina OK del runner (ejecución) con recuentos SOC de contexto, aunque runner OK no equivale a detección.
   - Evidencia: 05_LOGS/FINAL_RUN_20260616_175138/RESUMEN_EVIDENCIA_FINAL.md marca TEC-001..TEC-008 FAIL en evidencia Velociraptor esperada; 07_DOCS/MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT.md termina con Validación pendiente.
   - Corrección: Añadir CampaignId, Artifact, Source, DetectionTime, ruta/hash JSONL y referencia de evento por cada TEC. Separar ejecución, visibilidad, detección, alerta RT y salida externa.
   - Automático seguro: false
4. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - GRAFICAS - Chart 1; categorías A5:A13 y valores B5:B13
   - Problema: El gráfico Alertas ofensivas por TEC incluye la fila A5 de rótulo y omite TEC-009 en la fila 14.
   - Evidencia: Serie Excel: =SERIES("Alertas",GRAFICAS!$A$5:$A$13,GRAFICAS!$B$5:$B$13,1); datos reales TEC en A6:B14.
   - Corrección: Cambiar la serie a categorías A6:A14 y valores B6:B14 y volver a revisar visualmente.
   - Automático seguro: true
5. **03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx** - 04_Control_Publicos - ControlPublicosTable A4:X61 frente a datos A4:AE61
   - Problema: La tabla formal termina en X y deja fuera Y:AE (Visibilidad, Detección_Regla, Alerta_RT, Artifact_Alertante, Salida_Externa, Veredicto_Alertabilidad y Siguiente_Paso). Filtros y estructura no cubren las columnas añadidas.
   - Evidencia: Metadatos COM: tabla ControlPublicosTable=A4:X61; celdas con datos hasta AE61.
   - Corrección: Ampliar la tabla formal a A4:AE61 y verificar filtros/estilo.
   - Automático seguro: true
6. **03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx** - 09_Benchmark_Plan - BenchmarkPlanTable A4:J11 frente a datos A4:J13
   - Problema: La tabla formal excluye las filas B7/P2 High y B8/P1 Critical (filas 12 y 13).
   - Evidencia: Metadatos COM: BenchmarkPlanTable=A4:J11; celdas A12:J13 contienen escenarios B7 y B8.
   - Corrección: Ampliar la tabla a A4:J13 y revisar cualquier filtro o total dependiente.
   - Automático seguro: true
7. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - TEC_ATTACK_VALIDATION / TRAZABILIDAD - Fila TEC-009 (12); TRAZABILIDAD A4:E14
   - Problema: La transferencia TEC-009 es real, pero el Excel no enlaza la evidencia receptora original, por lo que su afirmación no es auditable desde el libro.
   - Evidencia: receiver_log.jsonl contiene POST /upload de 6245 bytes y SHA256 F7E58B...; upload_20260616_154943_944765.zip existe y su hash coincide; FINAL_RUN incluye summary, ZIP y 4104.
   - Corrección: Añadir al libro rutas y SHA256 de receiver_log.jsonl, ZIP recibido, tec009_transfer_summary.json, 4104 y campaña exacta.
   - Automático seguro: false

### MAJOR (11)

1. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - RESUMEN_EJECUTIVO / METRICAS / GRAFICAS - B11:B15; METRICAS B12:B15; GRAFICAS D5:E8, G5:H6, J5:K7
   - Problema: Los 15 hits FP v1.1 incluyen 3 eventos del RunId anterior TFM_FP_20260618_151244. Los agregados principales y gráficos usan P4=10 y total=15, en vez del conjunto limpio de 12 (P4=7, P3=5).
   - Evidencia: FP_HITS_RAW: EXACT_RUNID=3, TIME_WINDOW_INFERRED=9 y PREVIOUS_RUN_CONTAMINATION=3; filas 8,12,13 son contaminadas.
   - Corrección: Separar métricas BRUTAS (15) y LIMPIAS (12). Usar las limpias en indicadores principales y gráficos; mantener las 3 contaminadas solo en auditoría.
   - Automático seguro: false
2. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - FP_TEST_MATRIX - F4:F13
   - Problema: La asignación de hits por prueba suma 8, mientras el conjunto limpio tiene 12 hits. Cuatro P4 quedan sin asignar a FP_ID, pero la hoja no muestra una categoría No asignado ni el criterio temporal.
   - Evidencia: FP_TEST_MATRIX: FP-002=2, FP-003=1, FP-006=5; FP_HITS_RAW limpio=12.
   - Corrección: Añadir fila/categoría UNASSIGNED o un campo de calidad por prueba; no imputar hits inferidos a FP_ID sin regla reproducible.
   - Automático seguro: false
3. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - RESUMEN_EJECUTIVO / METRICAS - B7,B16:B19; METRICAS B4:B12
   - Problema: Las métricas derivadas están hardcodeadas aunque la fuente/criterio dice FORMULA. El libro tiene 0 celdas de fórmula.
   - Evidencia: Inspección Excel COM: FormulaCells=0 en las ocho hojas; valores 1, 0, 0,333... y 1,666... son constantes.
   - Corrección: Sustituir valores derivados por fórmulas enlazadas a las hojas de detalle y añadir controles de reconciliación.
   - Automático seguro: true
4. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - RESUMEN_EJECUTIVO / METRICAS / GRAFICAS - B7,B16:B19; METRICAS B4:B11; GRAFICAS B21:B24
   - Problema: Los porcentajes usan formato Estándar: se muestran como 1 o 0,333333333333, no como 100 % o 33,33 %.
   - Evidencia: NumberFormatLocal=Estándar en las celdas revisadas.
   - Corrección: Aplicar formato 0,00 % y mantener los valores base entre 0 y 1; ajustar también el eje del gráfico Cobertura vs tasas FP.
   - Automático seguro: true
5. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - METRICAS / CONCLUSIONES - METRICAS B10:B11; CONCLUSIONES A4,A8
   - Problema: Precisión/fiabilidad crítica del 100 % se apoya en TP_críticos=9 de contexto y una única campaña FP pequeña. No se documenta una unidad de análisis homogénea ni intervalo de confianza.
   - Evidencia: METRICAS nota TP_criticos=9; no existe agregado local exacto enlazado; las propias limitaciones reconocen muestra pequeña y contexto confirmado.
   - Corrección: Renombrar a tasa observada en esta muestra; definir TP/FP por evento o por técnica, enlazar el agregado y evitar presentarlo como fiabilidad general.
   - Automático seguro: false
6. **02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx** - PROCESS_ROLES / RAW_SAMPLES - PROCESS_ROLES A2:K698; RAW_SAMPLES A2:N261
   - Problema: Todas las filas de detalle se almacenan como texto, incluidos PID, CPU, RAM, elapsed y booleanos. Los decimales con coma no son valores numéricos Excel.
   - Evidencia: PROCESS_ROLES: 7.678 celdas String; RAW_SAMPLES: 3.654 celdas String. En origen los campos son parseables, pero en el libro no se tiparon.
   - Corrección: En copia revisada, convertir columnas numéricas a números, timestamps a fecha/hora y RunnerActive a booleano; conservar la hoja raw original si se requiere inmutabilidad.
   - Automático seguro: true
7. **03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx** - 01_Dashboard / 16_Scripts_Validados / 17_Incoherencias_Cerradas - A1:J13; A5:E11; A5:D14
   - Problema: El libro histórico usa títulos como definitivo/estado final y describe TEC-009 sin transferencia real, P1 Critical Correlation y acciones antiguas. Contradice el estado vigente aunque el nombre externo diga REFERENCIA.
   - Evidencia: Libro fechado 15/05/2026; proyecto vigente documenta HTTP ZIP real, P1 Critical Priority y EVENT/SOC_v3.
   - Corrección: Moverlo a anexo histórico o añadir portada visible NO VIGENTE / REFERENCIA HISTÓRICA y remitir al Excel 01.
   - Automático seguro: false
8. **04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx** - 01_Dashboard / 02_Control_Pruebas / 03_Matriz_TEC / 07_Benignas_FP - B5:B13; P5:AF51; D4:M12; E5:K19
   - Problema: Es una plantilla no ejecutada: 47/47 pruebas públicas pendientes, 0 completadas y 15 benignas pendientes. No es un resultado definitivo.
   - Evidencia: Fórmulas del Dashboard y estados Pendiente en todas las filas.
   - Corrección: Mantener solo como anexo metodológico claramente rotulado PLANTILLA/PLAN NO EJECUTADO o retirarlo del paquete de resultados definitivos.
   - Automático seguro: false
9. **MANIFEST_ENTREGA_MEMORIA.csv** - (CSV) - Fila MANIFEST_ENTREGA_MEMORIA.csv y fila SHA256SUMS_ENTREGA_MEMORIA.txt
   - Problema: El manifiesto está desactualizado respecto a sí mismo y al fichero de hashes: registra MANIFEST con 0 bytes y SHA256SUMS con 5.087 bytes, cuando son 4.593 y 5.673 bytes.
   - Evidencia: Comparación con Get-Item actual.
   - Corrección: Regenerar el manifiesto excluyendo el propio manifiesto y el fichero de hashes, o generarlos en orden y documentar la exclusión.
   - Automático seguro: true
10. **SHA256SUMS_ENTREGA_MEMORIA.txt** - (texto) - Última línea
   - Problema: El fichero incluye un hash de sí mismo que no coincide con su contenido final; una suma autocontenida no puede estabilizarse.
   - Evidencia: Esperado A5DA40...; SHA256 real 624CFE01563FCD3B5B3A9AC785B0BFD86AD5F4EB7BF333A2221F42C5A489B730. Los otros 29 hashes sí coinciden.
   - Corrección: Excluir SHA256SUMS_ENTREGA_MEMORIA.txt de su propia lista y regenerar el resto.
   - Automático seguro: true
11. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx** - TRAZABILIDAD - A4:E14
   - Problema: La trazabilidad usa la campaña v5 de 16/06 y no identifica los dos summaries más recientes de 18/06, aunque el criterio de auditoría exige priorizar la evidencia posterior o justificar por qué es auxiliar.
   - Evidencia: 05_LOGS/BENCHMARKS/TFM_TEC_Run_CANDIDATE_v5_20260618_192456/192553_summary.csv: 9 OK y uploads con hashes 7AF618.../86B416....
   - Corrección: Declarar explícitamente la campaña oficial. Si se mantiene 16/06, justificar y separar las ejecuciones del 18/06 como auxiliares; si se adopta la última, recalcular todo.
   - Automático seguro: false

### MINOR (5)

1. **01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx / 02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx** - Hojas de detalle - Rangos usados
   - Problema: Los rangos de datos no están definidos como tablas Excel, aunque sí hay freeze panes y filtros en el libro 02.
   - Evidencia: ListObjects=0 en los libros 01 y 02.
   - Corrección: Convertir los rangos de detalle en tablas con nombres únicos y filtros, sin alterar datos.
   - Automático seguro: true
2. **03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx** - 00_Guia..10_Gaps_Custom_P1P4 - UsedRange A1:Z200 / A1:AE200
   - Problema: Muchas hojas tienen UsedRange inflado hasta fila 200 y columna Z/AE pese a contener pocas celdas, lo que perjudica impresión y navegación.
   - Evidencia: Metadatos Excel COM de las primeras 12 hojas.
   - Corrección: Limpiar formato sobrante en una copia y definir áreas de impresión.
   - Automático seguro: true
3. **fuentes_analisis/*** - Todas - Archivos completos
   - Problema: Hay duplicados byte a byte: el libro 03 aparece tres veces y el libro 04 dos veces.
   - Evidencia: Hash 653133D7... en 03/V4_ORIGINAL/DEF_ORIGINAL; hash CD1C04BA... en 04/Artifacts_Publicos_ORIGINAL.
   - Corrección: Conservar una sola copia por hash o documentar explícitamente original versus copia de entrega.
   - Automático seguro: true
4. **00_CONTEXT/TEST_MATRIX.md** - N/A - Archivo ausente
   - Problema: La matriz obligatoria indicada por AGENTS.md no existe. La auditoría tuvo que usar CURRENT_TASK, DECISIONS, JSONL y MATRIZ_COBERTURA_FINAL.
   - Evidencia: Get-Item devuelve ruta inexistente; la ausencia ya estaba anotada en PORCENTAJE_FIABILIDAD_CONTEXT.md.
   - Corrección: Crear TEST_MATRIX.md vigente o corregir AGENTS.md para señalar la fuente canónica.
   - Automático seguro: false
5. **Todos los .xlsx** - Todas - Revisión visual
   - Problema: Excel abrió los 44 tabs y exportó 44 PDF temporales, pero el entorno no permitió renderizar esos PDF/PNG de forma fiable. Queda revisión visual humana de clipping, colores y paginación.
   - Evidencia: Excel COM 16.0; exportación 44/44; navegador integrado y Poppler no disponibles; CopyPicture produjo PNG en blanco.
   - Corrección: Abrir los cuatro libros en Excel/LibreOffice y revisar cada hoja y gráfico al 100 % de zoom antes de entregar.
   - Automático seguro: false

### INFO (4)

1. **Todos los .xlsx** - Todas - Estructura
   - Problema: No se detectaron #REF!, #VALUE!, #N/A, #DIV/0!, #NAME?, vínculos externos, hojas ocultas, pivotes ni macros.
   - Evidencia: Inspección Excel COM de 4 libros y 44 hojas.
   - Corrección: Sin acción.
   - Automático seguro: false
2. **02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx** - SCENARIO_SUMMARY / GRAFICAS - A1:R4; 4 gráficos
   - Problema: Los recuentos de escenario y roles coinciden con los CSV normalizados; los cuatro gráficos usan rangos coherentes.
   - Evidencia: 83/82/95 muestras; roles 520 notepad + 177 server/lab + 0 cliente; series A2:B4, D4:E7, G2:H4 y J2:K4.
   - Corrección: Mantener como evidencia de control de calidad, no como benchmark definitivo.
   - Automático seguro: false
3. **TEC-009 evidencia externa** - N/A - N/A
   - Problema: Sí existe transferencia real de ZIP controlado; el problema es la ausencia de enlace desde el Excel, no la falta de transferencia.
   - Evidencia: receiver_log.jsonl + upload_20260616_154943_944765.zip: 6.245 bytes, application/zip, POST /upload, SHA256 F7E58BFDFDADB96949D7BB55DD0DEEECAA00C50F7AF4652CC6E242E870747509.
   - Corrección: Incorporar esta trazabilidad al Excel 01.
   - Automático seguro: false
4. **04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx** - 01_Dashboard - Fórmulas y gráficos
   - Problema: Las 43 fórmulas calculan sin error y las dos series de gráficos cubren sus rangos previstos.
   - Evidencia: FormulaErrorCells=0; series A16:D20 e I13:J16.
   - Corrección: Sin acción estructural; solo reclasificar el libro como plantilla histórica/metodológica.
   - Automático seguro: false

## Comprobaciones estructurales

- Apertura con Microsoft Excel 16.0 en modo solo lectura: 4/4 libros.
- Hojas inspeccionadas: 44.
- Fórmulas con error: 0.
- Vínculos externos: 0.
- Hojas ocultas/VeryHidden: 0.
- Tablas dinámicas: 0.
- Macros/VBA: 0.
- Gráficos detectados: 13.
- Series de gráficos revisadas: el único rango erróneo localizado es `GRAFICAS` Chart 1 del libro 01.
- Se exportaron 44 PDF temporales desde Excel, pero no pudieron renderizarse en este entorno. La inspección visual humana sigue pendiente.

## Correcciones prioritarias antes de entregar

1. Repetir el benchmark real y regenerar el libro 02 solo con `CLIENT_SERVICE` observado.
2. Elegir una única campaña ofensiva oficial y reconstruir `TEC_ATTACK_VALIDATION` directamente desde su JSONL.
3. Añadir trazabilidad por técnica: runner, Artifact, Source, alerta CLIENT_EVENT, timestamp, ruta y SHA256.
4. Recalcular FP principales excluyendo los tres eventos contaminados; separar bruto y limpio.
5. Corregir el rango del gráfico ofensivo a `A6:B14`.
6. Corregir las dos tablas incompletas del libro 03 o retirar ese libro del paquete final.
7. Reclasificar el libro 04 como plantilla/anexo no ejecutado.
8. Regenerar manifiesto y hashes sin auto-incluirse.
9. Abrir los cuatro Excel en Excel/LibreOffice y completar la revisión visual de todas las hojas.

## Checklist final de entrega

- [ ] Benchmark definitivo válido con cliente activo.
- [ ] Campaña ofensiva oficial identificada por RunId/ventana.
- [ ] Recuentos SOC reconciliados con JSONL.
- [ ] TEC-001..TEC-009 con alerta CLIENT_EVENT trazable.
- [ ] TEC-009 enlazado a receiver log, ZIP y SHA256.
- [ ] FP bruto y limpio separados.
- [ ] Métricas derivadas mediante fórmulas y formatos porcentuales correctos.
- [ ] Gráfico ofensivo incluye las nueve TEC.
- [ ] Tablas del libro 03 ampliadas o libro retirado.
- [ ] Libro 04 rotulado como plantilla/anexo.
- [ ] Manifiesto e integridad regenerados.
- [ ] Revisión visual manual completada.

## Dictamen por libro

- Perfectos: ninguno.
- Con errores graves: 01, 02 y 03.
- Estructuralmente coherente pero no válido como resultado: 04.
- Copias corregidas generadas: ninguna.
