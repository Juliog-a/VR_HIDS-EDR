# PROJECT_STATE

## Estado global actual

- Proyecto: TFM Velociraptor como HIDS/DFIR.

- Arquitectura custom: P4 / P3 / P2 / P1.

- P4_v2: funcional.

- P3_v5: genera mas evidencia, pero CU-009 sigue pendiente.

- P3_v6: revisado; mezclaba deteccion fuerte CU-009 con contexto opcional.

- P3_v7: creado en candidate; intacto; superado por P3_v8 para corregir separacion crypto/hash en CU-009.

- P3_v8: creado en candidate; superado por P3_v9 para CU-009 4104.

- Debug PowerShell 4104 CU009Shape v4: validado experimentalmente por el investigador; confirma forma real de upload HTTP ZIP en 4104.

- P3_v9: creado en candidate; no validado experimentalmente. En prueba live, staging local emitio fila, pero CU-009 fuerte 4104 no emitio.

- P3_v10: creado en candidate; revisa solo integracion CU-009 fuerte 4104 usando `watch_evtx` directo sobre `EventData.ScriptBlockText`, sin wrapper `foreach` live. No validado; superado arquitectonicamente por P3_v11/P2_v1.

- P3_v11: creado en candidate. Reorganiza P3 como capa media de comportamiento general y contexto. Degrada CU-007/CU-008/CU-009 fuertes a contexto no responsable de validacion critica.

- P2_v1: creado en candidate. Artifact especializado, pequeno y auditable para CU-007, CU-008 y CU-009. Integra CU009Shape_v4 dentro de P2.

- Runner v4: creado en `03_RUNNERS` y copiado a `Carpeta_Compartida_TFM/runner_candidate`. Actualiza la campana completa a la ruta activa `01_ACTIVE_TESTS\Pruebas` y ejecuta TEC-009 mediante `exfiltracion_v3.ps1` con upload HTTP controlado.

- Runner v4 revision 20260612_1215: corregido TEC-009 para invocar `exfiltracion_v3.ps1` por ruta absoluta. Anade resumen final TXT/JSON/CSV por tecnica.

- Visibility checker: creado `TFM_Check_Sysmon_Visibility.ps1` como script de solo consulta para evaluar disponibilidad de Sysmon, PowerShell, Security, System y Application.

- Ultima iteracion 20260612_1258: creado paquete `ultima_iteracion` con runner v4 corregido, checker, scripts corregidos TEC-007/TEC-008/TEC-009, logs de referencia, hashes y README. No se ejecuto campana completa desde Codex.

- Router JSONL: funcional.

- Discord: desactivado.

- P2/P1: no validados experimentalmente. P2_v1 queda candidate pendiente de importacion/validacion; P1 reservado sin implementar.

- Iteracion artifacts historicos 20260612_1418:
  - P1_v3 creado en candidate como Low Basic historico con `parse_evtx()`.
  - P2_v2 creado en candidate como High Forensic historico con cobertura 9/9.
  - P3_v12 creado en candidate como Medium Behavioral historico con salida normalizada.
  - Debug RAW 4104 y Sysmon ID 1/3/11 creados en debug.
  - Paquete `ultima_iteracion_artifacts` creado con backups, analisis, README y hashes.
  - ZIP final `ultima_iteracion_artifacts.zip` creado.
  - No se modifica runner, scripts de ataque, receiver, router ni Discord.

- Iteracion EVENT final 20260615_2137:
  - Runner v5 creado en `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`.
  - TEC-009 en runner v5 se ejecuta mediante `powershell.exe -NoProfile -ExecutionPolicy Bypass -File` para favorecer 4104 fuerte.
  - P1/P2/P3 EVENT creados en `01_ARTIFACTS/candidate` como `CLIENT_EVENT` con `watch_evtx()`.
  - P1/P2/P3 EVENT cubren TEC-001 a TEC-009.
  - P3 EVENT incluye fuente `TEC009_HTTP_ZIP_Upload_4104` y Sysmon ID 1/3/11 como contexto.
  - Script `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1` creado para recopilar evidencia local post-campana.
  - Documentacion creada en `07_DOCS`.
  - Paquete `ultima_iteracion_event_final` y ZIP `ultima_iteracion_event_final.zip` creados con hashes.
  - Validacion local estatica realizada; validacion experimental pendiente por el autor en Velociraptor.

- Router SOC Discord 20260616_1944:
  - Creado `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v1.yaml`.
  - `SERVER_EVENT` basado en `watch_monitoring()` para alertas P1/P2/P3 EVENT.
  - Normaliza salida SOC con CU, TEC, MITRE, IOA, IOC, severidad, confianza y accion recomendada.
  - Escribe JSONL y puede enviar Discord si se aporta webhook como parametro.
  - No contiene webhook hardcodeado ni rutas/IOCs de laboratorio prohibidos.
  - Pendiente importar y validar en Velociraptor.

- Router SOC Discord 20260616_2027:
  - Creado `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v2.yaml`.
  - Copiado paquete operativo a `01_ARTIFACTS/candidate/last_version`.
  - Corrige causa probable de `SOC_v1` sin eventos: mismatch de P2 (`High.Forensic.Event_v1` vs `High.Forensic_v1`), ausencia de P4 y normalizacion fragil de columnas.
  - Anade `EnableP4`.
  - Escucha P1 legacy, P1 critical opcional, P2_v1, P3 EVENT y P4_v2.
  - Crea `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` como alias metodologico compatible.
  - Crea runner `03_RUNNERS/TFM_Run_All_TEC_Tests_v6.ps1` con correccion acotada del `TEC-009 EVIDENCE CHECK`.
  - Crea `07_DOCS/GUIA_ROUTER_SOC_DISCORD_v2.md`.
  - Validacion local: parser PowerShell OK para runner v6; comprobacion estatica de router sin `parse_evtx()`, sin `watch_evtx()`, sin webhook hardcodeado y sin IOCs de laboratorio prohibidos.
  - Pendiente importar/validar VQL en Velociraptor.

- Router SOC Discord 20260617_2027:
  - Creado `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`.
  - Copiado entregable operativo a `01_ARTIFACTS/candidate/last_version`.
  - Corrige el problema observado en SOC_v2: solo se enrutaba `P1_CRITICAL` aunque P2/P3/P4 tuviesen filas en GUI.
  - Causas identificadas:
    - P2 real desplegado es `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`, no `Custom.TFM.HIDS.P2.High.Forensic_v1`.
    - SOC_v2 usaba una unica source con muchos `watch_monitoring()` vivos dentro de `chain()`, lo que puede bloquear ramas posteriores en monitorizacion continua.
  - SOC_v3 usa una source independiente por cada source monitorizada:
    - P1 Critical: 4 sources.
    - P2 Event: 5 sources.
    - P3 Event: 13 sources.
    - P4 v2: 5 sources.
  - Defaults de validacion:
    - `EnableP1Critical=true`.
    - `EnableP1=false`.
    - `EnableP2Event=true`.
    - `EnableP3=true`.
    - `EnableP4=true`.
    - `ExcludeRouterSelfEvents=false`.
  - Crea `07_DOCS/GUIA_ROUTER_SOC_DISCORD_v3.md`.
  - Validacion local estatica:
    - `type: SERVER_EVENT`;
    - 27 route sources;
    - 27 `FROM watch_monitoring`;
    - no `parse_evtx()`;
    - no `watch_evtx()`;
    - sin webhook hardcodeado;
    - sin `DUMB_LAB`, `dump_exfil`, `datos_robados` ni `exfiltracion_v3.ps1`.
  - Pendiente importar y validar VQL/ejecucion en Velociraptor.

- Recalibracion P1 Critical 20260618_1253:
  - Recalibrado `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1`.
  - Backup del P1 ruidoso creado en `candidate` y `validated/last`.
  - P1 conserva artifact name y source names para compatibilidad con `SOC_v3`.
  - P1 deja de emitir `TEC-001`, `TEC-002`, `TEC-003` y `TEC-006` genericos.
  - P1 queda limitado a `TEC-004`, `TEC-005`, `TEC-007`, `TEC-008` y `TEC-009`
    con condiciones compuestas fuertes.
  - P2/P3/P4/SOC_v3 no modificados.
  - Guia creada: `07_DOCS/GUIA_P1_P2_CRITICAL_RECALIBRATION_v1.md`.
  - Hashes creados en `01_ARTIFACTS/validated/last/SHA256SUMS_P1_P2_RECALIBRATION.txt`.
  - Pendiente importacion y validacion experimental por el autor.

- Runner falsos positivos 20260618_1446:
  - Creado `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`.
  - Objetivo: medir falsos positivos con acciones benignas/administrativas.
  - No modifica runner TEC, artifacts, router, receiver ni scripts existentes.
  - Genera logs en `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\FPs` cuando se ejecuta en VM.
  - Lee `\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl`.
  - Produce log, summary TXT/JSON/CSV y `vr_hits` TXT/CSV.
  - Cubre `FP-001` a `FP-010` con actividad benigna.
  - Hash: `D552AA137F9A59D59574D2C86F934BCEDD00AF4A0C6785A6FEE67030E4388E17`.
  - Validacion local: parser PowerShell OK; no ejecutado experimentalmente.

- Runner falsos positivos v1.1 20260618_1504:
  - Corregido `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1` manteniendo el mismo nombre.
  - Backup del runner original:
    `03_RUNNERS/TFM_Run_FP_Tests_v1_OLD_BROKEN_POSTPROCESS_BACKUP.ps1`.
  - Hash runner corregido:
    `557B9799DAA62CFD68A31EF79296997137DBB8E355175615F20F44F8AD032528`.
  - Hash backup/original:
    `D552AA137F9A59D59574D2C86F934BCEDD00AF4A0C6785A6FEE67030E4388E17`.
  - Fix acotado a postproceso:
    lectura JSONL linea a linea, tolerancia a JSON corrupto, timestamps no
    parseables, JSONL vacio, `FP_ID=UNKNOWN` y normalizacion de arrays/singletons.
  - Validacion local:
    parser PowerShell OK y ejecucion controlada con JSONL vacio genera todos los
    outputs con 0 hits.
  - Evidencias generadas en `05_LOGS/FPS`:
    `_summary.txt`, `_summary.json`, `_summary.csv`, `_vr_hits.txt`,
    `_vr_hits.csv`, `FP_ANALYSIS_CONCLUSIONS_v1.md` y hashes.
  - Campana FP `TFM_FP_20260618_144027` documentada:
    29 alertas totales comunicadas, 16 `P2_EVENT`, 9 `P3`, 4 `P4`,
    0 `P1_CRITICAL`.
  - Interpretacion:
    no hay falso positivo critico; foco de revision en `TEC-006` y `TEC-002`;
    no modificar artifacts todavia.

- Informe fiabilidad/deteccion/FP 20260618_1602:
  - Generado `PORCENTAJE FIABILIDAD.xlsx`.
  - Generados:
    - `PORCENTAJE FIABILIDAD_RESUMEN.txt`;
    - `PORCENTAJE FIABILIDAD_CONCLUSIONES.md`;
    - `00_CONTEXT/PORCENTAJE_FIABILIDAD_CONTEXT.md`;
    - `05_LOGS/FPS/README_FIABILIDAD_FP.md`;
    - `SHA256SUMS_PORCENTAJE_FIABILIDAD.txt`.
  - Excel con 8 hojas:
    `RESUMEN_EJECUTIVO`, `TEC_ATTACK_VALIDATION`, `FP_TEST_MATRIX`,
    `FP_HITS_RAW`, `METRICAS`, `GRAFICAS`, `CONCLUSIONES`, `TRAZABILIDAD`.
  - Validacion local del paquete:
    8 hojas, 6 graficas, XML clave parseable y `GRAFICAS` sin filas duplicadas.
  - Datos FP v1.1 usados desde LOG:
    10 planificadas, 9 OK, 1 SKIPPED, 0 FAIL, 15 hits,
    `P4=10`, `P3=5`, `P2_EVENT=0`, `P1_CRITICAL=0`.
  - Calidad FP:
    `EXACT_RUNID=3`, `TIME_WINDOW_INFERRED=9`,
    `PREVIOUS_RUN_CONTAMINATION=3`, `UNKNOWN=0`.
  - Se separa contaminacion temporal de `TFM_FP_20260618_151244`.
  - Conclusion:
    cobertura ofensiva 9/9, fiabilidad critica 100 % en esta muestra,
    sin FP criticos ni forenses en FP v1.1.

- Benchmark rendimiento Velociraptor 20260618_1726:
  - Creado `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1`.
  - Creado `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1`.
  - Generados:
    - `VR_RESOURCE_BENCHMARK.xlsx`;
    - `VR_RESOURCE_BENCHMARK_RESUMEN.txt`;
    - `VR_RESOURCE_BENCHMARK_CONCLUSIONES.md`;
    - `00_CONTEXT/VR_RESOURCE_BENCHMARK_CONTEXT.md`;
    - `05_LOGS/BENCHMARKS/README_BENCHMARKS.md`;
    - `SHA256SUMS_VR_RESOURCE_BENCHMARK.txt`.
  - Estado actual:
    `PENDIENTE_DE_EJECUCION`, sin datos reales de benchmark en host/VM.
  - Excel generado como plantilla trazable con 8 hojas y 0 graficas reales.
  - Validacion local:
    parser PowerShell OK en ambos scripts y paquete OpenXML del Excel abrible.
  - No se modificaron artifacts, SOC_v3, runner TEC v6, runner FP, receiver,
    scripts TEC ni JSONL.

- Benchmark rendimiento Velociraptor fix 20260618_1900:
  - Corregido `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1`.
  - Se implementa `SkipServiceControl` en el `param()` real como `[bool]`.
  - Parametros finales:
    `ScenarioName`, `DurationSec`, `IntervalSec`, `RunnerPath`, `OutputDir`,
    `StopVRBeforeRun`, `StartVRAfterStop`, `SkipServiceControl`, `WarmupSec`,
    `IncludeServerGuiInTotal`.
  - Se separan roles:
    `CLIENT_SERVICE`, `SERVER_GUI`, `OTHER_VELOCIRAPTOR`.
  - `SERVER_GUI` queda excluido de metricas/conclusiones del agente cliente.
  - Se evitan ceros silenciosos para metricas no disponibles; se usa `NA/null`
    y se excluyen de medias.
  - Generador Excel actualizado con hoja `CALIDAD_DATOS`.
  - Benchmarks previos en `05_LOGS\BENCHMARKS` marcados como
    `DATOS_PARCIALES_O_LEGACY`.
  - Validacion local:
    parser OK, parametros detectados OK, `SkipServiceControl` invocable.

- Revision benchmarks existentes 20260618_1950:
  - Revisado `05_LOGS/BENCHMARKS`.
  - Generado paquete consolidado en `05_LOGS/BENCHMARKS/validado`.
  - Exceles generados:
    - `VR_RESOURCE_BENCHMARK_VALIDACION_DATOS.xlsx`;
    - `VR_RESOURCE_BENCHMARK.xlsx`.
  - Hojas/graficas:
    8 hojas y 4 graficas OpenXML.
  - Estado:
    `NO_VALIDO_PARA_CONCLUSIONES_DEFINITIVAS_DE_RENDIMIENTO`.
  - Motivos:
    - `SchemaVersion=1.0`;
    - `VR_IDLE` y `VR_TEC_RUNNER` con `ServiceStatus=Stopped`;
    - no se observa `CLIENT_SERVICE`;
    - `notepad.exe` capturado por ruta con palabra `Velociraptor`;
    - binario `velociraptor-v0.75.6-windows-amd64.exe` separado como
      server/laboratorio.
  - Decision:
    no usar estas metricas para afirmar impacto CPU/RAM del agente cliente;
    repetir benchmark con script corregido.

- Entrega Exceles memoria 20260618_1959:
  - Creada carpeta:
    `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS`.
  - Excel principal de deteccion/FP/fiabilidad:
    `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx`
    con 8 hojas y 6 graficas.
  - Excel benchmark:
    `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx`
    con 8 hojas y 4 graficas; estado de calidad de datos, no consumo definitivo.
  - Referencias/anexos:
    - `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx`;
    - `04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx`.
  - Generados:
    `README_ENTREGA_MEMORIA.md`, `NOTAS_CALIDAD_DATOS.md`,
    `MANIFEST_ENTREGA_MEMORIA.csv` y `SHA256SUMS_ENTREGA_MEMORIA.txt`.

- Auditoria Exceles definitivos 20260618_2105:
  - Revisados 4 Excel superiores, 3 Excel duplicados y 7 CSV.
  - Inspeccionadas 44 hojas y 13 graficos con Excel 16.0 en modo solo lectura.
  - Estado global: `NO APTO`.
  - Hallazgos: 7 BLOCKER, 11 MAJOR, 5 MINOR y 4 INFO.
  - Blockers principales:
    - falta benchmark definitivo del cliente;
    - recuentos SOC del libro 01 no reconciliados con un JSONL unico;
    - trazabilidad insuficiente desde TEC hasta alerta CLIENT_EVENT;
    - grafico ofensivo omite TEC-009;
    - dos tablas del libro 03 tienen rangos incompletos.
  - TEC-009 si tiene transferencia real y ZIP con hash coincidente; falta
    enlazar esta evidencia en el libro 01.
  - Entregables generados en:
    `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS_REVIEWED`.
  - Originales no modificados.

- Paquete de repetición controlada 20260619_1157:
  - Creada carpeta `06_CONTROLLED_RERUN`.
  - Incluye preflight, receiver aislado, wrapper TEC canónico, FP 3x,
    benchmark 3x3, validador, staging/transferencia y README.
  - Reconstruido `00_CONTEXT/TEST_MATRIX.md` con separación entre ejecución,
    `CLIENT_EVENT`, router, JSONL y evidencia externa.
  - FP-009 dispone de listener benigno controlado y no debe quedar `SKIPPED`.
  - Benchmark usa carga TEC sin exfiltración y separa escenario/repetición.
  - Verificación local: 13 scripts parseables, 9 hashes fuente coincidentes,
    package hash 15/15 y staging hash 25/25.
  - No se modificaron artifacts `validated` ni se regeneraron Excel.
  - Pendiente ejecución experimental por el autor y `VALIDACION PASS`.

- Patch runner TEC 4104 20260619_1223:
  - `TEC_20260619_FINAL01` ejecutó transferencia TEC-009 real, pero queda
    `NO APTA` por ausencia de summary global.
  - Causa: dos conversiones `Generic.List` con `@(...)` incompatibles con el
    binder de Windows PowerShell 5.1.
  - Creado sin sobrescribir v5/v6:
    `03_RUNNERS/TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1`.
  - Hash:
    `C023F8F7DC4D90977C4F6C76C8147A90C3558D62E14D9E1A265358DCDFF68367`.
  - Wrapper controlado actualizado para usar el patched.
  - Selftest mock: 49 eventos, 15 samples, 9 técnicas serializadas, PASS.
  - Pendiente repetir `TEC_20260619_FINAL02`.

- Validacion experimental: manual por el autor.

## Validación controlada FINAL02 20260619_2226

- Combinación canónica fijada:
  - TEC `TEC_20260619_FINAL02`;
  - FP `FP_20260619_FINAL01`;
  - benchmark `BENCH_20260619_FINAL01`;
  - transferencia final `RERUN_20260619_FINAL02` pendiente.
- TEC FINAL02 queda `APTO` con 9/9 runner OK, transferencia TEC-009 real y
  199 filas CLIENT_EVENT cronológicamente filtradas con cobertura 9/9.
- Benchmark queda `APTO CON OBSERVACIONES`: nueve runs válidos; el runner
  seguía activo al cerrar las ventanas de carga, sin invalidar las muestras.
- FP queda `NO APTO` hasta exportar CLIENT_EVENT posterior a REP_03 y adjudicar
  P2/UNKNOWN.
- El validador ya no aborta por propiedades `.Name` ausentes.
- El validador prioriza `DetectionTime` original frente a `Timestamp` del
  router y detecta contaminación tardía en REP_02/REP_03.
- Creado helper `10_Build_ClientEvent_Exports.ps1`; no sobrescribe y rechaza
  exports fuente anteriores a la campaña.
- Transferencia selectiva e importación idempotente implementadas en 07/08.
- Bundle `controlled_rerun_bundle_validation_patch`: 28/28 hashes correctos.
- No se regeneraron Excel.

## Flujo metodologico

Sysmon / PowerShell 4104

-> CLIENT_EVENT custom

-> SERVER_EVENT router

-> JSONL

-> analisis posterior

## Criterio critico

Telemetria no equivale a deteccion.

JSONL no equivale a deteccion.

Discord no equivale a deteccion.

Codex no valida resultados.

## Regeneración benchmark controlado (2026-06-19 23:32 CEST)

- Regenerado el Excel de rendimiento exclusivamente desde
  `06_CONTROLLED_RERUN/OUTPUT/BENCHMARK/BENCH_20260619_FINAL01`.
- No se usó ni modificó `05_LOGS/BENCHMARKS/validado/VR_RESOURCE_BENCHMARK.xlsx`.
- Fuente: 9 `summary.json` (3 `BASELINE_NO_VR`, 3 `VR_IDLE`, 3
  `VR_TEC_RUNNER`), todos `SchemaVersion=1.1` y `ScenarioValidity=VALID`.
- Control de cliente: 284 filas `CLIENT_SERVICE`, 0 `notepad.exe` como runner,
  `SERVER_GUI` excluido en 9/9 y runner observado en 3/3 ejecuciones de carga.
- `RunnerStillRunningAtEnd=True` aparece en las 3 ejecuciones
  `VR_TEC_RUNNER` y se conserva como `WARN` no bloqueante según D042.
- Entregable:
  `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS/02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`.
- Estado del Excel regenerado: `APTO CON OBSERVACIONES`.
- SHA-256:
  `885515561DC2096D4A310D08337CD86DED6D84B3579A4092641987271A652B8C`.

## Excel visibilidad/deteccion/FP/Wazuh V.5 (2026-06-26 11:43 CEST)

- Generado Excel final desde `04_EVIDENCE/Analisis_Tecnicas_TFM_V.4.xlsx`.
- Entregables principales:
  - `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`.
  - `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_V.5.xlsx`.
- Copias trazables y normalizados:
  - `04_EVIDENCE/Excel_visibilidad_26062026`.
- Contexto especifico:
  - `00_CONTEXT/CONTEXT_ESTADO_VISIBILIDAD_FP_26062026.md`.
- Estado VR REAL 24/06:
  - 9/9 TEC presentes.
  - 9 OK, 0 WARN, 0 FAIL.
  - 379 filas `CLIENT_EVENT` incluidas.
  - Perfiles: P1=19, P2=118, P3=109, P4=133.
- Estado FP 26/06:
  - 10/10 FP presentes.
  - 10 OK, 0 WARN, 0 FAIL, 0 SKIPPED.
  - `VelociraptorHits.Total=0`.
- Estado Wazuh:
  - Base: 902 archives, 58 alerts.
  - Custom: 18.285 archives, 110 alerts, 56 alertas 110xxx.
  - 110201=0 conservado como gap TEC-009.
- Verificacion Excel:
  - Apertura Excel COM readonly OK.
  - 31 hojas, 12/12 obligatorias.
  - 5 graficas nativas y 2 heatmaps.
  - 0 errores de formula buscados.
- Limitaciones:
  - 48 discrepancias de fechas/ventanas auditadas.
  - 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas con `vr_hits`; no
    se usan como FP definitivos.
- TEC-009 tiene HTTP 200, `UploadSucceeded=True`, SHA-256 y bytes en
  `ReceiverResponse`, pero no se localizo `receiver_log.jsonl` ni ZIP
  recibido; no afirmar exfiltracion contextual completa sin matiz.

## Excel benchmark final (2026-06-26 13:02 CEST)

- Generado Excel final de benchmark desde la fuente canonica
  `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS/02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`.
- Campana trazable:
  `06_CONTROLLED_RERUN/OUTPUT/BENCHMARK/BENCH_20260619_FINAL01`.
- Entregables principales:
  - `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`.
  - `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_AUDIT.md`.
  - `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_DATA_QUALITY.json`.
- Copias trazables, precheck, preview PDF y CSV normalizados:
  `04_EVIDENCE/Excel_benchmark_26062026`.
- Contexto especifico:
  `00_CONTEXT/CONTEXT_ESTADO_BENCHMARK_26062026.md`.
- Estado:
  - Precheck APTO.
  - 9/9 runs validos.
  - 3 escenarios y 3 repeticiones por escenario.
  - 422 filas `samples.csv` y 770 filas `process_samples.csv`.
  - 284 filas `CLIENT_SERVICE`.
  - `SERVER_GUI` excluido del calculo principal.
  - `notepad.exe` runner: 0.
- Metricas principales:
  - CPU media: `BASELINE_NO_VR=0`, `VR_IDLE=0`, `VR_TEC_RUNNER=0,95`.
  - RAM media MB: `BASELINE_NO_VR=0`, `VR_IDLE=57,48`, `VR_TEC_RUNNER=57,42`.
  - Pico CPU cliente VR: 54%.
  - Pico RAM cliente VR: 57,97 MB.
- Verificacion Excel:
  - 11 hojas obligatorias.
  - 6 graficas nativas en `GRAFICAS`.
  - 0 errores de formula.
  - fuentes con SHA-256.
- SHA-256:
  `AEB8835BAD0B2A8D015D399DDA11A2CD475B3F320750ABD333750026456A36EC`.
- Limitacion:
  - Benchmark de coste operativo del laboratorio, no de calidad de deteccion.
  - `RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER` se conserva como WARN no bloqueante.
