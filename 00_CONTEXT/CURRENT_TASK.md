# CURRENT_TASK

## Tarea activa actualizada 20260705_1216

Version beta documental generada: `CERRADA LOCALMENTE`.

Entregables:

- `08_MEMORIA/TFM_MEMORIA_BETA_v1.docx`.
- `08_MEMORIA/BETA_CHANGELOG.txt`.

Cambios principales:

- Capitulos VI, VII, IX y X completados con datos finales existentes.
- Comparativa Wazuh - Velociraptor actualizada.
- JSONL/Discord corregidos como salida externa, no deteccion.
- Referencias y anexos recomendados incorporados.

Control:

- `TFM.docx` original no modificado.
- `01_ARTIFACTS/validated` no modificado.
- No se han modificado scripts, artifacts, runners, evidencias ni Exceles.
- No se han ejecutado pruebas experimentales.

Pendiente:

1. Abrir la beta en Word y revisar visualmente pagina a pagina.
2. Normalizar bibliografia al formato requerido por la universidad.
3. Confirmar anexos y capturas finales.
4. Al volver al laboratorio, revisar si existe receiver_log.jsonl/ZIP final de
   TEC-009 o mantener el matiz documental.

## Tarea activa actualizada 20260629_1707

README GitHub y limpieza/ordenacion de raiz: `CERRADA LOCALMENTE`.

Entregables:

- `README.md`.
- `00_CONTEXT/CLEANUP_20260629.md`.
- `04_EVIDENCE/legacy_reports/README.md`.
- `99_ARCHIVE/README.md`.

Movimientos:

- Informes legacy de fiabilidad/FP:
  `04_EVIDENCE/legacy_reports/20260618_fiabilidad_fp`.
- Benchmark legacy:
  `04_EVIDENCE/legacy_reports/20260618_benchmark_legacy`.
- Iteraciones historicas:
  `99_ARCHIVE/legacy_iterations`.
- Prompts antiguos:
  `99_ARCHIVE/legacy_prompts_20260610`.
- Candidatos locales de descarte:
  `99_REVIEW_CLEANUP_20260629/delete_candidates`.

Control:

- `01_ARTIFACTS/validated` no modificado.
- VM local `lab/` no movida por riesgo de romper VirtualBox.
- `Carpeta_Compartida_TFM/` no movida por ser intercambio con VM.
- JSONL bruto, OneNote y `.lnk` retirados del indice Git sin borrado fisico.

Pendiente opcional:

- Revisar `git status` y confirmar/commit de la reorganizacion.
- Si GitHub rechaza por objetos historicos grandes, limpiar historial Git con
  herramienta especifica; no se hizo en esta sesion.
- `git gc --prune=now` ya se ejecuto: basura Git 0 bytes; pack historico
  restante 3.08 GiB.

## Tarea activa actualizada 20260626_1143

Generacion del Excel V.5/final de visibilidad, deteccion, falsos positivos y
comparacion Wazuh: `CERRADA LOCALMENTE`.

Entregables:

- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`
- `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_V.5.xlsx`
- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_AUDIT.md`
- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_DATA_QUALITY.json`
- Copias, normalizados, log y preview en:
  `04_EVIDENCE/Excel_visibilidad_26062026`

Control de aceptacion:

- TEC-001 a TEC-009: SI.
- FP-001 a FP-010: SI.
- VR real separado de FP: SI, mediante ventanas y discrepancias auditadas.
- Wazuh base/custom: SI.
- 110201=gap: SI.
- Graficas pobladas: SI, 5 graficas nativas.
- Hoja DISCREPANCIAS: SI.
- Hoja FUENTES con hashes: SI, 102 fuentes.
- Audit `.md`: SI.
- Context actualizado: SI.

Pendiente:

- Revision manual del autor antes de incorporar a memoria.
- Si se quiere afirmar TEC-009 como exfiltracion contextual completa, aportar
  `receiver_log.jsonl` y ZIP recibido; mientras tanto queda con limitacion.

## Tarea activa actualizada 20260619_2226

Cerrar la evidencia `CLIENT_EVENT` de FP y repetir la validación final sin
mezclar CampaignId.

Estado:

- `TEC_20260619_FINAL02`: `APTO`.
  - runner 9/9 OK;
  - JSONL 185 válido / 0 corrupto;
  - TEC-009 HTTP 200 y SHA-256 origen/destino;
  - CLIENT_EVENT consolidado: 199 filas, 0 fuera de ventana, TEC 9/9.
- `BENCH_20260619_FINAL01`: `APTO CON OBSERVACIONES`.
  - 9/9 SchemaVersion 1.1 y `VALID`;
  - cliente observado, SERVER_GUI excluido, notepad no runner;
  - `RunnerStillRunningAtEnd=True` en tres runs no invalida
    `RunnerObservedSamples > 0`.
- `FP_20260619_FINAL01`: `NO APTO` todavía.
  - faltan exports CLIENT_EVENT posteriores a las tres repeticiones;
  - REP_02 tiene 1 fila tardía de REP_01;
  - REP_03 tiene 36/47 filas con DetectionTime de REP_02;
  - P2 se conserva como FP/visibilidad forense; P1=0;
  - UNKNOWN requiere adjudicación con CLIENT_EVENT original.

Pendiente:

1. Instalar `controlled_rerun_bundle_validation_patch` en la VM.
2. Exportar 27 sources CLIENT_EVENT FP en `10:44:00Z..10:49:00Z`.
3. Ejecutar `10_Build_ClientEvent_Exports.ps1` para REP_01..REP_03.
4. Exportar selectivamente con `RERUN_20260619_FINAL02`.
5. Importar y validar TEC FINAL02 + FP FINAL01 + BENCH FINAL01.
6. No regenerar Excel hasta cerrar FP.

Guías:

- `06_CONTROLLED_RERUN/CLIENT_EVENT_EXPORT_GUIDE.md`.
- `06_CONTROLLED_RERUN/VALIDATION_STATUS_FINAL02.md`.

## Tarea activa actualizada 20260619_1223

Repetir campaña TEC con parche Windows PowerShell 5.1.

Estado:

- `TEC_20260619_FINAL01`: `NO APTA`; no usar como campaña final.
- Transferencia TEC-009 FINAL01 real, pero sin summary global.
- Runner patched creado y validado estáticamente:
  `03_RUNNERS/TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1`.
- Nuevo CampaignId obligatorio: `TEC_20260619_FINAL02`.

Orden:

1. Crear bundle `controlled_rerun_bundle_patch_4104` en host.
2. Instalar en VM con `-AllowPackageUpdate`.
3. Ejecutar `00_Preflight_Controlled_Rerun.ps1`.
4. Ejecutar `09_Test_Patched_4104.ps1`; exigir PASS.
5. Iniciar receiver host con `TEC_20260619_FINAL02`.
6. Ejecutar wrapper TEC con el mismo CampaignId.
7. No aceptar la campaña si runner/global summary no quedan `OK`.

## Tarea activa actualizada 20260619_1157

Ejecutar la repetición controlada preparada en:

```text
C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
```

Estado:

- Paquete listo y autocomprobado.
- `00_CONTEXT/TEST_MATRIX.md` reconstruido.
- No se han ejecutado nuevas pruebas experimentales.
- No se han regenerado CSV derivados ni Excel.
- Dictamen de entrega sigue `NO APTO` hasta validar datos nuevos.

Orden pendiente:

1. Ejecutar `05_Stage_For_VM.ps1` en host.
2. Ejecutar `06_Install_On_VM.ps1` y preflight en VM.
3. Confirmar monitoring activo en GUI.
4. Ejecutar receiver host y campaña TEC canónica VM.
5. Exportar P1/P2/P3/P4 `CLIENT_EVENT`.
6. Ejecutar FP tres repeticiones y exportar `CLIENT_EVENT` por ventana.
7. Ejecutar benchmark 3x3.
8. Transferir resultados e importar en host.
9. Ejecutar `04_Validate_Controlled_Rerun.ps1`.
10. Reconstruir CSV/Excel únicamente tras `VALIDACION PASS`.

## Tarea activa actualizada 20260618_2105

Corregir la entrega de Exceles tras auditoria tecnica.

Estado:

- Dictamen global: `NO APTO`.
- Informe:
  `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS_REVIEWED/AUDITORIA_EXCELES_DEFINITIVOS.md`.
- 7 BLOCKER, 11 MAJOR, 5 MINOR y 4 INFO.
- Excel originales intactos.

Pendiente prioritario:

1. Repetir benchmark con script corregido y `CLIENT_SERVICE` observado.
2. Elegir una campaña ofensiva canónica y recalcular `Hits SOC`/perfil por TEC
   directamente desde su JSONL.
3. Añadir trazabilidad por TEC hasta Artifact, Source, timestamp, JSONL y hash.
4. Separar FP bruto 15 de FP limpio 12.
5. Corregir el gráfico ofensivo del libro 01 para usar `A6:B14`.
6. Corregir tablas del libro 03 o retirarlo del paquete final.
7. Rotular el libro 04 como plantilla/anexo no ejecutado.
8. Regenerar manifiesto y hashes sin auto-inclusión.
9. Completar revisión visual manual.

## Tarea activa actualizada 20260618_1959

Carpeta de Exceles para memoria generada:

```text
C:\Users\julio\Desktop\TFM\04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_DEFINITIVOS
```

Exceles principales:

- `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx`
- `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx`

Exceles de referencia/anexo:

- `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx`
- `04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx`

Documentos de apoyo:

- `README_ENTREGA_MEMORIA.md`
- `NOTAS_CALIDAD_DATOS.md`
- `MANIFEST_ENTREGA_MEMORIA.csv`
- `SHA256SUMS_ENTREGA_MEMORIA.txt`

Estado:

- El Excel TEC+FP+fiabilidad esta listo para memoria.
- El Excel benchmark esta listo como validacion de calidad de datos, no como
  consumo definitivo del agente.

Pendiente:

1. Abrir visualmente los Excel en Excel/LibreOffice.
2. Repetir benchmark real cuando sea posible con `CLIENT_SERVICE` observado.
3. Si se repite benchmark correctamente, sustituir el Excel 02 por version final
   de consumo real.

## Tarea activa actualizada 20260618_1950

Revision de `05_LOGS/BENCHMARKS` y consolidacion en `validado`.

Entregables:

- `05_LOGS/BENCHMARKS/validado/VR_RESOURCE_BENCHMARK_VALIDACION_DATOS.xlsx`
- `05_LOGS/BENCHMARKS/validado/VR_RESOURCE_BENCHMARK.xlsx`
- `05_LOGS/BENCHMARKS/validado/VR_RESOURCE_BENCHMARK_VALIDACION_DATOS_RESUMEN.txt`
- `05_LOGS/BENCHMARKS/validado/VR_RESOURCE_BENCHMARK_VALIDACION_DATOS_CONCLUSIONES.md`
- `05_LOGS/BENCHMARKS/validado/scenario_quality.csv`
- `05_LOGS/BENCHMARKS/validado/process_role_classification.csv`
- `05_LOGS/BENCHMARKS/validado/raw_samples_normalized.csv`
- `05_LOGS/BENCHMARKS/validado/SHA256SUMS_BENCHMARKS_VALIDADO.txt`
- `05_LOGS/BENCHMARKS/validado/fuentes/`

Estado:

- `NO_VALIDO_PARA_CONCLUSIONES_DEFINITIVAS_DE_RENDIMIENTO`.
- Se localizaron los tres escenarios, pero proceden de esquema legacy.
- No se observo `CLIENT_SERVICE`.
- `VR_IDLE` y `VR_TEC_RUNNER` muestran `ServiceStatus=Stopped`.
- El Excel generado documenta la invalidacion con 8 hojas y 4 graficas.

Pendiente inmediato:

1. Copiar a la VM el benchmark corregido:

```text
C:\Users\julio\Desktop\TFM\03_RUNNERS\TFM_Benchmark_VR_Resource_Usage_v1.ps1
```

2. Ejecutar prueba corta en VM:

```powershell
cd "C:\Users\seguridad\Desktop\TFM\03_RUNNERS"

.\TFM_Benchmark_VR_Resource_Usage_v1.ps1 `
  -ScenarioName "TEST_SHORT" `
  -DurationSec 10 `
  -IntervalSec 2 `
  -SkipServiceControl $true
```

3. Repetir los tres escenarios solo si aparece `CLIENT_SERVICE` cuando
   Velociraptor esta activo.

4. Copiar resultados nuevos a `05_LOGS/BENCHMARKS` y regenerar Excel definitivo.

## Tarea activa actualizada 20260618_1900

Benchmark de rendimiento corregido tras error `NamedParameterNotFound`:

- `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1`
- `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1`

Backups:

- `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1_PRE_SKIP_PARAM_BROKEN_BACKUP.ps1`
- `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1_PRE_ROLE_FIX_BACKUP.ps1`

Validacion local:

- `POWERSHELL_PARSE_OK` benchmark.
- `POWERSHELL_PARSE_OK` generador Excel.
- `SkipServiceControlExists=True`.
- `RequiredParametersAllPresent=True`.
- `TopLevelParamBlockCount=1`.
- Excel regenerado con 9 hojas.

Estado de datos:

- `VR_RESOURCE_BENCHMARK.xlsx` queda en estado `DATOS_PARCIALES_O_LEGACY`.
- Los benchmarks anteriores no son definitivos porque no separaban correctamente
  cliente Velociraptor HIDS y server/GUI local.

Pendiente inmediato:

1. Copiar a VM:

```text
C:\Users\julio\Desktop\TFM\03_RUNNERS\TFM_Benchmark_VR_Resource_Usage_v1.ps1
```

2. Ejecutar prueba corta:

```powershell
cd "C:\Users\seguridad\Desktop\TFM\03_RUNNERS"

.\TFM_Benchmark_VR_Resource_Usage_v1.ps1 `
  -ScenarioName "TEST_SHORT" `
  -DurationSec 10 `
  -IntervalSec 2 `
  -SkipServiceControl $true
```

3. Si la prueba corta funciona, repetir los tres escenarios de
   `05_LOGS\BENCHMARKS\INSTRUCCIONES_REPETIR_BENCHMARK_v3.md`.

4. Copiar las carpetas nuevas `TFM_BENCH_*` al host y regenerar Excel.

## Tarea activa actualizada 20260618_1726

Benchmark de impacto en rendimiento de Velociraptor preparado:

- `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1`
- `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1.ps1`
- `VR_RESOURCE_BENCHMARK.xlsx`
- `VR_RESOURCE_BENCHMARK_RESUMEN.txt`
- `VR_RESOURCE_BENCHMARK_CONCLUSIONES.md`
- `00_CONTEXT/VR_RESOURCE_BENCHMARK_CONTEXT.md`
- `05_LOGS/BENCHMARKS/README_BENCHMARKS.md`
- `SHA256SUMS_VR_RESOURCE_BENCHMARK.txt`

Estado:

- No hay ejecuciones reales de benchmark todavia.
- El Excel esta en estado `PENDIENTE_DE_EJECUCION`.
- No se han calculado metricas reales de CPU/RAM/IO.
- No hay graficas reales porque no hay datos.

Pendiente inmediato:

1. Ejecutar en VM:

```powershell
cd C:\Users\seguridad\Desktop\TFM\03_RUNNERS
.\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName "BASELINE_NO_VR" -DurationSec 180 -IntervalSec 2 -StopVRBeforeRun $true
.\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName "VR_IDLE" -DurationSec 180 -IntervalSec 2 -StartVRAfterStop $true -WarmupSec 60
.\TFM_Benchmark_VR_Resource_Usage_v1.ps1 -ScenarioName "VR_TEC_RUNNER" -DurationSec 300 -IntervalSec 2 -RunnerPath "C:\Users\seguridad\Desktop\TFM\03_RUNNERS\TFM_Run_All_TEC_Tests_v6.ps1"
```

2. Copiar resultados desde:

```text
C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\BENCHMARKS
```

a:

```text
C:\Users\julio\Desktop\TFM\05_LOGS\BENCHMARKS
```

3. Reejecutar:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\julio\Desktop\TFM\03_RUNNERS\TFM_Generate_VR_Benchmark_Excel_v1.ps1
```

4. Revisar visualmente `VR_RESOURCE_BENCHMARK.xlsx` con datos reales.

## Tarea activa actualizada 20260618_1602

Informe de fiabilidad/deteccion/falsos positivos generado:

- `PORCENTAJE FIABILIDAD.xlsx`
- `PORCENTAJE FIABILIDAD_RESUMEN.txt`
- `PORCENTAJE FIABILIDAD_CONCLUSIONES.md`
- `00_CONTEXT/PORCENTAJE_FIABILIDAD_CONTEXT.md`
- `05_LOGS/FPS/README_FIABILIDAD_FP.md`
- `SHA256SUMS_PORCENTAJE_FIABILIDAD.txt`

Estado:

- Cobertura ofensiva documentada: 9/9 = 100 %.
- FP v1.1 documentado:
  - 10 planificadas;
  - 9 OK;
  - 1 SKIPPED;
  - 0 FAIL;
  - 15 hits;
  - `P4=10`;
  - `P3=5`;
  - `P2_EVENT=0`;
  - `P1_CRITICAL=0`.
- Fiabilidad critica: 100 % para esta muestra.

Calidad de datos:

- `EXACT_RUNID=3`.
- `TIME_WINDOW_INFERRED=9`.
- `PREVIOUS_RUN_CONTAMINATION=3`.
- `UNKNOWN=0`.
- RunId anterior contaminante separado: `TFM_FP_20260618_151244`.

Pendiente inmediato:

1. Abrir visualmente `PORCENTAJE FIABILIDAD.xlsx` en Excel/LibreOffice.
2. Revisar graficas y formato final.
3. Si se requiere validez estadistica, ampliar muestra FP.
4. No modificar P1/P2/P3/P4 solo con estos resultados.

## Tarea activa actualizada 20260618_1504

Runner FP corregido:

- `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`
- Hash: `557B9799DAA62CFD68A31EF79296997137DBB8E355175615F20F44F8AD032528`

Backup del runner con postproceso roto:

- `03_RUNNERS/TFM_Run_FP_Tests_v1_OLD_BROKEN_POSTPROCESS_BACKUP.ps1`
- Hash: `D552AA137F9A59D59574D2C86F934BCEDD00AF4A0C6785A6FEE67030E4388E17`

Objetivo inmediato:

1. Copiar el runner corregido a la VM si procede.
2. Mantener P1 recalibrado, P2/P3/P4 y SOC_v3 activos.
3. Limpiar `soc_alerts.jsonl` con backup previo si procede.
4. Ejecutar de nuevo:

```powershell
cd C:\Users\seguridad\Desktop\TFM\03_RUNNERS
.\TFM_Run_FP_Tests_v1.ps1 -ClearJsonlBeforeRun $true -KeepWorkspace $true
```

5. Confirmar que se generan sin error:
   - `_summary.txt`
   - `_summary.json`
   - `_summary.csv`
   - `_vr_hits.txt`
   - `_vr_hits.csv`
6. Confirmar que `P1_CRITICAL` sigue en 0.
7. Comparar outputs nativos contra `05_LOGS/FPS/FP_ANALYSIS_CONCLUSIONS_v1.md`.

Estado documentado de la campana `TFM_FP_20260618_144027`:

- Total alertas comunicado: 29.
- `P2_EVENT`: 16.
- `P3`: 9.
- `P4`: 4.
- `P1_CRITICAL`: 0.
- Por `RunId` reconstruido desde evidencia local: `P2_EVENT=6`, `P4=4`, `P1_CRITICAL=0`.

Decision operativa:

- No modificar P1/P2/P3/P4 todavia.
- No modificar SOC_v3.
- Documentar `TEC-006` y `TEC-002` como foco principal de revision FP.

## Tarea activa

Validar experimentalmente la iteracion final EVENT:

- `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`
- `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1`

Paquete:

- `ultima_iteracion_event_final`
- `ultima_iteracion_event_final.zip`
- `ultima_iteracion_event_final.zip.sha256`

## Objetivo inmediato

1. Revisar checklist previo:
   - `07_DOCS/CHECKLIST_PREVIO_LOGGING_VR.md`
2. Importar P1/P2/P3 EVENT en Velociraptor.
3. Activarlos como Client Event Monitoring antes de ejecutar la campana.
4. Arrancar receiver en el host.
5. Ejecutar runner v5 en la VM Windows.
6. Confirmar que TEC-009 genera:
   - ZIP real;
   - copia ZIP;
   - SHA256;
   - HTTP 2xx;
   - `UploadSucceeded=True`;
   - summary JSON;
   - 4104 con `Invoke-WebRequest`, `-Method POST`, `-InFile`, `application/zip`, `ReceiverUrl` y `.zip`.
7. Revisar eventos Velociraptor, especialmente P3:
   - `TEC009_HTTP_ZIP_Upload_4104`;
   - `TEC009_HTTP_Upload_Sysmon_Process_Context`;
   - `TEC009_HTTP_Network_Sysmon_Context`;
   - `TEC009_ZIP_FileCreate_Sysmon_Context`.
8. Ejecutar recogida de evidencia:

```powershell
.\03_RUNNERS\TFM_Collect_Final_Evidence_v1.ps1
```

## Artifacts EVENT creados

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`

## Router SOC candidate

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v1.yaml`
- Guia: `07_DOCS/GUIA_ROUTER_SOC_DISCORD_v1.md`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v2.yaml`
- Guia: `07_DOCS/GUIA_ROUTER_SOC_DISCORD_v2.md`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`
- Guia: `07_DOCS/GUIA_ROUTER_SOC_DISCORD_v3.md`
- Paquete operativo: `01_ARTIFACTS/candidate/last_version`

Pendiente:

1. Importar `SOC_v3` en Velociraptor.
2. Desactivar `SOC_v2` y activar `SOC_v3` como Server Event Monitoring despues de P1/P2/P3/P4 EVENT.
3. Probar primero con `EnableDiscord=false` y `EnableJSONL=true`.
4. Confirmar escritura de `\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl`.
5. Activar Discord con `EnableDiscord=true` y `DiscordWebhook=<webhook>` solo en despliegue.
6. Usar runner `03_RUNNERS/TFM_Run_All_TEC_Tests_v6.ps1` para evitar el fallo del bloque `TEC-009 EVIDENCE CHECK`.
7. Validar que `Group-Object Profile` muestra mas de `P1_CRITICAL`, idealmente `P2_EVENT`, `P3` y `P4`.

## P1 Critical recalibrado

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml`
- `01_ARTIFACTS/validated/last/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml`
- Backup ruidoso:
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1_PRE_NOISY_BACKUP.yaml`
  - `01_ARTIFACTS/validated/last/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1_PRE_NOISY_BACKUP.yaml`
- Guia:
  - `07_DOCS/GUIA_P1_P2_CRITICAL_RECALIBRATION_v1.md`
- Hashes:
  - `01_ARTIFACTS/validated/last/SHA256SUMS_P1_P2_RECALIBRATION.txt`

Pendiente:

1. Importar el P1 recalibrado.
2. Mantener P2/P3/P4 y SOC_v3 activos.
3. Limpiar `soc_alerts.jsonl`.
4. Ejecutar runner v6.
5. Confirmar que `P1_CRITICAL` baja de forma drastica.
6. Confirmar que P2/P3/P4 mantienen vision 9/9.

## Runner y evidencia

- `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`
- `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1`
- `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`

## Runner de falsos positivos

Archivo creado:

- `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`

Hash:

```text
D552AA137F9A59D59574D2C86F934BCEDD00AF4A0C6785A6FEE67030E4388E17
```

Uso esperado en VM:

```powershell
cd C:\Users\seguridad\Desktop\TFM\03_RUNNERS
.\TFM_Run_FP_Tests_v1.ps1 -ClearJsonlBeforeRun $true -KeepWorkspace $true
```

Salidas por defecto:

- `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\FPs`

Pendiente:

1. Copiar/confirmar el runner en `C:\Users\seguridad\Desktop\TFM\03_RUNNERS`.
2. Activar P1 recalibrado, P2/P3/P4 y SOC_v3.
3. Ejecutar el runner FP.
4. Revisar `_summary.txt`, `_summary.json`, `_summary.csv`, `_vr_hits.txt` y `_vr_hits.csv`.
5. Evaluar falsos positivos por perfil:
   - P1_CRITICAL: FP grave.
   - P2_EVENT: FP forense relevante.
   - P3: FP medio/comportamental.
   - P4: FP leve o visibilidad amplia.

## Documentacion operativa

- `07_DOCS/CHECKLIST_PREVIO_LOGGING_VR.md`
- `07_DOCS/EJECUCION_FINAL_P1_P2_P3_EVENT_Y_RUNNER_V5.md`
- `07_DOCS/MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT.md`
- `07_DOCS/RESUMEN_TECNICO_CAMBIOS_EVENT_FINAL.md`
- `07_DOCS/VALIDACION_LOCAL_EVENT_FINAL.md`

## Cambio arquitectonico vigente

Se mantienen dos familias:

- Backtesting/hunts historicos:
  - artifacts `CLIENT`;
  - `parse_evtx()`;
  - P1_v3/P2_v2/P3_v12.
- Monitorizacion live:
  - artifacts `CLIENT_EVENT`;
  - `watch_evtx()`;
  - P1/P2/P3 EVENT v1.

## Restricciones

- No mover P1/P2/P3 EVENT a validated sin importacion y validacion experimental, salvo orden directa y documentada del autor.
- No modificar receiver, router ni Discord salvo necesidad justificada.
- No usar rutas de laboratorio, nombres de scripts, `DUMB_LAB`, `dump_exfil`, `datos_robados` ni etiquetas TEC como IOC.
- Mantener cobertura 9/9.
- Separar confianza de severidad.
- Sysmon ID 3/11/23/26 enriquecen contexto, no bloquean deteccion fuerte si hay 4104 suficiente.

## Criterio de exito pendiente

1. P1/P2/P3 EVENT cargan sin error en Velociraptor.
2. Client Event Monitoring activo antes de runner v5.
3. Runner v5 finaliza campana 9/9.
4. Runner v5 imprime `TEC-009 EVIDENCE CHECK` con ZIP, copia, SHA256, HTTP 2xx y upload OK.
5. P3 EVENT emite `TEC009_HTTP_ZIP_Upload_4104` si existe 4104 fuerte.
6. TEC-009 sigue detectable aunque no aparezca Sysmon ID 3.
7. TEC-007/TEC-008 detectan por 4104 fuerte aunque falten Sysmon ID 11 o 23/26.
8. Se exportan eventos Velociraptor y se guardan junto a `FINAL_RUN_<timestamp>`.

## Regeneración controlada posterior a la auditoría (2026-06-19)

La tarea abierta pasa a ser la reconstrucción reproducible de las fuentes y
Excel definitivos. No se deben regenerar todavía los libros con los datos
mezclados existentes.

Orden acordado:

1. Crear wrappers/checkers nuevos únicamente en rutas `candidate`.
2. Generar una campaña TEC-001..TEC-009 con JSONL limpio y transferencia
   TEC-009 verificable.
3. Generar una campaña FP independiente con JSONL limpio.
4. Generar benchmark v1.1 válido de baseline, idle y runner; objetivo recomendado:
   tres repeticiones por escenario.
5. Validar automáticamente cobertura, timestamps, contaminación, roles de
   proceso, hashes y trazabilidad.
6. Solo después reconstruir CSV y copias Excel en `REVIEWED`.

Responsabilidad experimental:

- El autor ejecuta en VM/Velociraptor.

## Actualización 2026-06-19 23:32 CEST

Completada la regeneración del Excel de benchmark desde la campaña controlada
`BENCH_20260619_FINAL01`.

- Excel generado y verificado estructuralmente: 9 hojas, 4 gráficos con datos,
  422 muestras y 770 filas de procesos.
- Los 9 runs pasan los controles bloqueantes de cliente.
- Se mantienen 3 `RunnerStillRunningAtEnd` como `WARN` documentado.
- El trabajo pendiente continúa siendo FP: disponer de exports CLIENT_EVENT
  posteriores a REP_03, filtrar por ventana individual y adjudicar
  P2/UNKNOWN antes de regenerar los Excel restantes.
- Codex no atribuye validación experimental; prepara los scripts, revisa las
  salidas y genera los entregables derivados.

## Actualizacion 2026-06-26 13:02 CEST

Completado el Excel final de benchmark de rendimiento.

- Precheck: APTO.
- Excel final:
  `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`.
- Trazabilidad:
  `04_EVIDENCE/Excel_benchmark_26062026`.
- Contexto:
  `00_CONTEXT/CONTEXT_ESTADO_BENCHMARK_26062026.md`.
- 9/9 runs validos, 3 escenarios y 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner: 0.
- `RunnerStillRunningAtEnd=True` queda como WARN no bloqueante.

Tarea abierta actual:

- Revision manual del autor.
- Redactar Capitulo VII con las hojas del Excel benchmark, sin mezclarlo con deteccion/visibilidad/FP/Wazuh.
