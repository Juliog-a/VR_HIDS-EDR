\# EVIDENCE\_INDEX



\## TEC-009

Actualización documental 2026-07-08:

\- Para la memoria beta, TEC-009 se formula como exfiltración HTTP controlada en laboratorio.

\- La evidencia mínima defendible en el paquete de memoria es: `ReceiverReachable=True`, `HTTPStatus=200`, `UploadSucceeded=True`, SHA-256 local y bytes transferidos en `ReceiverResponse`.

\- Si en el paquete revisado no aparece `receiver_log.jsonl` o ZIP recibido, no invalida la técnica. Debe tratarse como limitación documental secundaria sobre conservación del receptor final.

Evidencia externa validada:

\- POST /upload recibido.

\- ZIP dummy recibido.

\- bytes\_received registrado.

\- SHA256 registrado.

\- content\_type = application/zip.

\- user\_agent registrado.

\- receiver\_log.jsonl confirma transferencia.



Interpretación:

\- Sí hay transferencia efectiva de ZIP dummy.

\- No hay datos sensibles reales.

\- No hay terceros comprometidos.

\- El receptor pertenece al investigador.



\## Evidencias clave

\- receiver\_log.jsonl

\- upload\_\*.zip

\- alerts.jsonl

\- logs de ejecución

\- eventos 4104

\- eventos Sysmon ID 1/3/11/22

## Revision P3_v8

- Evidencia experimental comunicada por el investigador:
  - 4104 real de `exfiltracion_v3.ps1`.
  - Regex positivas sobre ese 4104: UploadHttp, UploadFile, UploadPost, UploadZipContent y CryptoStrong por SHA256/Get-FileHash.
  - P3_v8 importado y source CU009_HTTP_Upload_Strong_4104 seleccionada en Velociraptor UI.
  - TEC-009 ejecuta correctamente con HTTP 200 y upload correcto, pero P3_v8 no muestra filas en esa source.
- Interpretacion:
  - El hashing SHA256 corresponde a evidencia de integridad.
  - No debe tratarse como cifrado/encryption ni bloquear CU-009.
  - El fallo actual apunta a esquema/campo/captacion de 4104 en Velociraptor, no a receiver ni red.
  - La validacion experimental de P3_v8 no queda cerrada.

## Debug PowerShell 4104 Raw v1

- Artifact creado para comparar:
  - `watch_evtx` live.
  - `parse_evtx` historico/backtest.
  - `ScriptText` normalizado con fallback.
- Resultado experimental comunicado:
  - `WatchEvtx_Raw` no emitio filas relevantes tras TEC-009.
  - `Normalized_Test` no emitio deteccion.
  - Logs mostraron error de ruta: `C:\Windows\System32\Winevt\Logs\Microsoft-Windows-PowerShell\Operational`.
  - `ParseEvtx_Backtest_Raw` si mostro 4104 historicos.
  - Se confirmaron campos `EventData.ScriptBlockText`, `EventData.ScriptBlockId`, `EventData.MessageNumber`, `EventData.MessageTotal`, `Message`, `EventDataJSON` y `SystemJSON`.

## Debug PowerShell 4104 Raw v2

- Artifact creado para separar:
  - `ChannelName = Microsoft-Windows-PowerShell/Operational`.
  - `EvtxPath = C:\Windows\System32\winevt\Logs\Microsoft-Windows-PowerShell%4Operational.evtx`.
- Resultado operativo:
  - Al ser `type: CLIENT_EVENT`, no aparece como artifact ejecutable en Hunt Manager / Collection manual.

## Debug PowerShell 4104 ParseCU009 v1

- Artifact `type: CLIENT` creado para Hunt/Collection manual.
- Usa solo `parse_evtx`.
- Source: `PowerShell4104_ParseEvtx_CU009_Filtered`.
- Validacion pendiente:
  - Ejecutar con `SinceMinutes=240`.
  - Interpretar `CU009StrongCandidate`.

## Debug PowerShell 4104 WatchCU009 v1

- Artifact `type: CLIENT_EVENT` creado para Client Monitoring live.
- Usa solo `watch_evtx`.
- Activar solo si ParseCU009 v1 confirma `CU009StrongCandidate=True`.

## Debug PowerShell 4104 CU009Shape v3

- Resultado experimental previo comunicado:
  - parse_evtx y watch_evtx emiten 4104.
  - Falsos positivos por comandos de validacion `Get-WinEvent`.
  - Falsos positivos por texto de prompt/documentacion con listas de herramientas HTTP.
- Resultado experimental posterior comunicado:
  - El 4104 real de `exfiltracion_v3.ps1` se lee correctamente por `parse_evtx`.
  - `UploadHttpRegex=True`.
  - `UploadFileRegex=True`.
  - `UploadPostRegex=True`.
  - `UploadZipContentRegex=True`.
  - `HashEvidenceRegex=True`.
  - `CryptoEncryptionRegex=False`.
  - `DetectionEngineeringNoiseRegex=False`.
  - `DocumentationNoiseRegex=False`.
  - `ActualInvokeWebRequestShape=False`.
  - `ActualUploadCommandShape=False`.
  - `DestructionStrongRegex=True`.
  - `CU009StrongCandidate_v3=False`.
- Interpretacion:
  - El problema ya no esta en EVTX, ruta, campo `ScriptBlockText` ni receiver.
  - El fallo esta en la logica de forma real del comando de v3.
  - `DestructionStrongRegex=True` corresponde probablemente a limpieza local con `Remove-Item` y no debe bloquear CU-009.
- Artifact `type: CLIENT` creado para backtest manual.
- Source: `PowerShell4104_ParseEvtx_CU009_CommandShape_v3`.
- Logica:
  - exige forma real `Invoke-WebRequest/iwr` con `-Uri`, `-Method POST`, `-InFile` y `-ContentType application/zip`;
  - permite `HashEvidenceRegex`;
  - excluye `CryptoEncryptionRegex`;
  - excluye `DetectionEngineeringNoiseRegex`;
  - excluye `DocumentationNoiseRegex`.
- Validacion estatica local:
  - `exfiltracion_v3.ps1` -> `CU009StrongCandidate_v3=True`.
  - comando de validacion -> `DetectionEngineeringNoiseRegex=True` y candidato fuerte falso.
  - texto documental -> `DocumentationNoiseRegex=True` y candidato fuerte falso.
- Validacion experimental de v3 no valida CU-009; superado por v4.

## Debug PowerShell 4104 CU009Shape v4

- Artifact `type: CLIENT` creado para backtest manual.
- Source: `PowerShell4104_ParseEvtx_CU009_CommandShape_v4`.
- Logica:
  - normaliza `ScriptText` con fallback sobre campos 4104;
  - calcula `ActualInvokeWebRequestShape` por coexistencia de herramienta HTTP y `-Uri`/destino;
  - calcula `ActualUploadCommandShape` por coexistencia en el mismo `ScriptText` de herramienta HTTP, URI/destino, POST, `InFile` y `application/zip`;
  - permite `HashEvidenceRegex`;
  - excluye solo `CryptoEncryptionRegex`;
  - excluye ruido de validacion y documentacion;
  - emite `DestructionStrongRegex`, pero no lo usa como bloqueo final de CU-009.
- Validacion estatica local:
  - `exfiltracion_v3.ps1` -> `ActualInvokeWebRequestShape=True`, `ActualUploadCommandShape=True`, `CU009StrongCandidate_v4=True`.
  - comando de validacion -> `DetectionEngineeringNoiseRegex=True`, candidato fuerte falso.
  - texto documental/lista de herramientas -> `DocumentationNoiseRegex=True`, candidato fuerte falso.
- Validacion experimental comunicada por el investigador:
  - `ActualInvokeWebRequestShape=True`.
  - `ActualUploadCommandShape=True`.
  - `UploadHttpRegex=True`.
  - `UploadFileRegex=True`.
  - `UploadPostRegex=True`.
  - `UploadZipContentRegex=True`.
  - `HashEvidenceRegex=True`.
  - `CryptoEncryptionRegex=False`.
  - `DetectionEngineeringNoiseRegex=False`.
  - `DocumentationNoiseRegex=False`.
  - `CU009StrongCandidate_v4=True`.
- Evidencia local revisada:
  - `05_LOGS/CU009Shape_v4_validacion_4104_TEC009.json`.
  - 20 filas totales.
  - 4 filas fuertes con `CU009StrongCandidate_v4=True`.
  - 16 filas de ruido descartadas por regex de validacion/documentacion.
- Interpretacion:
  - La forma validada de CU-009 debe integrarse en P3_v9.

## P3 Medium Behavioral v9

- Artifact candidate creado:
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v9.yaml`.
- Integra en `CU009_Source_TEC009_HTTP_Upload_Strong_4104` la logica validada de CU009Shape v4:
  - `ScriptText` normalizado;
  - herramienta HTTP;
  - URI/destino;
  - POST;
  - `InFile`;
  - `application/zip`;
  - hashing permitido;
  - cifrado real excluido;
  - ruido de validacion/documentacion excluido;
  - `DestructionStrongRegex` emitido como contexto, no como bloqueo.
- Validacion pendiente:
  - superado por P3_v10 para nueva prueba live.
- Resultado experimental comunicado:
  - importado y activado como `CLIENT_EVENT`;
  - TEC-009 ejecuto correctamente con HTTP 200 y upload correcto;
  - Windows genero 4104 real con `Invoke-WebRequest`, `-Uri`, `-Method POST`, `-InFile` y `application/zip`;
  - ScriptBlockId observado: `73f0278c-4a97-49e8-8590-eb5a5e9fcfb6`;
  - `CU009_Source_TEC009_Local_Archive_Staging_4104_Context` emitio fila;
  - `CU009_Source_TEC009_HTTP_Upload_Strong_4104` no emitio filas;
  - fuentes opcionales Sysmon/DNS/network no emitieron filas.
- Interpretacion:
  - P3_v9 no queda validado.
  - La causa probable esta en la integracion live de CU-009 fuerte, especialmente el wrapper `foreach` sobre `watch_evtx`, no en el script ni en la telemetria 4104.

## P3 Medium Behavioral v10

- Artifact candidate creado:
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v10.yaml`.
- Cambio respecto a v9:
  - `CU009_Source_TEC009_HTTP_Upload_Strong_4104` usa `watch_evtx(filename=PowerShellLog)` directamente.
  - Filtra sobre `EventData.ScriptBlockText`, campo confirmado en el 4104 real.
  - Mantiene la forma validada de Debug v4: HTTP tool + URI + POST + `InFile` + `application/zip`.
  - Mantiene hashing permitido y cifrado real excluido.
  - Mantiene exclusion de ruido de validacion/documentacion.
  - Emite `DestructionStrongRegex` como contexto, pero no lo usa como bloqueo final.
- Validacion estatica local:
  - `exfiltracion_v3.ps1` oficial produce `CU009StrongCandidate=True` con la logica v10.
- Validacion pendiente:
  - importar como `CLIENT_EVENT`;
  - ejecutar TEC-009 una vez;
  - confirmar fila CU-009 fuerte;
  - validar regresion CU-006/CU-007/CU-008.

## Reorganizacion P3_v11 / P2_v1

- Cambio arquitectonico aplicado en candidate:
  - `P3_v11` queda como capa media de comportamiento general y contexto.
  - `P2_v1` queda como artifact especializado para CU-007, CU-008 y CU-009.
  - P1 queda reservado para correlacion critica final.
- `P3_v11`:
  - conserva CU-001 generico y CU-006 discovery;
  - conserva contexto de crypto/file-impact, destruccion, staging/archive y red PowerShell;
  - no emite deteccion fuerte de CU-007/CU-008/CU-009.
- `P2_v1`:
  - CU-007: `CU007_Source_RansomwareLike_Encryption_4104`;
  - CU-008: `CU008_Source_Sabotage_Destruction_4104` y `CU008_Source_Sabotage_Destruction_Sysmon_Process`;
  - CU-008 Sysmon ID 26: `CU008_Source_FileDelete_Sysmon26_Forensic`, evidencia forense no alerta individual;
  - CU-009: `CU009_Source_Staging_ZIP_HTTP_Upload_4104`, con logica CU009Shape_v4;
  - CU-009 Sysmon ID 1/3/11 como proceso fuerte y contexto de red/ZIP.
- Validacion estatica local:
  - regex P2 CU-007 positiva sobre `Scriptransom_v2.ps1`;
  - regex P2 CU-008 positiva sobre `sabotaje.ps1`;
  - regex P2 CU-009 positiva sobre `exfiltracion_v3.ps1`;
  - no se detectaron en los YAML nuevos rutas de laboratorio, nombres de scripts, endpoint fijo, usuario ni hostname como IOC.
- Validacion pendiente:
  - importar `P3_v11` y `P2_v1`;
  - activar como `CLIENT_EVENT` candidate;
  - ejecutar campana TEC-007/TEC-008/TEC-009;
  - validar regresion CU-001/CU-006;
  - no mover a `validated` sin campana completa.

## Runner TFM_Run_All_TEC_Tests_v4

- Runner creado:
  - `03_RUNNERS/TFM_Run_All_TEC_Tests_v4.ps1`.
  - Copia para VM: `Carpeta_Compartida_TFM/runner_candidate/TFM_Run_All_TEC_Tests_v4.ps1`.
- Cambio principal:
  - `BasePath` actualizado a `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas`.
  - TEC-009 ya no usa staging inline antiguo.
  - TEC-009 ejecuta `exfiltracion_v3.ps1` desde `TEC-009_Exfiltracion`.
- Parametros configurables:
  - `ReceiverUrl`, default `http://192.168.1.129:8088/upload`.
  - `ExfilTimeoutSec`, default `60`.
  - `KeepExfilArtifacts`, default `$true`.
  - `EnableExfilUpload`, default `$true`.
- Evidencia esperada en log:
  - inicio TEC-009;
  - `ReceiverUrl`;
  - `runStart`;
  - ruta de ejecucion;
  - exit code de `exfiltracion_v3.ps1`;
  - HTTP status;
  - Upload succeeded;
  - eventos 4104 coincidentes;
  - restauracion final.
- Validacion estatica:
  - parser PowerShell OK.
  - hash identico entre `03_RUNNERS` y `runner_candidate`.
- Validacion experimental pendiente:
  - ejecutar en VM como administrador con receiver activo;
  - confirmar HTTP 200 y `Upload succeeded True`;
  - confirmar 4104 con `Invoke-WebRequest`, `-Method POST`, `-InFile` y `application/zip`.

## Revision runner v4 20260612_1215

- Evidencia revisada:
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_115408.log`.
- Hallazgos:
  - `TEC-009 receiver reachable => False`.
  - `powershell.exe -File ".\exfiltracion_v3.ps1"` fallo con parametro `-File` porque la ruta relativa no se resolvio de forma fiable.
  - No hubo eventos 4104 coincidentes porque `exfiltracion_v3.ps1` no llego a ejecutarse.
- Correccion:
  - TEC-009 invoca ahora `exfiltracion_v3.ps1` mediante ruta absoluta.
  - El precheck del receiver se conserva y queda incluido en el summary.
  - El runner genera:
    - log completo;
    - summary TXT;
    - summary JSON;
    - summary CSV.
- Hashes finales:
  - `TFM_Run_All_TEC_Tests_v4.ps1`: `E762A2451AD39E5E5136FA1530D635D289B276C124DC775B9AD853CC994704E2`.
  - `TFM_Check_Sysmon_Visibility.ps1`: `82C50549F79D8C102D4086934ED1C321A718A6726CC21590E5911E997136D2D5`.

## TFM_Check_Sysmon_Visibility

- Script creado:
  - `03_RUNNERS/TFM_Check_Sysmon_Visibility.ps1`.
  - `Carpeta_Compartida_TFM/runner_candidate/TFM_Check_Sysmon_Visibility.ps1`.
- Tipo:
  - Solo consulta.
  - No ejecuta tecnicas.
  - No modifica laboratorio.
- Fuentes consultadas:
  - Sysmon Operational.
  - PowerShell Operational.
  - Security.
  - System.
  - Application.
- Salidas:
  - `TFM_Sysmon_Visibility_YYYYMMDD_HHMMSS.log`.
  - `TFM_Sysmon_Visibility_YYYYMMDD_HHMMSS_summary.txt`.
  - `TFM_Sysmon_Visibility_YYYYMMDD_HHMMSS_summary.json`.
  - `TFM_Sysmon_Visibility_YYYYMMDD_HHMMSS_events.csv`.
- Validacion estatica:
  - Parser PowerShell OK.
  - Copia identica entre `03_RUNNERS` y `runner_candidate`.

## Ultima iteracion campana v4 20260612_1258

- Evidencia revisada:
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_123701.log`.
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_123701_summary.txt`.
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_123701_summary.csv`.
  - `05_LOGS/TFM_Sysmon_Visibility_20260612_124318_summary.txt`.
- Resultado observado:
  - Global `FAIL`.
  - OK/WARN/FAIL: `6/1/2`.
  - TEC-007 `FAIL` por `.aesCount=0`.
  - TEC-008 `WARN` por `FilesBefore=0`.
  - TEC-009 `FAIL` por transformacion de booleano en `-EnableUpload:$true`.
  - Receiver TEC-009 alcanzable: `ReceiverReachable=True`.
- Correcciones preparadas:
  - runner v4 con `TFM_BASEPATH`, switches para TEC-009 y `QueryError` 4104.
  - scripts TEC-007 corregidos para dataset activo, enumeracion coherente y AES real sobre dummy.
  - restore TEC-008 corregido para poblar `DUMB_LAB`.
  - `exfiltracion_v3.ps1` con switches.
- Validacion realizada:
  - Parser PowerShell OK en entregables.
  - Hashes SHA256 registrados en `ultima_iteracion/hashes/SHA256SUMS.txt`.
- Validacion pendiente:
  - ejecutar pruebas unitarias TEC-007 y TEC-009 en VM.
  - ejecutar campana completa v4 en VM con receiver activo.

## Campaña TEC_20260619_FINAL01

Estado: `NO APTA` como campaña canónica final.

Evidencia existente comunicada por el autor:

- TEC-001..TEC-009 ejecutadas.
- Receiver TEC-009 alcanzable.
- HTTP 200.
- `UploadSucceeded=True`.
- ZIP origen:
  `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion\staging_controlled\tec009_controlled_archive.zip`.
- SHA-256 origen:
  `e50dff5188761c1b94ab0b94a7eb4a3ddec1715acc9c6bc854e0e942ab5e6f84`.
- PowerShell 4104 comunicado: 49 eventos.

Limitación bloqueante:

- `ArgumentException: Los tipos de argumentos no coinciden` durante conversión
  de listas en Windows PowerShell 5.1.
- No existe summary global final.
- La transferencia se conserva como evidencia diagnóstica, no como resultado
  final de campaña.

## Memoria beta con referencias y anexos 20260708_1250

Documento:

- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx`

Hash SHA-256:

- `DFB82ECF760E56C9F8E8D78485CF7D9571DCB6C6254D6281C9BF3BB95D9025C3`

Evidencia documental:

- Referencias APA insertadas.
- Anexos A-L desarrollados.
- Word abrio, actualizo indices/listas y guardo correctamente.
- Validacion OpenXML correcta.

Fuentes internas citadas:

- `04_EVIDENCE/Excel_visibilidad_26062026/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`
- `04_EVIDENCE/Excel_benchmark_26062026/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`
- `01_ARTIFACTS/validated`
- `10_WAZUH/tfm_wazuh_custom_rules_v1.xml`
- `00_CONTEXT/DECISIONS.md`, `TEST_MATRIX.md`, `EVIDENCE_INDEX.md`
