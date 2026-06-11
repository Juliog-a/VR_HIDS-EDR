\# EVIDENCE\_INDEX



\## TEC-009

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
- Validacion experimental pendiente en VM.
