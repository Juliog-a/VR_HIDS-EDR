# CURRENT_TASK

## Tarea activa

Importar y validar en Velociraptor los artifacts historicos nuevos:

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic_v3.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic_v2.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml`
- `01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.PowerShell4104.LastHours.yaml`
- `01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours.yaml`

Paquete:

- `ultima_iteracion_artifacts`

Antes de validar artifacts, si hace falta repetir campana completa:

`03_RUNNERS/TFM_Run_All_TEC_Tests_v4.ps1`

Antes de la siguiente campana, copiar a la VM el paquete:

`C:\Users\julio\Desktop\TFM\ultima_iteracion`

Objetivo inmediato:

- Ejecutar TEC-001 a TEC-009 desde `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas`.
- Confirmar que TEC-009 ejecuta `TEC-009_Exfiltracion\exfiltracion_v3.ps1`.
- Confirmar que TEC-009 genera ZIP, copia local, SHA256, HTTP status 200, `Upload succeeded True` y 4104 con `Invoke-WebRequest`, `-Method POST`, `-InFile`, `application/zip`.
- Usar la campana para validar P1_v3/P2_v2/P3_v12 como backtesting historico.
- Ejecutar `TFM_Check_Sysmon_Visibility.ps1` antes o despues de la campana para documentar fuentes de telemetria disponibles.

## Contexto de artifacts para validar con el runner

- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic_v3.yaml
- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic_v2.yaml
- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml
- 01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.PowerShell4104.LastHours.yaml
- 01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours.yaml

## Artifacts nuevos creados

- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic_v3.yaml
- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic_v2.yaml
- 01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml
- 01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.PowerShell4104.LastHours.yaml
- 01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours.yaml

## Runner nuevo creado

- 03_RUNNERS/TFM_Run_All_TEC_Tests_v4.ps1
- Carpeta_Compartida_TFM/runner_candidate/TFM_Run_All_TEC_Tests_v4.ps1

## Script de visibilidad creado

- 03_RUNNERS/TFM_Check_Sysmon_Visibility.ps1
- Carpeta_Compartida_TFM/runner_candidate/TFM_Check_Sysmon_Visibility.ps1

## Paquete ultima iteracion

- `ultima_iteracion/runners/TFM_Run_All_TEC_Tests_v4.ps1`
- `ultima_iteracion/checks/TFM_Check_Sysmon_Visibility.ps1`
- `ultima_iteracion/scripts_modificados/TEC-007_Ransomware/crear_directorio_dummy.ps1`
- `ultima_iteracion/scripts_modificados/TEC-007_Ransomware/enumeracion.ps1`
- `ultima_iteracion/scripts_modificados/TEC-007_Ransomware/Scriptransom.ps1`
- `ultima_iteracion/scripts_modificados/TEC-008_Sabotaje/crear_directorio_dummy.ps1`
- `ultima_iteracion/scripts_modificados/TEC-009_Exfiltracion/exfiltracion_v3.ps1`
- `ultima_iteracion/hashes/SHA256SUMS.txt`
- `ultima_iteracion/hashes/PARSER_VALIDATION.txt`

## Restricciones

- No mover P1_v3/P2_v2/P3_v12 a validated sin importacion y validacion experimental en Velociraptor.

- No tocar router.

- No activar Discord.

- Mantener cobertura 9/9 en P1/P2/P3.

- No usar rutas/scripts/lab markers como IOC.

## Resultado ya confirmado

`Custom.TFM.Debug.PowerShell4104.CU009Shape_v4` fue validado experimentalmente por el investigador sobre 4104 real de TEC-009:

- `UploadHttpRegex=True`
- `UploadFileRegex=True`
- `UploadPostRegex=True`
- `UploadZipContentRegex=True`
- `HashEvidenceRegex=True`
- `CryptoEncryptionRegex=False`
- `DetectionEngineeringNoiseRegex=False`
- `DocumentationNoiseRegex=False`
- `ActualInvokeWebRequestShape=True`
- `ActualUploadCommandShape=True`
- `CU009StrongCandidate_v4=True`

## Resultado P3_v9

P3_v9 fue importado y activado como `CLIENT_EVENT`, pero no queda validado:

- TEC-009 ejecuto correctamente.
- HTTP status 200.
- Upload succeeded `True`.
- Windows genero 4104 real con `Invoke-WebRequest`, `-Uri`, `-Method POST`, `-InFile` y `application/zip`.
- ScriptBlockId observado: `73f0278c-4a97-49e8-8590-eb5a5e9fcfb6`.
- `CU009_Source_TEC009_Local_Archive_Staging_4104_Context` emitio fila.
- `CU009_Source_TEC009_HTTP_Upload_Strong_4104` no emitio filas.

Causa probable: integracion de CU-009 fuerte en P3_v9 con `foreach` envolviendo `watch_evtx` live, no equivalente operativamente al backtest `parse_evtx` de Debug v4.

## Cambio arquitectonico vigente

P1/P2/P3 nuevos son artifacts historicos `CLIENT` con `parse_evtx()`.

P3_v12 emite salida normalizada y alertable para 9/9 tecnicas.

P2_v2 mantiene detalle forense y contexto ampliado.

P1_v3 mantiene logica simple y estable.

## Criterio de exito pendiente

1. Copiar archivos desde `ultima_iteracion` a la VM.
2. Verificar que `exfiltracion_v3.ps1` existe en `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion`.
3. Ejecutar prueba unitaria TEC-007 y confirmar `.aesCount > 0`.
4. Ejecutar prueba unitaria TEC-009 con receiver activo y confirmar `Upload succeeded True`.
5. Ejecutar checker de visibilidad:
   - `.\TFM_Check_Sysmon_Visibility.ps1 -LookbackHours 24 -IncludeSamples -MaxSamplesPerQuery 10`
6. Levantar receiver local controlado.
7. Ejecutar runner v4 como administrador.
8. Confirmar para TEC-007 en el log y summary:
   - `.aesCount > 0`.
9. Confirmar para TEC-008 en el log y summary:
   - `FilesBefore > 0`.
   - `FilesAfter < FilesBefore`.
10. Confirmar para TEC-009 en el log y summary:
   - ZIP creado.
   - ZIP copy creado.
   - ZIP SHA256.
   - HTTP status 200.
   - Upload succeeded True.
   - eventos 4104 coincidentes.
11. Confirmar que el runner genera:
   - summary TXT;
   - summary JSON;
   - summary CSV.
12. Importar debug RAW 4104 y Sysmon.
13. Importar P1_v3/P2_v2/P3_v12.
14. Ejecutar como Hunt/Collection manual con `LookbackHours=2`.
15. Revisar TEC-001 a TEC-009 sobre la campana historica.
16. Confirmar en P2/P3:
   - TEC-007 por fuente 4104 fuerte de crypto/AES.
   - TEC-008 por fuente 4104 fuerte de borrado.
   - TEC-009 por fuente 4104 fuerte de HTTP ZIP upload.
17. Confirmar para CU-009:
   - `ActualInvokeWebRequestShape=True`
   - `ActualUploadCommandShape=True`
   - `CryptoEncryptionRegex=False`
   - `DetectionEngineeringNoiseRegex=False`
   - `DocumentationNoiseRegex=False`
   - `CU009StrongCandidate=True`
18. Validar que P3_v12 emite al menos una fila TEC-009 si existe 4104 fuerte, aunque no aparezca Sysmon ID 3.
19. Validar regresion TEC-001 a TEC-006.
