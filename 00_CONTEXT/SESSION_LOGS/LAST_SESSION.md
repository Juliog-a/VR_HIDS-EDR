# LAST_SESSION

## Resumen

Se recibio resultado experimental de `Custom.TFM.Debug.PowerShell4104.CU009Shape_v3`: el artifact lee el 4104 real de `exfiltracion_v3.ps1`, pero `ActualInvokeWebRequestShape=False`, `ActualUploadCommandShape=False` y `CU009StrongCandidate_v3=False`. Se creo `Custom.TFM.Debug.PowerShell4104.CU009Shape_v4` para corregir la logica de forma real del comando.

## Estado actual

- Fuente oficial: `C:\Users\julio\Desktop\TFM`.
- `P3_v8`: intacto; no se modifica.
- `P3_v9`: no creado.
- Debug actual: `01_ARTIFACTS/debug/Custom.TFM.Debug.PowerShell4104.CU009Shape_v4.yaml`.
- Router JSONL: no tocado.
- Discord: desactivado; no tocado.
- P2/P1: desactivados; no tocados.

## Diagnostico actual

- La visibilidad 4104 funciona.
- `EventData.ScriptBlockText` existe.
- `parse_evtx` funciona.
- El receiver y la transferencia HTTP no son el problema actual.
- El fallo de v3 esta en la forma real de comando: la regex ordenada no marca el bloque multilinea con backticks de `Invoke-WebRequest`.
- `DestructionStrongRegex=True` puede aparecer por `Remove-Item` de limpieza local; no debe bloquear CU-009 si el upload fuerte esta presente.

## Cambios clave

- `CU009Shape_v4` es `type: CLIENT`.
- Source: `PowerShell4104_ParseEvtx_CU009_CommandShape_v4`.
- Usa `parse_evtx` sobre `C:\Windows\System32\winevt\Logs\Microsoft-Windows-PowerShell%4Operational.evtx`.
- `ActualInvokeWebRequestShape` se calcula por coexistencia de herramienta HTTP y URI/destino.
- `ActualUploadCommandShape` se calcula por coexistencia de herramienta HTTP, URI/destino, POST, `InFile` y `application/zip`.
- `CU009StrongCandidate_v4` no bloquea por `DestructionStrongRegex`.
- Se mantienen filtros contra ruido de validacion y documentacion.

## Validacion realizada

- Prueba estatica local contra `exfiltracion_v3.ps1`:
  - `ActualInvokeWebRequestShape=True`;
  - `ActualUploadCommandShape=True`;
  - `CU009StrongCandidate_v4=True`.
- Prueba estatica local de ruido:
  - comando `Get-WinEvent` de validacion -> candidato fuerte falso;
  - texto documental/lista de herramientas -> candidato fuerte falso.
- No se ha realizado validacion experimental de v4.

## Pendiente inmediato

1. Importar `Custom.TFM.Debug.PowerShell4104.CU009Shape_v4`.
2. Ejecutarlo como Hunt/Collection manual con `SinceMinutes=240`.
3. Confirmar si el 4104 real de `exfiltracion_v3.ps1` muestra `CU009StrongCandidate_v4=True`.
4. Crear `P3_v9` solo si v4 confirma la forma real del comando en laboratorio.

## Cierre de trazabilidad

- Cierre adicional registrado en `00_CONTEXT/SESSION_LOGS/session_20260611_1909.md`.
- Validacion experimental sigue pendiente por el autor.
