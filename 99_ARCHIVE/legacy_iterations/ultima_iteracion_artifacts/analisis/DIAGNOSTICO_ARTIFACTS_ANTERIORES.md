# Diagnostico de artifacts anteriores

## Alcance revisado

- Contexto obligatorio en `00_CONTEXT`.
- Artifacts actuales en `01_ARTIFACTS`.
- Evidencia de campana v4 en `05_LOGS`.
- Excel de referencia en `04_EVIDENCE`.
- Debug validado `CU009Shape_v4`.

## Hallazgos principales

1. `P2_v1`, `P3_v11` y `P4_v2` son `type: CLIENT_EVENT` y usan `watch_evtx()`.
   - Sirven para monitorizacion live.
   - No sirven para backtesting de una campana ya ejecutada.
   - La iteracion actual requiere `type: CLIENT` y `parse_evtx()`.

2. `P3_v11` degrada TEC-007, TEC-008 y TEC-009 a contexto.
   - Esto contradice el nuevo requisito de cobertura 9/9.
   - TEC-007/008/009 deben poder emitir HIGH por PowerShell 4104 fuerte.

3. TEC-009 fallaba por mezcla de modelo live y modelo historico.
   - `Debug PowerShell4104 CU009Shape_v4` si valido la forma real por `parse_evtx()`.
   - `P3_v9` no emitio la fuente fuerte live aunque el 4104 real existia.
   - Causa probable documentada: integracion live con `watch_evtx()`/wrapper y ventana temporal.

4. Sysmon ID 3 no debe ser condicion obligatoria para TEC-009.
   - La exfiltracion HTTP controlada dura poco.
   - La red puede perderse por polling, ventana o fuente no activa.
   - La campana v4 confirma 4104 fuerte y upload HTTP correcto.

5. Sysmon ID 23/26 no debe ser condicion obligatoria para TEC-008.
   - La visibilidad de `2026-06-12 13:26` muestra Sysmon ID 23/26 con count 0.
   - Debe quedar como evidencia forense si existe.

6. Sysmon ID 11 no debe ser condicion unica para TEC-007.
   - TEC-007 genera `.aesCount=50`.
   - La fuente principal estable es PowerShell 4104 con AES/CryptoStream/CreateEncryptor.

## Campos reales inferidos

Desde `EVIDENCE_INDEX.md`, debug CU009Shape v4 y CSV/logs:

- PowerShell 4104:
  - `EventData.ScriptBlockText`
  - `EventData.ScriptBlockId`
  - `EventData.MessageNumber`
  - `EventData.MessageTotal`
  - `EventData.Path`
  - `Message`
  - `System.TimeCreated.SystemTime`
  - `System.EventID.Value`
  - `System.Provider.Name`
  - `System.Channel`
  - `System.Computer`

- Sysmon:
  - `EventData.CommandLine`
  - `EventData.Image`
  - `EventData.ParentImage`
  - `EventData.ParentCommandLine`
  - `EventData.TargetFilename`
  - `EventData.DestinationIp`
  - `EventData.DestinationPort`
  - `EventData.DestinationHostname`
  - `EventData.Details`
  - `EventData.Hashes`
  - `EventData.ProcessGuid`
  - `EventData.ProcessId`
  - `EventData.ParentProcessId`

## Cambio aplicado

- P1/P2/P3 nuevos son `type: CLIENT`.
- Todas las fuentes usan `parse_evtx()`.
- `LookbackHours` acota la ventana.
- TEC-009 HIGH depende de 4104 fuerte, no de Sysmon ID 3.
- TEC-008 HIGH depende de 4104 fuerte, no de Sysmon ID 23/26.
- TEC-007 HIGH depende de 4104 fuerte, no solo de Sysmon ID 11.
- Sysmon ID 3/11/23/26 queda como contexto o evidencia forense.

## Diferencia parse_evtx() vs watch_evtx()

- `parse_evtx()`:
  - lectura historica de ficheros EVTX;
  - util para Hunt/Collection manual;
  - permite revisar campanas ya ejecutadas.

- `watch_evtx()`:
  - monitorizacion live;
  - solo ve eventos mientras el artifact esta activo;
  - no recupera eventos pasados.

## Limitacion de Generic.Events.TrackNetworkConnections

- Es util como contexto de red, no como detector principal de exfiltracion.
- Puede perder conexiones breves.
- No prueba por si solo que se haya transferido un ZIP.
- Para TEC-009, la evidencia fuerte es la forma de comando 4104 y el receiver/log si se revisa externamente.

## Conflicto Excel vs campana v4

El Excel `Analisis_Tecnicas_TFM_Velociraptor_DEF.xlsx` define TEC-009 como staging/archive host-based, y versiones previas describen copia local o TCP contextual.

La campana v4 actual confirma una evolucion experimental:

- `ReceiverReachable=True`
- `HTTPStatus=200`
- `UploadSucceeded=True`
- `ZipCount=6`
- `PS4104=1`
- 4104 con `Invoke-WebRequest`, `-Method POST`, `-InFile`, `application/zip`

Se mantiene el mapeo MITRE base del Excel (`T1074.001 / T1560.001`) y se documenta `T1048.003` como contexto de exfiltracion HTTP controlada en la campana v4.
