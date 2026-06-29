# VALIDACION LOCAL EVENT FINAL

Fecha: 2026-06-15

## Alcance

Validacion local estatica de la iteracion final EVENT:

- Runner v5.
- Script auxiliar de recogida de evidencia.
- Artifacts CLIENT_EVENT P1/P2/P3.
- Documentacion operativa.

Esta validacion no sustituye la prueba final en Velociraptor ni la ejecucion de la campana en laboratorio.

## Resultado

| Elemento | Validacion | Resultado |
|---|---|---|
| TFM_Run_All_TEC_Tests_v5.ps1 | Parser PowerShell | OK |
| TFM_Collect_Final_Evidence_v1.ps1 | Parser PowerShell | OK |
| P1 EVENT | `type: CLIENT_EVENT` | OK |
| P2 EVENT | `type: CLIENT_EVENT` | OK |
| P3 EVENT | `type: CLIENT_EVENT` | OK |
| P1/P2/P3 EVENT | Uso de `watch_evtx()` para live monitoring | OK |
| P1/P2/P3 EVENT | Ausencia de `parse_evtx()` | OK |
| P1/P2/P3 EVENT | Cobertura textual TEC-001 a TEC-009 | OK |
| P1/P2/P3 EVENT | Ausencia de `Generic.Events.TrackNetworkConnections` | OK |
| P1/P2/P3 EVENT | No uso de `DUMB_LAB`, `dump_exfil`, `datos_robados` ni nombre de script TEC-009 como IOC | OK |

## Validaciones no realizadas

| Validacion | Motivo |
|---|---|
| Compilacion VQL con binario Velociraptor | No hay `velociraptor.exe` disponible en el entorno local de Codex. |
| Parse YAML con PyYAML/Ruby | Python/py no arranca en esta sesion y Ruby no esta instalado. |
| Ejecucion real de CLIENT_EVENT | Requiere importar artifacts en Velociraptor y activar Client Event Monitoring en el laboratorio. |
| Confirmacion experimental 9/9 | La validacion final debe hacerla el autor ejecutando runner v5 con los artifacts EVENT activos. |

## Justificacion de `watch_evtx()`

Los artifacts finales son `CLIENT_EVENT`, por lo que su objetivo es monitorizacion en vivo. Para ese diseno, `watch_evtx()` es apropiado porque emite eventos nuevos mientras el artifact esta activo.

Para backtesting historico siguen existiendo las versiones `CLIENT` basadas en `parse_evtx()`:

- `Custom.TFM.HIDS.P1.Low.Basic_v3.yaml`
- `Custom.TFM.HIDS.P2.High.Forensic_v2.yaml`
- `Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml`

## Puntos pendientes en GUI Velociraptor

1. Importar `Custom.TFM.HIDS.P1.Low.Basic.Event_v1`.
2. Importar `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`.
3. Importar `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`.
4. Activarlos en Client Event Monitoring antes de ejecutar runner v5.
5. Ejecutar `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`.
6. Confirmar que P3 emite `TEC009_HTTP_ZIP_Upload_4104` con `Confidence=HIGH` cuando 4104 contiene `Invoke-WebRequest`, `-InFile` y `application/zip`, `.zip`, `upload` o `ReceiverUrl`.
7. Confirmar que Sysmon ID 1/3/11 aparece como contexto y no como condicion obligatoria para TEC-009.
