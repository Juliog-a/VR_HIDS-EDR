# TEST_MATRIX

Actualizado: 2026-07-10 10:26 CEST.

## Separación obligatoria de niveles

| Nivel | Qué demuestra | Qué no demuestra |
|---|---|---|
| Runner | El escenario se ejecutó y produjo evidencia local esperada. | Que Velociraptor detectó. |
| `CLIENT_EVENT` | Artifact de cliente emitió detección o señal. | Que el router persistió o notificó. |
| `SERVER_EVENT` | Enrutamiento, normalización, persistencia o notificación. | Detección original. |
| JSONL | Persistencia externa de salida SOC. | Detección por sí solo. |
| Receiver TEC-009 | Transferencia HTTP controlada. | Detección HIDS por sí solo. |
| Discord | Notificación. | Detección. |

## Perfiles Velociraptor finales

| Perfil | Artifact | Tipo | Estado |
|---|---|---|---|
| P1 | `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` | `CLIENT_EVENT` | FINAL |
| P2 | `Custom.TFM.HIDS.P2.High.Forensic.Event_v1` | `CLIENT_EVENT` | FINAL |
| P3 | `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1` | `CLIENT_EVENT` | FINAL |
| P4 | `Custom.TFM.HIDS.P4.Low.Basic_v2` | `CLIENT_EVENT` | FINAL |
| Router | `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3` | `SERVER_EVENT` | FINAL, no detecta |

## Técnicas TEC

| TEC | MITRE | Criterio resumido |
|---|---|---|
| TEC-001 | T1059.001 | PowerShell 4104/proceso asociado. |
| TEC-002 | T1059.003 | `cmd.exe` controlado y proceso observado. |
| TEC-003 | T1053.005 | Tarea programada temporal creada, ejecutada y retirada. |
| TEC-004 | T1547.001 | Run Key temporal creada, consultada y retirada. |
| TEC-005 | T1569.002 | Servicio temporal creado y eliminado; System 7045 si disponible. |
| TEC-006 | T1518.001 | Discovery de seguridad con 4104/proceso asociado. |
| TEC-007 | T1486 | Cifrado AES de datos dummy; `.aes` como apoyo, no único criterio. |
| TEC-008 | T1485 | Borrado controlado con `FilesBefore > 0`; Sysmon ID 26 solo forense. |
| TEC-009 | T1074.001/T1560.001/T1048.003 | ZIP dummy, HTTP 2xx, `UploadSucceeded=True`, SHA-256 y bytes transferidos. |

## Falsos positivos

Campaña final:

- FP-001 a FP-010 presentes.
- 10/10 OK.
- 0 WARN, 0 FAIL, 0 SKIPPED.
- `VelociraptorHits.Total = 0`.

Criterio:

- P1 crítico en FP sería hallazgo grave.
- P2 requiere adjudicación por posible señal forense dual-use.
- P3/P4 se documentan como visibilidad o posible ruido, sin inflar criticidad.

## Benchmark

| Escenario | Repeticiones | Criterio |
|---|---:|---|
| `BASELINE_NO_VR` | 3 | Sin cliente Velociraptor. |
| `VR_IDLE` | 3 | Cliente Velociraptor observado. |
| `VR_TEC_RUNNER` | 3 | Cliente observado y runner observado. |

Estado final:

- 9/9 runs válidos.
- `SchemaVersion=1.1`.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- `RunnerStillRunningAtEnd=True` es WARN no bloqueante.

## Criterio global

- No reconstruir cifras finales con fuentes mezcladas.
- No usar JSONL/Discord como detección primaria.
- No usar benchmark para calidad de detección.
- No aceptar conclusiones nuevas sin evidencia trazable.
