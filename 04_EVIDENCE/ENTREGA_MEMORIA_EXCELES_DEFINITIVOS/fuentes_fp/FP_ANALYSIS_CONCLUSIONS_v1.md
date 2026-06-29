# FP_ANALYSIS_CONCLUSIONS_v1

## Alcance

- RunId analizado: `TFM_FP_20260618_144027`.
- Runner: `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`.
- La ejecucion FP completo las pruebas benignas, pero fallo en la fase final del runner v1.
- Fase fallida: lectura de `soc_alerts.jsonl`, filtrado por ventana/RunId y generacion de `_summary.*` / `_vr_hits.*`.
- Codex no valida experimentalmente. La ejecucion fue realizada por el autor en laboratorio.

## Fuentes disponibles

- `TFM_FP_Run_20260618_144027_vr_hits_MANUAL.csv`.
- `TFM_FP_Run_20260618_144027_vr_hits_MANUAL.txt`.
- Agregado comunicado por el autor tras lectura manual del JSONL.
- No esta disponible desde este host la ruta de VM:
  `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\FPs`.

## Resultado de pruebas benignas

| FP_ID | Estado | Interpretacion |
|---|---:|---|
| FP-001 | OK | PowerShell administrativo legitimo |
| FP-002 | OK | CMD legitimo |
| FP-003 | OK | Scheduled Task benigno temporal |
| FP-004 | OK | Registro benigno no persistente |
| FP-005 | OK | Consulta de servicios |
| FP-006 | OK | Discovery legitimo de seguridad |
| FP-007 | OK | ZIP benigno local |
| FP-008 | OK | Borrado benigno controlado |
| FP-009 | SKIPPED | HTTP benigno local no ejecutado; no evaluar |
| FP-010 | OK | Operaciones normales masivas |

## Agregado global comunicado

| Profile | Alertas |
|---|---:|
| P2_EVENT | 16 |
| P3 | 9 |
| P4 | 4 |
| P1_CRITICAL | 0 |
| Total | 29 |

Top `DetectionName` comunicado:

| DetectionName | Alertas |
|---|---:|
| TEC-006 forensic PowerShell 4104 | 8 |
| TEC-006 Security Software Discovery | 8 |
| TEC-002 forensic Sysmon ID 1 | 6 |
| PowerShell suspicious execution - basic profile v2 | 4 |
| TEC-003 forensic Sysmon ID 1 | 2 |
| TEC-002 Windows Command Shell execution | 1 |

## Hits reconstruidos por RunId desde evidencia local

Busqueda por `RunId` reconstruida desde el CSV manual:

| Profile | Hits |
|---|---:|
| P2_EVENT | 6 |
| P4 | 4 |
| P1_CRITICAL | 0 |
| Total | 10 |

Asociacion local por FP:

| FP_ID | Profile | Interpretacion |
|---|---|---|
| FP-001 | P4 | Visibilidad amplia PowerShell; aceptable salvo volumen excesivo |
| FP-002 | P2_EVENT | CMD legitimo genera TEC-002 forense; revisar sensibilidad, no critico |
| FP-003 | P2_EVENT | Scheduled task temporal genera TEC-002/TEC-003; esperado como evidencia forense |
| FP-005 | P4 | Consulta de servicios via PowerShell; visibilidad amplia |
| FP-006 | P4 local / P2 agregado | Discovery legitimo de seguridad; foco relevante TEC-006 |
| FP-008 | P4 | Borrado controlado visible; no P1 |

## Interpretacion por severidad metodologica

- `P1_CRITICAL = 0`: no hay falso positivo critico soportado por los datos disponibles.
- `P2_EVENT = 16` en el agregado: hay falsos positivos o alertas benignas forenses relevantes que deben documentarse.
- `P3 = 9`: comportamiento medio/comportamental; aceptable si el volumen queda controlado.
- `P4 = 4`: visibilidad amplia esperable.
- El foco principal de revision esta en `TEC-006` y `TEC-002`.
- `FP-006` puede explicar `TEC-006` por discovery legitimo de seguridad.
- `FP-002` puede explicar `TEC-002` por uso legitimo de `cmd.exe`.
- `FP-003` puede explicar `TEC-003` por scheduled task temporal.
- `FP-007` no muestra evidencia local de `TEC-009` ni `P1`.
- `FP-008` no muestra evidencia local de `P1`.
- `FP-009` fue `SKIPPED` y no debe incluirse como prueba ejecutada.

## Decision tecnica

- No modificar P1.
- No modificar P2/P3/P4 todavia.
- No modificar `SOC_v3`.
- No modificar runner ofensivo TEC, receiver ni scripts TEC.
- Corregir solo el postproceso del runner FP.
- Repetir campana FP con runner v1.1 para obtener outputs nativos del runner.

## Evidencias generadas

- `TFM_FP_20260618_144027_summary.txt`
- `TFM_FP_20260618_144027_summary.json`
- `TFM_FP_20260618_144027_summary.csv`
- `TFM_FP_20260618_144027_vr_hits.txt`
- `TFM_FP_20260618_144027_vr_hits.csv`

