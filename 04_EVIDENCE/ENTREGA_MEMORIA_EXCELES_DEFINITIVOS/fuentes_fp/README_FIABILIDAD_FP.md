# README_FIABILIDAD_FP

## Campana FP principal

- RunId: TFM_FP_20260618_151316
- Fuente LOG:
  - TFM_FP_20260618_151316_summary.json
  - TFM_FP_20260618_151316_vr_hits.csv
  - TFM_FP_20260618_151316_vr_hits.txt

## Resultado

- Planificadas: 10
- OK: 9
- SKIPPED: 1
- FAIL: 0
- Hits VR: 15
- P4: 10
- P3: 5
- P2_EVENT: 0
- P1_CRITICAL: 0

## Calidad de asociacion

- EXACT_RUNID: 3
- TIME_WINDOW_INFERRED: 9
- PREVIOUS_RUN_CONTAMINATION: 3
- UNKNOWN: 0

## Interpretacion

- No hay FP critico.
- No hay FP forense.
- Ruido benigno limitado a P3/P4.
- TEC-002 corresponde a uso benigno de cmd.exe.
- TEC-006 corresponde a discovery legitimo de seguridad.

## Nota

No mezclar TFM_FP_20260618_151316 con TFM_FP_20260618_151244 sin mantener AssociationQuality.
