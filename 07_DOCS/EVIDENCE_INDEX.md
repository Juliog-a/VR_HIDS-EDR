# EVIDENCE_INDEX

Actualizado: 2026-07-10 10:26 CEST.

## Fuentes finales principales

| Evidencia | Ruta | Estado |
|---|---|---|
| Visibilidad/detección/FP/Wazuh | `04_EVIDENCE/Excel_visibilidad_26062026` | FINAL |
| Excel visibilidad en memoria | `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx` | FINAL |
| Benchmark | `04_EVIDENCE/Excel_benchmark_26062026` | FINAL |
| Excel benchmark en memoria | `08_MEMORIA/TFM_BENCHMARK_RENDIMIENTO.xlsx` | FINAL |
| Artifacts Velociraptor | `01_ARTIFACTS/validated` | FINAL |
| Reglas Wazuh | `10_WAZUH/tfm_wazuh_custom_rules_v1.xml` | FINAL |
| Memoria entregada | `08_MEMORIA/TFM.docx` | FINAL |

## Resultados Velociraptor

- TEC-001 a TEC-009: 9/9 presentes.
- Filas `CLIENT_EVENT`: 379.
- P1 = 19.
- P2 = 118.
- P3 = 109.
- P4 = 133.

Fuente:

- `04_EVIDENCE/Excel_visibilidad_26062026`.
- `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx`.

## Falsos positivos

- FP-001 a FP-010: 10/10 OK.
- 0 WARN, 0 FAIL, 0 SKIPPED.
- `VelociraptorHits.Total = 0`.
- 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas: documentadas como
  discrepancia, no usadas como FP definitivo.

## Benchmark

- 9/9 runs válidos.
- 3 escenarios y 3 repeticiones por escenario.
- CPU media `VR_TEC_RUNNER`: 0,95 %.
- RAM media cliente: ~57 MB.
- Pico CPU cliente: 54 %.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.

Interpretación:

- Coste operativo observado en laboratorio.
- No calidad de detección.

## Wazuh

- Base: 902 archives, 58 alerts.
- Custom: 18.285 archives, 110 alerts, 56 alertas 110xxx.
- `110201=0` se conserva como gap TEC-009.
- No comparar `archives/alerts` como unidades equivalentes a `CLIENT_EVENT`.

## TEC-009

Formulación vigente:

- Exfiltración HTTP controlada en laboratorio.
- Staging, archivado ZIP y transferencia HTTP controlada con datos dummy.

Evidencia mínima defendible en la memoria:

- `ReceiverReachable=True`.
- `HTTPStatus=200`.
- `UploadSucceeded=True`.
- SHA-256 local.
- Bytes transferidos en `ReceiverResponse`.

Criterio documental:

- La ausencia de `receiver_log.jsonl` o ZIP recibido en el paquete revisado no
  invalida TEC-009.
- Si esos artefactos aparecen en evidencia histórica, se tratan como apoyo
  externo adicional, no como requisito único.

## Evidencia histórica

No leer por defecto:

- `04_EVIDENCE/legacy_reports`.
- `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS`.
- `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS_REVIEWED`.
- `04_EVIDENCE/diag_p3v3`.
- `04_EVIDENCE/TEC009`.
- `99_ARCHIVE`.

Estas rutas se conservan por trazabilidad histórica o diagnóstica.
