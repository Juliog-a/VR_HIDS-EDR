# PROJECT_STATE

Actualizado: 2026-07-10 10:26 CEST.

## Estado global

TFM sobre evaluación de Velociraptor como HIDS/DFIR en Windows para observar,
detectar y analizar un subconjunto controlado de técnicas MITRE ATT&CK mediante
artifacts personalizados.

Estado actual de la memoria:

- `APTO PARA REVISIÓN FINAL HUMANA`.
- `NO APTO PARA PDF FINAL TODAVÍA`.
- Documento vigente para continuar:
  `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
- Documento original beta que no debe modificarse sin orden expresa:
  `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx`.

## Alcance real

- Sistema evaluado: Velociraptor en laboratorio Windows.
- Comparación auxiliar: Wazuh como HIDS/SIEM, sin equivalencia directa de
  unidades con Velociraptor.
- Técnicas evaluadas: `TEC-001` a `TEC-009`.
- No extrapolar a toda MITRE ATT&CK Enterprise.
- No extrapolar a producción.
- Codex no valida experimentalmente; la validación experimental corresponde al
  autor en laboratorio.

## Técnicas evaluadas

| TEC | MITRE ATT&CK | Escenario resumido |
|---|---|---|
| TEC-001 | T1059.001 PowerShell | Ejecución PowerShell con logging 4104. |
| TEC-002 | T1059.003 Windows Command Shell | Uso controlado de `cmd.exe`. |
| TEC-003 | T1053.005 Scheduled Task | Creación, ejecución y retirada de tarea temporal. |
| TEC-004 | T1547.001 Registry Run Keys | Valor Run temporal controlado. |
| TEC-005 | T1569.002 Service Execution | Servicio temporal controlado. |
| TEC-006 | T1518.001 Security Software Discovery | Discovery benigno de software/servicios de seguridad. |
| TEC-007 | T1486 Data Encrypted for Impact | Cifrado AES de datos dummy controlados. |
| TEC-008 | T1485 Data Destruction | Borrado controlado dentro del dataset dummy. |
| TEC-009 | T1074.001, T1560.001, T1048.003 | Staging, ZIP y exfiltración HTTP controlada en laboratorio. |

## Resultados finales consolidados

Velociraptor:

- Cobertura real documentada: 9/9 técnicas.
- Filas consolidadas `CLIENT_EVENT`: 379.
- Distribución por perfil:
  - P1 = 19.
  - P2 = 118.
  - P3 = 109.
  - P4 = 133.

Falsos positivos:

- FP-001 a FP-010 presentes.
- 10/10 OK.
- 0 WARN, 0 FAIL, 0 SKIPPED.
- `VelociraptorHits.Total = 0`.
- 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas con `vr_hits`; se
  documentan como discrepancia y no se usan como FP definitivo.

Benchmark:

- 9/9 runs válidos.
- Escenarios: `BASELINE_NO_VR`, `VR_IDLE`, `VR_TEC_RUNNER`.
- 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- CPU media en `VR_TEC_RUNNER`: 0,95 %.
- RAM media aproximada del cliente Velociraptor: 57 MB.
- Pico CPU cliente: 54 %.
- `RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER` es WARN no bloqueante.

Wazuh:

- Base: 902 archives, 58 alerts.
- Custom: 18.285 archives, 110 alerts, 56 alertas 110xxx.
- Regla `110201=0` conservada como gap TEC-009.
- Wazuh se usa como comparación HIDS/SIEM, no como sustituto ni unidad
  equivalente a `CLIENT_EVENT`.

TEC-009:

- Se mantiene como exfiltración HTTP controlada en laboratorio.
- Evidencia mínima defendible en la memoria: `ReceiverReachable=True`,
  `HTTPStatus=200`, `UploadSucceeded=True`, SHA-256 local y bytes transferidos
  en `ReceiverResponse`.
- Si en el paquete revisado no aparece `receiver_log.jsonl` o ZIP recibido, se
  trata como limitación documental secundaria, no como invalidación de TEC-009.

## Estado de Velociraptor

Artifacts fuente de verdad:

- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`

Criterio:

- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, normaliza, persiste o notifica.
- JSONL y Discord son salidas externas, no detección primaria.
- `01_ARTIFACTS/candidate` y `01_ARTIFACTS/debug` no deben leerse por defecto.

## Estado de Wazuh

Fuente principal:

- `10_WAZUH/tfm_wazuh_custom_rules_v1.xml`.

Evidencia documental:

- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`.
- `04_EVIDENCE/Excel_visibilidad_26062026`.

Uso metodológico:

- Comparación de capacidades HIDS/SIEM.
- No equivalencia directa entre `archives/alerts` y `CLIENT_EVENT`.
- Gap TEC-009 mantenido en regla `110201=0`.

## Estado de benchmark

Fuente final:

- `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`.

Trazabilidad:

- `04_EVIDENCE/Excel_benchmark_26062026`.
- Fuente canónica regenerada:
  `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS/02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`.

Interpretación:

- Benchmark de coste operativo observado en laboratorio.
- No mide calidad de detección.
- No debe usarse para afirmar rendimiento universal en producción.

## Estado de falsos positivos

Fuente final:

- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`.
- Trazabilidad en `04_EVIDENCE/Excel_visibilidad_26062026`.

Criterio:

- FP definitivo se basa en `summary.json` y `vr_hits.csv` reconciliados.
- Filas fuera de ventana, duplicadas o no reconciliadas se documentan, pero no
  inflan resultados.

## Estado de la memoria

Memoria vigente:

- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.

Estado:

- Contenido técnico y metodológico avanzado.
- No exportar PDF final todavía.
- Bloqueantes antes de PDF:
  - cerrar Word y confirmar desaparición del temporal
    `~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`;
  - eliminar índice general duplicado;
  - actualizar índice general, tablas, figuras y código;
  - reescribir frase del capítulo VI sobre JSONL/Discord;
  - revisión visual de tablas, captions, anexos y PDF provisional.

## Rutas importantes

| Ruta | Uso |
|---|---|
| `00_CONTEXT/README_CONTEXT.md` | Puerta de entrada al contexto limpio. |
| `00_CONTEXT/OLD_CONTEXT_USELESS` | Contexto histórico/no leer por defecto. |
| `01_ARTIFACTS/validated` | Artifacts finales/validados. |
| `01_ARTIFACTS/candidate` | Candidatos históricos o no finales; no leer por defecto. |
| `02_SCRIPTS/validated` | Scripts validados de apoyo. |
| `03_RUNNERS/validate` | Runners/scripts de validación controlada. |
| `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO` | Runners antiguos y backups. |
| `04_EVIDENCE/Excel_visibilidad_26062026` | Evidencia final visibilidad/detección/FP/Wazuh. |
| `04_EVIDENCE/Excel_benchmark_26062026` | Evidencia final benchmark. |
| `06_CONTROLLED_RERUN` | Paquete reproducible de repetición controlada. |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10` | Iteración vigente de memoria pre-PDF. |
| `10_WAZUH` | Material de comparación Wazuh. |
| `99_ARCHIVE` | Histórico ya archivado. |

## Fuentes de verdad

Leer primero:

1. `00_CONTEXT/README_CONTEXT.md`.
2. `00_CONTEXT/PROJECT_STATE.md`.
3. `00_CONTEXT/CURRENT_TASK.md`.
4. `00_CONTEXT/DECISIONS.md`.

Fuentes secundarias:

- `00_CONTEXT/ARTIFACT_INDEX.md`.
- `00_CONTEXT/TEST_MATRIX.md`.
- `00_CONTEXT/EVIDENCE_INDEX.md`.
- `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`.
- `00_CONTEXT/CODEX_RULES.md`.
- `00_CONTEXT/VM_ACTIVE_PATHS.md`.

No leer por defecto:

- `00_CONTEXT/OLD_CONTEXT_USELESS`.
- Cualquier carpeta `candidate`.
- `01_ARTIFACTS/debug`.
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`.
- `99_ARCHIVE`.
