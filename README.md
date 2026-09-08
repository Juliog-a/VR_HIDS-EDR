# VR_HIDS-EDR

Repositorio del Trabajo Fin de Máster que evalúa **Velociraptor como HIDS/DFIR en Windows** mediante artifacts personalizados, técnicas MITRE ATT&CK, evidencia experimental, análisis de falsos positivos, benchmark de rendimiento y una comparación auxiliar con Wazuh.

El proyecto distingue expresamente cinco niveles: visibilidad de telemetría, detección emitida por `CLIENT_EVENT`, alerta en tiempo real, evidencia forense y salida externa. Los artifacts `SERVER_EVENT` enrutan o persisten eventos; JSONL y Discord no constituyen detección.

## Técnicas evaluadas

| Caso | MITRE ATT&CK | Escenario controlado |
|---|---|---|
| TEC-001 | T1059.001 | PowerShell |
| TEC-002 | T1059.003 | Windows Command Shell |
| TEC-003 | T1053.005 | Scheduled Task |
| TEC-004 | T1547.001 | Registry Run Keys / Startup Folder |
| TEC-005 | T1569.002 | Service Execution |
| TEC-006 | T1518.001 | Security Software Discovery |
| TEC-007 | T1486 | Data Encrypted for Impact sobre datos ficticios |
| TEC-008 | T1485 | Data Destruction controlada |
| TEC-009 | T1074.001 / T1560.001 / T1048.003 | Staging, ZIP y transferencia HTTP controlada |

La matriz de pruebas y los criterios de evidencia están en [`00_CONTEXT/TEST_MATRIX.md`](00_CONTEXT/TEST_MATRIX.md). Sysmon ID 26 se usa únicamente como evidencia forense, no como alerta individual.

## Arquitectura experimental

```text
Windows + Sysmon + PowerShell 4104
  -> Velociraptor CLIENT_EVENT (detección)
  -> Velociraptor SERVER_EVENT (enrutado/persistencia)
  -> JSONL o notificación auxiliar
  -> análisis, evidencias y memoria
```

Los perfiles P1, P2, P3 y P4 separan señales críticas, detalle forense, comportamiento y visibilidad básica. Los cinco artifacts finales se encuentran en [`01_ARTIFACTS/validated`](01_ARTIFACTS/validated).

## Estructura

| Ruta | Contenido |
|---|---|
| `00_CONTEXT/` | Estado, decisiones, matrices e índices de trazabilidad. |
| `01_ARTIFACTS/validated/` | Artifacts Velociraptor finales. |
| `02_SCRIPTS/validated/` | Scripts auxiliares validados para las pruebas controladas. |
| `03_RUNNERS/` | Campaña TEC, captura de evidencia y benchmark. |
| `04_EVIDENCE/` | Evidencias experimentales y datasets analíticos. |
| `06_CONTROLLED_RERUN/` | Paquete de repetición controlada y validaciones estáticas. |
| `07_DOCS/` | Guías operativas y documentación técnica. |
| `08_MEMORIA/ENTREGA_TFM_Julgarrei/` | Entregables finales del TFM. |
| `10_WAZUH/` | Evidencia y documentación de la comparación con Wazuh. |

`lab/`, `Carpeta_Compartida_TFM/` y `05_LOGS/` son áreas locales de laboratorio excluidas de Git.

## Scripts principales y reproducción

- `03_RUNNERS/TFM_Run_All_TEC_Tests_v6.ps1`: campaña TEC-001 a TEC-009.
- `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1`: recopilación de evidencia final.
- `03_RUNNERS/validate/TFM_Run_FP_Tests_v1.ps1`: pruebas controladas de falsos positivos.
- `03_RUNNERS/validate/TFM_Benchmark_VR_Resource_Usage_v1.ps1`: captura del benchmark.
- `06_CONTROLLED_RERUN/00_Preflight_Controlled_Rerun.ps1`: comprobaciones previas.
- `06_CONTROLLED_RERUN/01_Run_TEC_Canonical.ps1`: ejecución canónica de las técnicas.
- `06_CONTROLLED_RERUN/04_Validate_Controlled_Rerun.ps1`: validación del rerun.
- `06_CONTROLLED_RERUN/99_Static_SelfTest.ps1`: autocomprobación estática del paquete.

La secuencia, prerrequisitos, separación host/VM y controles de seguridad están documentados en [`06_CONTROLLED_RERUN/README.md`](06_CONTROLLED_RERUN/README.md). Los scripts ofensivos simulados deben ejecutarse exclusivamente en el laboratorio controlado descrito por el proyecto.

## Resultados y comparación

- El Excel de visibilidad reúne dashboard, técnicas, resultados, artifacts, falsos positivos, fuentes de evidencia y comparación con Wazuh.
- El Excel de rendimiento conserva runs válidos, muestras, calidad de datos, gráficas, trazabilidad, limitaciones y discrepancias.
- La comparación con Wazuh es auxiliar y no sustituye la evaluación principal de Velociraptor.
- Las métricas describen el laboratorio evaluado y no representan rendimiento universal.

No se reproducen cifras en este README: los resultados verificables están en los Excel finales y sus evidencias de origen.

## Entregables finales

La carpeta [`08_MEMORIA/ENTREGA_TFM_Julgarrei`](08_MEMORIA/ENTREGA_TFM_Julgarrei) contiene:

- `TFM.docx`: memoria final.
- `DEFENSA_TFM_Julio_Garcia_Amorena.pptx`: presentación de defensa.
- `Analisis_Tecnicas_TFM_Velociraptor.xlsx`: análisis final de visibilidad, detección y comparación.
- `TFM_BENCHMARK_RENDIMIENTO.xlsx`: benchmark final de rendimiento.

El inventario de evidencias se mantiene en [`00_CONTEXT/EVIDENCE_INDEX.md`](00_CONTEXT/EVIDENCE_INDEX.md) y el informe de limpieza final en [`CLEANUP_REPORT.md`](CLEANUP_REPORT.md).
