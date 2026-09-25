# Análisis de la capacidad de Velociraptor para la detección de técnicas MITRE ATT&CK asociadas a ransomware

Repositorio del Trabajo Fin de Máster sobre el uso de **Velociraptor como HIDS/DFIR en Windows**. Evalúa, en un laboratorio controlado, la visibilidad, la detección y el coste operativo de artifacts propios ante un subconjunto de técnicas MITRE ATT&CK relacionadas con la cadena de un ransomware. Incluye pruebas benignas de falsos positivos y una comparación auxiliar con Wazuh.

El proyecto distingue visibilidad de telemetría, detección, alerta en tiempo real, evidencia forense y salida externa. Los artifacts `CLIENT_EVENT` emiten las detecciones; el artifact `SERVER_EVENT` enruta, persiste o notifica. JSONL y Discord son salidas, no detectores. Sysmon ID 26 se usa como evidencia forense, no como alerta individual.

## Laboratorio y técnicas

La arquitectura emplea Windows, Sysmon y el registro PowerShell 4104, un cliente y servidor Velociraptor, scripts de escenarios controlados y un receiver HTTP para TEC-009. Wazuh se analiza por separado como comparación HIDS/SIEM. Los detalles de cada condición de prueba están en la [matriz de pruebas](07_DOCS/TEST_MATRIX.md).

| Caso | Técnica ATT&CK | Escenario |
|---|---|---|
| TEC-001 | T1059.001 | PowerShell |
| TEC-002 | T1059.003 | Windows Command Shell |
| TEC-003 | T1053.005 | Tarea programada |
| TEC-004 | T1547.001 | Run Key |
| TEC-005 | T1569.002 | Ejecución mediante servicio |
| TEC-006 | T1518.001 | Descubrimiento de software de seguridad |
| TEC-007 | T1486 | Cifrado de datos ficticios |
| TEC-008 | T1485 | Borrado controlado |
| TEC-009 | T1074.001 / T1560.001 / T1048.003 | Staging, ZIP y transferencia HTTP controlada |

## Contenido

| Ruta | Función |
|---|---|
| [`01_ARTIFACTS/validated/`](01_ARTIFACTS/validated/) | Cuatro perfiles `CLIENT_EVENT` (P1–P4) y router `SERVER_EVENT` finales. |
| [`02_SCRIPTS/validated/`](02_SCRIPTS/validated/) | Escenarios y receiver de laboratorio. |
| [`03_RUNNERS/`](03_RUNNERS/) | Runners TEC, FP y benchmark; el runner usado por la repetición controlada está en `validate/`. |
| [`06_CONTROLLED_RERUN/`](06_CONTROLLED_RERUN/) | Preflight, preparación, ejecución y verificación de la campaña. |
| [`07_DOCS/`](07_DOCS/) | Guías de configuración, artifacts y cobertura. |
| [`04_EVIDENCE/`](04_EVIDENCE/) | Exportaciones, tablas normalizadas y fuentes de los libros de resultados. |
| [`08_MEMORIA/`](08_MEMORIA/) | Memoria entregada en PDF/DOCX y libros de visibilidad y rendimiento. |
| [`10_WAZUH/`](10_WAZUH/) | Reglas y evidencia de la comparación auxiliar. |

## Reproducción

1. Preparar un laboratorio Windows aislado con Velociraptor, Sysmon y PowerShell 5.1; habilitar la telemetría descrita en la [guía de repetición](06_CONTROLLED_RERUN/README.md). La prueba TEC-009 requiere un receptor HTTP controlado. Wazuh es necesario solo para repetir la comparación auxiliar.
2. Importar los [cinco artifacts finales](01_ARTIFACTS/validated/) y configurar por separado los cuatro `CLIENT_EVENT` y el router `SERVER_EVENT`. Mantener desactivada la notificación Discord salvo que se configure una salida propia.
3. Ejecutar el [preflight](06_CONTROLLED_RERUN/00_Preflight_Controlled_Rerun.ps1) y seguir la secuencia TEC, FP y benchmark de la guía. El [runner TEC canónico](03_RUNNERS/validate/TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1), el [runner FP](03_RUNNERS/validate/TFM_Run_FP_Tests_v1.ps1) y el [runner de benchmark](03_RUNNERS/validate/TFM_Benchmark_VR_Resource_Usage_v1.ps1) están incluidos. Los hashes de los componentes constan en [`CONFIG_SOURCE_MANIFEST.csv`](06_CONTROLLED_RERUN/CONFIG_SOURCE_MANIFEST.csv).
4. Verificar las salidas con [`04_Validate_Controlled_Rerun.ps1`](06_CONTROLLED_RERUN/04_Validate_Controlled_Rerun.ps1) y contrastarlas con la [matriz](07_DOCS/TEST_MATRIX.md). La validación experimental requiere ejecutar la campaña en el laboratorio.

Los [resultados finales](08_MEMORIA/) se presentan en la memoria y en dos libros: visibilidad/detección/FP/Wazuh y rendimiento. Las fuentes tabulares están en [`04_EVIDENCE/Excel_visibilidad_26062026/`](04_EVIDENCE/Excel_visibilidad_26062026/) y [`04_EVIDENCE/Excel_benchmark_26062026/`](04_EVIDENCE/Excel_benchmark_26062026/). El [índice de evidencias](07_DOCS/EVIDENCE_INDEX.md) explica su alcance. Las conclusiones se limitan a la configuración y al laboratorio evaluados.
