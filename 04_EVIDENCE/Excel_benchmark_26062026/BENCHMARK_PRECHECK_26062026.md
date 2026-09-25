# BENCHMARK_PRECHECK_26062026

Fecha generacion: 2026-06-26T10:50:53.824Z

Decision: **APTO**

## Fuente canonica

- <RAIZ_TFM>\04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_REGENERADOS\02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx
- Analisis_Benchmark.xlsx queda limitado a referencia historica/formato.

## Validaciones obligatorias

| Check | Status | Evidence | Severity | Decision |
| --- | --- | --- | --- | --- |
| Fuente canonica XLSX disponible | OK | <RAIZ_TFM>\04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_REGENERADOS\02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx | INFO | Debe existir; no se usa Analisis_Benchmark.xlsx como resultado definitivo. |
| SchemaVersion 1.1 presente | OK | 1.1 | CRITICA | Todos los runs validos deben estar en schema 1.1. |
| Exactamente 3 escenarios | OK | {"BASELINE_NO_VR":3,"VR_IDLE":3,"VR_TEC_RUNNER":3} | CRITICA | Escenarios esperados: BASELINE_NO_VR, VR_IDLE, VR_TEC_RUNNER. |
| 3 repeticiones por escenario | OK | {"BASELINE_NO_VR":3,"VR_IDLE":3,"VR_TEC_RUNNER":3} | CRITICA | Cada escenario debe tener REP_01, REP_02 y REP_03. |
| 9/9 runs validos | OK | 9/9 | CRITICA | No se genera Excel final si faltan runs validos. |
| Muestras CPU suficientes | OK | samples.csv combinados=422 | CRITICA | CPU/RAM se interpretan como metricas de laboratorio, no universales. |
| Muestras RAM suficientes | OK | samples.csv combinados=422 | CRITICA | Se exige muestra observada por run. |
| BASELINE_NO_VR presente | OK | 3 runs | CRITICA | Referencia sin cliente VR activo. |
| VR_IDLE presente | OK | 3 runs | CRITICA | Cliente VR activo en reposo. |
| VR_TEC_RUNNER presente | OK | 3 runs | CRITICA | Carga controlada de runner TEC. |
| SERVER_GUI excluido del calculo principal | OK | SERVER_GUI rows en process_samples=422 | CRITICA | SERVER_GUI puede estar observado, pero no suma en metricas de cliente. |
| notepad.exe runner = 0 | OK | notepad rows=0 | CRITICA | Evita contaminacion del runner por notepad.exe. |
| RunnerStillRunningAtEnd evaluado | WARN | true=3; false=6 | NO_BLOQUEANTE | Si aparece en VR_TEC_RUNNER se trata como WARN no bloqueante. |
| ClientRows > 0 cuando aplica | OK | TFM_BENCH_BASELINE_NO_VR_20260619_130840:0; TFM_BENCH_BASELINE_NO_VR_20260619_132157:0; TFM_BENCH_BASELINE_NO_VR_20260619_133538:0; TFM_BENCH_VR_IDLE_20260619_131216:49; TFM_BENCH_VR_IDLE_20260619_132534:49; TFM_BENCH_VR_IDLE_20260619_133913:47; TFM_BENCH_VR_TEC_RUNNER_20260619_131620:45; TFM_BENCH_VR_TEC_RUNNER_20260619_132936:51; TFM_BENCH_VR_TEC_RUNNER_20260619_134315:43 | CRITICA | BASELINE debe no observar cliente; VR_IDLE/VR_TEC_RUNNER deben observarlo. |
| Duplicados e inconsistencias | OK | duplicados=0; missing=0 | CRITICA | Duplicados o ausencias bloquean el Excel final. |
| Graficas generables con datos reales | OK | [{"Scenario":"BASELINE_NO_VR","AvgCPU":0,"AvgRAM":0},{"Scenario":"VR_IDLE","AvgCPU":0,"AvgRAM":57.48},{"Scenario":"VR_TEC_RUNNER","AvgCPU":0.95,"AvgRAM":57.42}] | CRITICA | No se generan graficas vacias. |

## KPIs del precheck

| Metrica | Valor |
| --- | --- |
| TotalRuns | 9 |
| ValidRuns | 9 |
| ScenarioCount | 3 |
| RepetitionsPerScenario | {"BASELINE_NO_VR":3,"VR_IDLE":3,"VR_TEC_RUNNER":3} |
| AvgCPU_BASELINE_NO_VR | 0 |
| AvgCPU_VR_IDLE | 0 |
| AvgCPU_VR_TEC_RUNNER | 0.95 |
| AvgRAM_BASELINE_NO_VR | 0 |
| AvgRAM_VR_IDLE | 57.48 |
| AvgRAM_VR_TEC_RUNNER | 57.42 |
| PeakCPU | 54 |
| PeakRAM | 57.97 |

## Resumen por escenario

| Scenario | RepetitionsExpected | RepetitionsFound | ValidRuns | AvgDurationSeconds | AvgCPUPercent | MaxCPUPercent | AvgRAMMB | MaxRAMMB | RunnerObservedSamples | Warnings | Interpretation |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| BASELINE_NO_VR | 3 | 3 | 3 | 182.69 | 0 | 0 | 0 | 0 | 0 |  | Referencia de laboratorio con servicio Velociraptor detenido; CPU/RAM del cliente igual a 0 por diseno. |
| VR_IDLE | 3 | 3 | 3 | 181.57 | 0 | 0 | 57.48 | 57.92 | 0 |  | Cliente Velociraptor activo en reposo; coste operativo observado bajo en CPU y estable en RAM. |
| VR_TEC_RUNNER | 3 | 3 | 3 | 315.6 | 0.95 | 54 | 57.42 | 57.97 | 139 | El runner sigue activo al finalizar la ventana de benchmark. No se mata el proceso. \| RunnerStillRunningAtEnd=True tratado como WARN no bloqueante. | Cliente bajo carga controlada de runner TEC; RunnerStillRunningAtEnd se conserva como WARN no bloqueante. |

## Discrepancias y decisiones

| Tipo | Ambito | Detalle | Decision | Fuente |
| --- | --- | --- | --- | --- |
| WARN_NO_BLOQUEANTE | VR_TEC_RUNNER | RunnerStillRunningAtEnd=True en runs VR_TEC_RUNNER. | Se conserva como advertencia metodologica; no invalida el benchmark porque no afecta a CPU/RAM del cliente Velociraptor ya muestreadas. | <RAIZ_TFM>\06_CONTROLLED_RERUN\OUTPUT\BENCHMARK\BENCH_20260619_FINAL01\REP_01\VR_TEC_RUNNER\TFM_BENCH_VR_TEC_RUNNER_20260619_131620\summary.json \| <RAIZ_TFM>\06_CONTROLLED_RERUN\OUTPUT\BENCHMARK\BENCH_20260619_FINAL01\REP_02\VR_TEC_RUNNER\TFM_BENCH_VR_TEC_RUNNER_20260619_132936\summary.json \| <RAIZ_TFM>\06_CONTROLLED_RERUN\OUTPUT\BENCHMARK\BENCH_20260619_FINAL01\REP_03\VR_TEC_RUNNER\TFM_BENCH_VR_TEC_RUNNER_20260619_134315\summary.json |
| INFO_METODOLOGICA | SERVER_GUI | SERVER_GUI detectado en process_samples. | Excluido del calculo principal mediante IncludeServerGuiInTotal=False y ServerGuiExcludedFromClientMetrics=True. | <RAIZ_TFM>\06_CONTROLLED_RERUN\OUTPUT\BENCHMARK\BENCH_20260619_FINAL01 |
| INFO_REFERENCIA | Analisis_Benchmark.xlsx | Excel historico localizado. | Usado solo como referencia historica/formato, no como resultado definitivo. | <RAIZ_TFM>\04_EVIDENCE\Analisis_Benchmark.xlsx |

## Decision metodologica

Los datos actuales son suficientes para generar el Excel definitivo. RunnerStillRunningAtEnd=True se mantiene como WARN no bloqueante en VR_TEC_RUNNER. SERVER_GUI esta observado pero excluido del calculo principal. CPU/RAM se interpretan solo como coste operativo observado en laboratorio.
