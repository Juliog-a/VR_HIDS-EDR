# TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_AUDIT

Fecha generacion: 2026-06-26 13:01:26 +02:00

## Decision

APTO. El precheck confirma 9/9 runs validos, 3 escenarios y 3 repeticiones por escenario. No se han detectado fallos criticos.

## Salidas

- Excel final memoria: C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_benchmark\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx
- Excel trazabilidad: C:\Users\julio\Desktop\TFM\04_EVIDENCE\Excel_benchmark_26062026\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx
- Data quality JSON: C:\Users\julio\Desktop\TFM\04_EVIDENCE\Excel_benchmark_26062026\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_DATA_QUALITY.json
- Precheck: C:\Users\julio\Desktop\TFM\04_EVIDENCE\Excel_benchmark_26062026\BENCHMARK_PRECHECK_26062026.md
- Preview PDF graficas: C:\Users\julio\Desktop\TFM\04_EVIDENCE\Excel_benchmark_26062026\GRAFICAS_PREVIEW_26062026.pdf

## Fuente canonica

C:\Users\julio\Desktop\TFM\04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_REGENERADOS\02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx

Analisis_Benchmark.xlsx se mantiene solo como referencia historica/formato.

## KPIs

- Runs totales: 9
- Runs validos: 9
- Escenarios: 3
- CPU media BASELINE_NO_VR: 0 %
- CPU media VR_IDLE: 0 %
- CPU media VR_TEC_RUNNER: 0.95 %
- RAM media BASELINE_NO_VR: 0 MB
- RAM media VR_IDLE: 57.48 MB
- RAM media VR_TEC_RUNNER: 57.42 MB
- Pico CPU cliente VR: 54 %
- Pico RAM cliente VR: 57.97 MB

## Validacion final

- Apertura Excel: OK
- Hojas obligatorias presentes: True
- Hojas: 11
- Formulas: 0
- Errores de formula: 0
- Graficas en GRAFICAS: 6
- Graficas pobladas: True
- SERVER_GUI excluido: True
- notepad.exe runner: 0
- Fuentes con SHA-256: True

## Limitaciones metodologicas

- Laboratorio Windows 10 virtualizado.
- Mediciones relativas, no extrapolables directamente a produccion.
- Benchmark centrado en agente cliente Velociraptor.
- SERVER_GUI excluido del calculo principal.
- RunnerStillRunningAtEnd conservado como WARN no bloqueante.
- No mide calidad de deteccion, alertas, evidencia ni falsos positivos.
- Tres repeticiones por escenario.
- Posible ruido residual del sistema operativo.

## Discrepancias

- RunnerStillRunningAtEnd=True en VR_TEC_RUNNER: WARN no bloqueante.
- SERVER_GUI observado y excluido del calculo principal.
- Analisis_Benchmark.xlsx usado solo como referencia historica/formato.
