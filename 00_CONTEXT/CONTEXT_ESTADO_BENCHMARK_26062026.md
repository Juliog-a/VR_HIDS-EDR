# CONTEXT_ESTADO_BENCHMARK_26062026

Fecha: 2026-06-26 13:02 CEST

## Excel generado

Se genero el Excel definitivo de benchmark:

- `C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_benchmark\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`
- Copia trazable:
  `C:\Users\julio\Desktop\TFM\04_EVIDENCE\Excel_benchmark_26062026\TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`
- SHA-256:
  `AEB8835BAD0B2A8D015D399DDA11A2CD475B3F320750ABD333750026456A36EC`

Artefactos asociados:

- `TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_AUDIT.md`
- `TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_DATA_QUALITY.json`
- `BENCHMARK_PRECHECK_26062026.md`
- `benchmark_generation_26062026.log`
- `GRAFICAS_PREVIEW_26062026.pdf`
- CSV normalizados en `04_EVIDENCE\Excel_benchmark_26062026\normalized`.

## Fuente canonica usada

Fuente canonica:

- `C:\Users\julio\Desktop\TFM\04_EVIDENCE\ENTREGA_MEMORIA_EXCELES_REGENERADOS\02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`

Fuente de verificacion y trazabilidad:

- `C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN\OUTPUT\BENCHMARK\BENCH_20260619_FINAL01`

`Analisis_Benchmark.xlsx` queda solo como referencia historica/formato.

## Escenarios

Escenarios incluidos:

- `BASELINE_NO_VR`
- `VR_IDLE`
- `VR_TEC_RUNNER`

Cada escenario tiene 3 repeticiones validas.

## Runs validos

- Runs totales: 9.
- Runs validos: 9.
- `SchemaVersion=1.1`: 9/9.
- `ScenarioValidity=VALID`: 9/9.
- `ValidForClientBenchmark=True`: 9/9.
- Muestras `samples.csv`: 422 filas.
- Muestras `process_samples.csv`: 770 filas.
- Filas `CLIENT_SERVICE`: 284.
- `SERVER_GUI`: observado, pero excluido del calculo principal.
- `notepad.exe` como runner: 0.

## Metricas principales

- CPU media `BASELINE_NO_VR`: 0 %.
- CPU media `VR_IDLE`: 0 %.
- CPU media `VR_TEC_RUNNER`: 0,95 %.
- RAM media `BASELINE_NO_VR`: 0 MB.
- RAM media `VR_IDLE`: 57,48 MB.
- RAM media `VR_TEC_RUNNER`: 57,42 MB.
- Pico CPU cliente Velociraptor: 54 %.
- Pico RAM cliente Velociraptor: 57,97 MB.
- Duracion media `BASELINE_NO_VR`: 182,69 s.
- Duracion media `VR_IDLE`: 181,57 s.
- Duracion media `VR_TEC_RUNNER`: 315,6 s.

## Limitaciones

- Laboratorio Windows 10 virtualizado.
- Mediciones relativas, no extrapolables directamente a produccion.
- Benchmark centrado en el agente cliente Velociraptor.
- `SERVER_GUI` excluido del calculo principal.
- `RunnerStillRunningAtEnd=True` conservado como WARN no bloqueante.
- El benchmark no mide calidad de deteccion, alertas, evidencia ni falsos positivos.
- Tres repeticiones por escenario.
- Posible ruido residual del sistema operativo.

## Discrepancias

- `RunnerStillRunningAtEnd=True` en los 3 runs `VR_TEC_RUNNER`; se conserva como WARN no bloqueante.
- `SERVER_GUI` aparece en `process_samples.csv`, pero no suma en metricas de cliente.
- `Analisis_Benchmark.xlsx` no se usa como resultado definitivo.

No se detectaron discrepancias criticas bloqueantes.

## Decision de entrega

Decision: APTO para memoria.

El Excel final se acepta porque cumple:

- 9/9 runs validos.
- 3 escenarios.
- 3 repeticiones por escenario.
- Graficas pobladas.
- Hojas obligatorias presentes.
- Formula check sin errores.
- Trazabilidad con SHA-256.
- Audit y data quality generados.

## Proximos pasos Capitulo VII

- Usar `RESUMEN_EJECUTIVO`, `RESUMEN_ESCENARIO`, `GRAFICAS`, `CALIDAD_DATOS` y `LIMITACIONES` para redactar el apartado de coste operativo.
- Presentar CPU/RAM como medicion del laboratorio, no como conclusion universal.
- Separar este resultado del capitulo de deteccion/visibilidad.
- Citar `RunnerStillRunningAtEnd=True` como advertencia metodologica no bloqueante.
