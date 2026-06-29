# NOTAS DE CALIDAD DE DATOS

## Tecnicas, FP y fiabilidad

El Excel `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx` se basa en:

- campana ofensiva TEC validada con 9/9 tecnicas detectadas;
- FP v1.1 `TFM_FP_20260618_151316`;
- logs de `05_LOGS/FPS`;
- separacion de calidad de asociacion:
  - `EXACT_RUNID`;
  - `TIME_WINDOW_INFERRED`;
  - `PREVIOUS_RUN_CONTAMINATION`;
  - `UNKNOWN`.

Conclusion soportada:

- cobertura ofensiva: 100 % en laboratorio;
- P1_CRITICAL en FP v1.1: 0;
- P2_EVENT en FP v1.1: 0;
- fiabilidad critica: 100 % para esta muestra;
- ruido benigno concentrado en P3/P4.

Limitaciones:

- muestra FP pequena;
- laboratorio controlado;
- parte de los hits FP se asocia por ventana temporal;
- hay contaminacion temporal documentada de un RunId anterior.

## Benchmark

El Excel `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx` no debe presentarse
como medicion definitiva de consumo del agente.

Motivos:

- los logs son `SchemaVersion=1.0`;
- `VR_IDLE` y `VR_TEC_RUNNER` muestran `ServiceStatus=Stopped`;
- no se observa el proceso cliente:
  `C:\Program Files\Velociraptor\Velociraptor.exe` con
  `client.config.yaml service run`;
- se detectaron falsos matches de `notepad.exe`;
- el binario `velociraptor-v0.75.6-windows-amd64.exe` se clasifica como
  server/laboratorio, no como agente HIDS.

Conclusion soportada:

- el benchmark debe repetirse antes de cerrar el apartado de rendimiento;
- el Excel actual sirve como evidencia metodologica y control de calidad.
