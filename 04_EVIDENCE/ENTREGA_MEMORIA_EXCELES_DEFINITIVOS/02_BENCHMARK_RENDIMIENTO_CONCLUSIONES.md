# VR resource benchmark - conclusiones de validacion

## Estado

Los logs de `05_LOGS/BENCHMARKS` no son validos para cerrar el benchmark definitivo de consumo del agente cliente Velociraptor.

## Evidencia

- Los tres escenarios existen y tienen estructura completa.
- Los tres `summary.json` declaran `SchemaVersion=1.0`.
- `VR_IDLE` y `VR_TEC_RUNNER` muestran `ServiceStatus=Stopped`.
- No aparece `C:\Program Files\Velociraptor\Velociraptor.exe` ni comando con `client.config.yaml service run`.
- `notepad.exe` fue incluido por contener rutas con la palabra Velociraptor.
- `velociraptor-v0.75.6-windows-amd64.exe` se clasifica como binario server/laboratorio y no como agente HIDS.

## Interpretacion

Las metricas legacy de CPU/RAM no deben usarse como coste del agente cliente. Sirven como evidencia de que el benchmark anterior mezclaba roles y necesitaba la correccion ya documentada.

## Pendiente

Repetir `BASELINE_NO_VR`, `VR_IDLE` y `VR_TEC_RUNNER` con el script corregido y `-SkipServiceControl $true`, confirmando que `CLIENT_SERVICE` aparece en idle/runner y no aparece en baseline.