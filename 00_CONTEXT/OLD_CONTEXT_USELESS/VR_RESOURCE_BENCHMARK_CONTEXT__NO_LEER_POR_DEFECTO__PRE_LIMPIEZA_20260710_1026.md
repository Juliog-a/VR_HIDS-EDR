# VR_RESOURCE_BENCHMARK_CONTEXT

## Correccion vigente

- TFM_Benchmark_VR_Resource_Usage_v1.ps1 incluye SkipServiceControl en param().
- El benchmark separa CLIENT_SERVICE, SERVER_GUI y OTHER_VELOCIRAPTOR.
- El Excel tiene hoja CALIDAD_DATOS.
- No usar benchmarks anteriores como definitivos si no proceden del script corregido.

## Estado Excel

- Estado: DATOS_PARCIALES_O_LEGACY.
- Hojas: 9.

## Interpretacion

- CPU media cliente idle < 2 %: impacto bajo.
- CPU media cliente runner < 10 %: impacto bajo/moderado.
- RAM estable sin crecimiento progresivo: aceptable.
- Server/GUI local no forma parte del coste del agente HIDS.
