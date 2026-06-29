# ENTREGA MEMORIA - EXCELES DEFINITIVOS

## Proposito

Carpeta consolidada para entregar en la memoria del TFM los Excel finales y la
trazabilidad minima asociada.

## Exceles principales

1. `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx`
   - Entregable principal de deteccion, tecnicas MITRE, falsos positivos y
     fiabilidad.
   - Generado desde `PORCENTAJE FIABILIDAD.xlsx`.
   - Integra la campana ofensiva validada y la campana FP v1.1.

2. `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx`
   - Entregable de benchmark de rendimiento en estado de validacion de datos.
   - Documenta que los logs actuales no sirven para concluir consumo definitivo
     del agente cliente Velociraptor.
   - Incluye tablas y graficas de calidad de datos.

## Excel de referencia recomendado

3. `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx`
   - Copia de `04_EVIDENCE/Analisis_Tecnicas_TFM_V.4.xlsx`.
   - Se conserva como referencia historica del analisis de tecnicas.
   - No sustituye al Excel integrado final `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx`.

4. `04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx`
   - Copia de `04_EVIDENCE/Analisis_Artifacts_Publicos_Velociraptor_TFM_10_10.xlsx`.
   - Recomendado como anexo metodologico si la memoria explica por que se
     construyeron artifacts custom frente a artifacts publicos.
   - No es un resultado de campana FP/TEC ni de benchmark.

## Fuentes incluidas

- `fuentes_fp/`: summaries, `vr_hits`, conclusiones y hashes de falsos positivos.
- `fuentes_benchmark/`: tablas normalizadas y hashes del benchmark revisado.
- `fuentes_analisis/`: copias originales de los Excel historicos V4/DEF.

## Criterio de uso en memoria

- Usar `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx` para:
  - cobertura ofensiva;
  - deteccion TEC-001 a TEC-009;
  - falsos positivos;
  - fiabilidad critica;
  - trazabilidad de FP.

- Usar `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx` con cautela:
  - evidencia la revision del benchmark;
  - no permite afirmar consumo real definitivo del agente cliente;
  - debe describirse como pendiente de repeticion con `CLIENT_SERVICE` observado.

- Usar `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx` solo como antecedente o
  apoyo historico.

## Decision tecnica

No recomiendo entregar mas Exceles principales. Con estos dos bloques se cubren:

- eficacia/deteccion/FP/fiabilidad;
- rendimiento, aunque actualmente como validacion de calidad de datos.

El V4 y el analisis de artifacts publicos quedan como referencias/anexos para
auditoria documental y metodologia.
