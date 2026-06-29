# PORCENTAJE FIABILIDAD - CONCLUSIONES

## Resultado ejecutivo

- Cobertura ofensiva: **9/9 tecnicas = 100 %**.
- Runner ofensivo: **OK/WARN/FAIL = 9/0/0**.
- FP v1.1: **9 OK, 1 SKIPPED, 0 FAIL**.
- Hits FP v1.1: **15**.
- P1_CRITICAL en FP: **0**.
- P2_EVENT en FP: **0**.
- Fiabilidad critica: **100 %** para esta muestra.

## Interpretacion tecnica

- Velociraptor detecta TEC-001 a TEC-009 en el laboratorio.
- Tras recalibracion, P1 deja de ser catch-all y opera como perfil critico.
- En FP v1.1 no se observaron falsos positivos criticos ni forenses.
- El ruido benigno se concentra en:
  - TEC-002: cmd.exe legitimo.
  - TEC-006: discovery legitimo de seguridad/Defender.
- ZIP local, borrado controlado y registro no persistente no escalaron a P1.

## Calidad de datos

- EXACT_RUNID: 3.
- TIME_WINDOW_INFERRED: 9.
- PREVIOUS_RUN_CONTAMINATION: 3.
- Los hits contaminados por TFM_FP_20260618_151244 estan marcados y no deben usarse para porcentajes principales salvo como bloque separado.

## Limitaciones

- Laboratorio controlado.
- Muestra pequena de falsos positivos.
- Parte de los hits queda como UNKNOWN por falta de RunId embebido en la evidencia.
- Los recuentos SOC ofensivos posteriores a P1 recalibrado proceden de contexto confirmado por el autor; no se localizo un CSV local con ese agregado exacto.
- No sustituye una validacion estadistica amplia en produccion.
