# PORCENTAJE_FIABILIDAD_CONTEXT

## Archivos revisados

- 00_CONTEXT/PROJECT_STATE.md
- 00_CONTEXT/CURRENT_TASK.md
- 00_CONTEXT/DECISIONS.md
- 00_CONTEXT/SESSION_LOGS/LAST_SESSION.md
- 00_CONTEXT/EVIDENCE_INDEX.md
- 00_CONTEXT/CODEX_RULES.md
- 05_LOGS/FPS/TFM_FP_20260618_151316_summary.json
- 05_LOGS/FPS/TFM_FP_20260618_151316_vr_hits.csv
- 05_LOGS/FPS/TFM_FP_20260618_151316_summary.txt
- 05_LOGS/TFM_TEC_Run_CANDIDATE_v5_20260616_174626_summary.txt
- 05_LOGS/TFM_TEC_Run_CANDIDATE_v5_20260616_174626_summary.csv
- 07_DOCS/MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT.md
- 01_ARTIFACTS/validated/last
- 01_ARTIFACTS/candidate/last_version

## Excel generado

- C:\Users\julio\Desktop\TFM\PORCENTAJE FIABILIDAD.xlsx

## Hojas

1. RESUMEN_EJECUTIVO
2. TEC_ATTACK_VALIDATION
3. FP_TEST_MATRIX
4. FP_HITS_RAW
5. METRICAS
6. GRAFICAS
7. CONCLUSIONES
8. TRAZABILIDAD

## Metricas calculadas

- Cobertura ofensiva.
- Tasa de ejecucion FP.
- Tasa de deteccion benigna.
- Tasa FP critico.
- Tasa FP forense.
- Tasa FP medio/bajo.
- Precision critica aproximada.
- Fiabilidad critica.
- Alertas FP por prueba ejecutada.
- Alertas por perfil, tecnica y DetectionName.

## Cautelas de calidad

- FP_HITS_RAW separa EXACT_RUNID, TIME_WINDOW_INFERRED, PREVIOUS_RUN_CONTAMINATION y UNKNOWN.
- En TFM_FP_20260618_151316_vr_hits.csv hay evidencias del RunId anterior TFM_FP_20260618_151244.
- Los hits contaminados se conservan para auditoria, pero quedan marcados.
- La ruta 01_ACTIVE_TESTS de la VM no existe en este host.
- TEST_MATRIX.md no existe en 00_CONTEXT.
- Los recuentos SOC ofensivos posteriores a P1 recalibrado proceden de contexto confirmado por el autor.

## Conclusiones utilizables en memoria

- Velociraptor cubre las 9 tecnicas ofensivas evaluadas.
- P1 recalibrado no genera FP criticos en FP v1.1.
- No hay FP forenses en FP v1.1.
- El ruido benigno queda en P3/P4 y se concentra en cmd.exe legitimo y discovery de seguridad.
- La fiabilidad critica observada en esta muestra es 100 %.

## Pendiente

- Conservar las evidencias de la VM si se requiere trazabilidad completa de soc_alerts.jsonl.
- Repetir FP con mayor muestra si se desea validez estadistica.
- No modificar P1/P2/P3/P4 solo con estos FP; documentar y ampliar muestra.
