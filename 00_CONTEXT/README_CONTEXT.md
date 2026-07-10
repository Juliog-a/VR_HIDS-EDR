# README_CONTEXT

Actualizado: 2026-07-10 10:26 CEST.

## Orden de lectura recomendado

1. `AGENTS.md`.
2. `00_CONTEXT/README_CONTEXT.md`.
3. `00_CONTEXT/PROJECT_STATE.md`.
4. `00_CONTEXT/CURRENT_TASK.md`.
5. `00_CONTEXT/DECISIONS.md`.
6. `00_CONTEXT/ARTIFACT_INDEX.md`.
7. `00_CONTEXT/TEST_MATRIX.md`.
8. `00_CONTEXT/EVIDENCE_INDEX.md`.
9. `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`.
10. `00_CONTEXT/CODEX_RULES.md`.
11. `00_CONTEXT/VM_ACTIVE_PATHS.md`.

Para una sesión rápida, los puntos 2 a 5 suelen ser suficientes salvo que la
tarea afecte a artifacts, evidencias, pruebas o VM.

## Carpetas que deben ignorarse por defecto

- `00_CONTEXT/OLD_CONTEXT_USELESS`.
- Cualquier carpeta `candidate`.
- `01_ARTIFACTS/debug`.
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`.
- `99_ARCHIVE`.
- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/_ARCHIVO_OBSOLETO_NO_USAR`.

Estas rutas solo se leen si el usuario lo pide explícitamente o si la tarea es
recuperar histórico.

## Contexto histórico

Histórico principal:

- `00_CONTEXT/OLD_CONTEXT_USELESS`: contextos largos, contradictorios o
  superados por el estado actual.
- `99_ARCHIVE`: iteraciones y prompts antiguos.
- `04_EVIDENCE/legacy_reports`: informes legacy conservados por trazabilidad.
- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/_ARCHIVO_OBSOLETO_NO_USAR`:
  documentos de memoria antiguos ya aislados.

No leer estas rutas por defecto.

## Evidencias finales

Evidencia final de visibilidad/detección/FP/Wazuh:

- `04_EVIDENCE/Excel_visibilidad_26062026`.
- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`.

Evidencia final de benchmark:

- `04_EVIDENCE/Excel_benchmark_26062026`.
- `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`.

Artifacts finales:

- `01_ARTIFACTS/validated`.

Memoria vigente:

- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.

## Candidatos no finales

No leer por defecto:

- `01_ARTIFACTS/candidate`.
- `02_SCRIPTS/candidate`.
- `07_DOCS/candidate`.

Solo se leen si la tarea solicita desarrollo, comparación o recuperación de
versiones candidatas.

## Regla de resolución de contradicciones

Si dos documentos se contradicen, prevalece este orden:

1. Memoria beta revisada pre-PDF.
2. Documentos de revisión pre-PDF.
3. `PROJECT_STATE.md` actual.
4. `LAST_SESSION.md` actual.
5. Evidencias finales validadas.
6. Contextos históricos solo como trazabilidad.

## Restricciones fijas

- No tocar `01_ARTIFACTS/validated`.
- No cambiar cifras finales sin evidencia nueva explícita.
- No degradar TEC-009.
- No tratar JSONL/Discord como detección primaria.
- No tratar `SERVER_EVENT` como detector.
- No extrapolar a toda MITRE ni a producción.
