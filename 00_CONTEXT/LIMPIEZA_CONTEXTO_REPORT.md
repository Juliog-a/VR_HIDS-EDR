# LIMPIEZA_CONTEXTO_REPORT

Fecha: 2026-07-10 10:26 CEST.

## Objetivo

Reducir ruido contextual para futuras iteraciones de Codex, separar fuentes
vigentes de histórico y dejar explícito qué debe leerse y qué no.

## Archivos movidos

### A `00_CONTEXT/OLD_CONTEXT_USELESS`

- `00_CONTEXT/PROJECT_STATE.md`
- `00_CONTEXT/CURRENT_TASK.md`
- `00_CONTEXT/DECISIONS.md`
- `00_CONTEXT/ARTIFACT_INDEX.md`
- `00_CONTEXT/TEST_MATRIX.md`
- `00_CONTEXT/EVIDENCE_INDEX.md`
- `00_CONTEXT/CODEX_RULES.md`
- `00_CONTEXT/ANALISIS_DIFERENCIAS_TFM_vs_BETA_PREPDF.txt`
- `00_CONTEXT/CLEANUP_20260629.md`
- `00_CONTEXT/CLOSE.md`
- `00_CONTEXT/OPEN.md`
- `00_CONTEXT/CONTEXT_LAST.txt`
- `00_CONTEXT/CONTEXT_ESTADO_BENCHMARK_26062026.md`
- `00_CONTEXT/CONTEXT_ESTADO_VISIBILIDAD_FP_26062026.md`
- `00_CONTEXT/PORCENTAJE_FIABILIDAD_CONTEXT.md`
- `00_CONTEXT/VR_RESOURCE_BENCHMARK_CONTEXT.md`
- `00_CONTEXT/RESUMEN_CONTEXTO_CHATGPT_CODEX_PREPDF.md`
- `00_CONTEXT/estructura_COMPARTIDA_detallada.txt`
- `00_CONTEXT/estructura_COMPARTIDA_tree.txt`
- `00_CONTEXT/estructura_HOST_TFM_detallada.txt`
- `00_CONTEXT/estructura_HOST_TFM_tree.txt`

Todos se renombraron con sufijo:

`__NO_LEER_POR_DEFECTO__PRE_LIMPIEZA_20260710_1026`

### A `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`

- `03_RUNNERS/TFM_Run_All_TEC_Tests_v2.ps1`
- `03_RUNNERS/TFM_Run_All_TEC_Tests_v3.ps1`
- `03_RUNNERS/TFM_Run_All_TEC_Tests_v4.ps1`
- `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`
- `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1_PRE_SKIP_PARAM_BROKEN_BACKUP.ps1`
- `03_RUNNERS/TFM_Generate_VR_Benchmark_Excel_v1_PRE_ROLE_FIX_BACKUP.ps1`
- `03_RUNNERS/TFM_Run_FP_Tests_v1_OLD_BROKEN_POSTPROCESS_BACKUP.ps1`

## Archivos conservados

- `01_ARTIFACTS/validated`: no modificado.
- `04_EVIDENCE`: no se movieron evidencias finales.
- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`: vigente.
- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx`: original conservado.
- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/_ARCHIVO_OBSOLETO_NO_USAR`: ya estaba aislado.
- `99_ARCHIVE`: conservado.
- `10_WAZUH`: conservado.
- `00_CONTEXT/SESSION_LOGS/session_*.md`: conservados por trazabilidad.

## Archivos creados o actualizados

- `00_CONTEXT/README_CONTEXT.md`
- `00_CONTEXT/PROJECT_STATE.md`
- `00_CONTEXT/CURRENT_TASK.md`
- `00_CONTEXT/DECISIONS.md`
- `00_CONTEXT/ARTIFACT_INDEX.md`
- `00_CONTEXT/TEST_MATRIX.md`
- `00_CONTEXT/EVIDENCE_INDEX.md`
- `00_CONTEXT/CODEX_RULES.md`
- `00_CONTEXT/INVENTARIO_PROYECTO_LIMPIO.md`
- `00_CONTEXT/INSTRUCCIONES_CODEX_FUTURAS.md`
- `00_CONTEXT/LIMPIEZA_CONTEXTO_REPORT.md`
- `00_CONTEXT/OLD_CONTEXT_USELESS/README.md`
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO/README.md`
- `01_ARTIFACTS/candidate/README_NO_LEER_POR_DEFECTO.md`
- `01_ARTIFACTS/debug/README_NO_LEER_POR_DEFECTO.md`
- `02_SCRIPTS/candidate/README_NO_LEER_POR_DEFECTO.md`
- `07_DOCS/candidate/README_NO_LEER_POR_DEFECTO.md`
- `AGENTS.md`
- `00_CONTEXT/SESSION_LOGS/session_20260710_1026.md`
- `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`

## Contradicciones detectadas

| Tema | Contradicción | Criterio aplicado |
|---|---|---|
| TEC-009 | Contextos antiguos exigían o afirmaban `receiver_log.jsonl`/ZIP recibido; la revisión pre-PDF permite su ausencia como limitación secundaria. | Prevalece memoria beta revisada: TEC-009 sigue siendo exfiltración HTTP controlada si hay HTTP 200, `UploadSucceeded=True`, SHA-256 y bytes transferidos. |
| Benchmark | Contextos antiguos marcaban benchmark pendiente/no válido; contexto final 26062026 lo acepta. | Prevalece benchmark final 26062026: 9/9 runs válidos, coste operativo, no calidad de detección. |
| FP | Contextos antiguos tenían 9 OK/1 SKIPPED o hits contaminados; final 26062026 tiene 10/10 OK y `VelociraptorHits.Total=0`. | Prevalece Excel final V.5 y normalizados de 26062026. |
| Artifacts | Índices antiguos daban peso a P3/P2/P1 candidates y debug. | Prevalece `01_ARTIFACTS/validated`; candidate/debug no se lee por defecto. |
| Memoria | README beta 20260708 indicaba abrir `TFM_MEMORIA_BETA_v1.docx`; revisión 20260709 indica trabajar en `REVISION_PREPDF`. | Prevalece `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`; original no se toca sin orden expresa. |
| JSONL/Discord | Una frase de capítulo VI puede tratarlos como resultado final primario. | Debe reescribirse: son salidas externas trazables derivadas de `CLIENT_EVENT`. |

## Criterio usado

1. Lo más reciente prevalece si no contradice evidencias finales.
2. Memoria beta revisada pre-PDF prevalece sobre contextos antiguos.
3. `PROJECT_STATE.md` nuevo resume el estado actual.
4. Evidencias finales y Exceles finales prevalecen sobre informes legacy.
5. No se modifican resultados técnicos validados.
6. No se borran ficheros; se mueven a carpetas de no lectura por defecto.

## Dudas pendientes

- Confirmar manualmente si Word sigue abierto y si desaparece
  `~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
- Confirmar visualmente en Word la eliminación del índice duplicado y la
  actualización de índices/listas.
- Confirmar si se desea actualizar la memoria original `TFM_MEMORIA_BETA_v1.docx`
  después de aprobar la copia `REVISION_PREPDF`.

## Recomendaciones

- Usar `README_CONTEXT.md` como entrada única.
- No leer `OLD_CONTEXT_USELESS` ni `candidate` salvo petición explícita.
- Resolver primero los P1 de `CHECKLIST_P1_ANTES_PDF.md`.
- Mantener `01_ARTIFACTS/validated` intacto.
- No modificar cifras finales sin nueva evidencia explícita.
- Antes de una sesión larga, revisar `INVENTARIO_PROYECTO_LIMPIO.md`.
