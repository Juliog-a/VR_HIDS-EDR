# INVENTARIO_PROYECTO_LIMPIO

Actualizado: 2026-07-10 10:26 CEST.

| Ruta | Tipo de archivo/carpeta | Estado | Motivo | Leer en futuras iteraciones |
|---|---|---|---|---|
| `AGENTS.md` | Instrucciones raíz | ACTIVO | Reglas obligatorias del repositorio. | Sí |
| `00_CONTEXT/README_CONTEXT.md` | Contexto maestro | ACTIVO | Puerta de entrada limpia. | Sí |
| `00_CONTEXT/PROJECT_STATE.md` | Estado maestro | ACTIVO | Estado actual sin cronología ruidosa. | Sí |
| `00_CONTEXT/CURRENT_TASK.md` | Tarea viva | ACTIVO | Pendientes reales antes del PDF. | Sí |
| `00_CONTEXT/DECISIONS.md` | Decisiones finales | ACTIVO | Metodología cerrada. | Sí |
| `00_CONTEXT/ARTIFACT_INDEX.md` | Índice de artifacts | ACTIVO | Distingue validated/candidate/debug. | Sí, si hay tarea de artifacts |
| `00_CONTEXT/TEST_MATRIX.md` | Matriz de pruebas | ACTIVO | Separación runner/CLIENT_EVENT/router/JSONL/evidencia. | Sí, si hay tarea técnica |
| `00_CONTEXT/EVIDENCE_INDEX.md` | Índice de evidencias | ACTIVO | Rutas finales y criterios de evidencia. | Sí, si hay tarea de resultados |
| `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md` | Última sesión | ACTIVO | Cierre más reciente. | Sí |
| `00_CONTEXT/CODEX_RULES.md` | Reglas operativas | ACTIVO | Reglas compactas de trabajo y cierre. | Sí |
| `00_CONTEXT/VM_ACTIVE_PATHS.md` | Rutas VM | ACTIVO | Aclara que la VM no es fuente oficial. | Sí, si hay laboratorio |
| `00_CONTEXT/OLD_CONTEXT_USELESS` | Histórico contextual | NO_LEER_POR_DEFECTO | Contextos superados, contradictorios o largos. | No |
| `00_CONTEXT/SESSION_LOGS/session_*.md` | Histórico de sesiones | HISTÓRICO | Trazabilidad. | No, salvo auditoría |
| `01_ARTIFACTS/validated` | Artifacts Velociraptor | FINAL | Fuente final de artifacts. | Sí, solo si la tarea lo requiere |
| `01_ARTIFACTS/candidate` | Artifacts candidatos | CANDIDATE / NO_LEER_POR_DEFECTO | Versiones no finales y backups. | No |
| `01_ARTIFACTS/debug` | Artifacts diagnóstico | HISTÓRICO / NO_LEER_POR_DEFECTO | Debug de 4104/Sysmon. | No |
| `02_SCRIPTS/validated` | Scripts validados | ACTIVO | Scripts de apoyo validados. | Sí, si la tarea lo requiere |
| `02_SCRIPTS/candidate` | Scripts candidatos | CANDIDATE / NO_LEER_POR_DEFECTO | Duplicados o versiones no finales. | No |
| `03_RUNNERS/validate` | Runners validación | ACTIVO | Scripts de validación/control. | Sí, si hay rerun/validación |
| `03_RUNNERS/TFM_Run_All_TEC_Tests_v6.ps1` | Runner TEC | ACTIVO/HISTÓRICO OPERATIVO | Última versión no archivada en raíz. | Solo si la tarea trata runners |
| `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1` | Runner FP | ACTIVO | Runner FP corregido. | Solo si la tarea trata FP |
| `03_RUNNERS/TFM_Benchmark_VR_Resource_Usage_v1.ps1` | Benchmark | ACTIVO | Medición de coste operativo. | Solo si la tarea trata benchmark |
| `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO` | Runners antiguos | OBSOLETO / NO_LEER_POR_DEFECTO | v2-v5 y backups rotos/superados. | No |
| `04_EVIDENCE/Excel_visibilidad_26062026` | Evidencia final | FINAL | Resultados VR/FP/Wazuh y normalizados. | Sí, si se verifican resultados |
| `04_EVIDENCE/Excel_benchmark_26062026` | Evidencia final | FINAL | Benchmark final y normalizados. | Sí, si se verifica benchmark |
| `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS` | Evidencia derivada | FINAL | Fuente canónica benchmark regenerado. | Solo si hace falta trazabilidad |
| `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS` | Entrega antigua | HISTÓRICO | Paquete previo auditado. | No |
| `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS_REVIEWED` | Auditoría antigua | HISTÓRICO | Marcó bloqueantes ya superados por finales. | No |
| `04_EVIDENCE/legacy_reports` | Informes legacy | HISTÓRICO / NO_LEER_POR_DEFECTO | Conservados por trazabilidad. | No |
| `04_EVIDENCE/diag_p3v3` | Diagnóstico | HISTÓRICO / NO_LEER_POR_DEFECTO | Debug antiguo P3. | No |
| `04_EVIDENCE/TEC009` | Evidencia específica | HISTÓRICO/INTERMEDIA | Diagnósticos TEC-009 previos. | No, salvo petición TEC-009 |
| `06_CONTROLLED_RERUN` | Paquete reproducible | ACTIVO/HISTÓRICO | Scripts de repetición controlada y validación. | Sí, si hay rerun |
| `07_DOCS` | Guías técnicas | ACTIVO/HISTÓRICO | Documentación auxiliar. | Solo si la tarea lo requiere |
| `07_DOCS/candidate` | Docs candidatos | CANDIDATE / NO_LEER_POR_DEFECTO | Planes y notas no finales. | No |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx` | Word memoria | ACTIVO | Documento vigente pre-PDF. | Sí |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx` | Word beta original | FINAL BASE / NO_EDITAR | Original intacto para referencia. | Solo comparación |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/CHECKLIST_P1_ANTES_PDF.md` | Checklist | ACTIVO | Bloqueantes exactos antes del PDF. | Sí, si se trabaja memoria |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/REVISION_ESTADO_FINAL_BETA_10_10.md` | Revisión | ACTIVO | Estado académico/técnico pre-PDF. | Sí, si se trabaja memoria |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/_ARCHIVO_OBSOLETO_NO_USAR` | Memoria antigua | NO_LEER_POR_DEFECTO | Históricos ya archivados. | No |
| `08_MEMORIA/Excel_visibilidad_FP` | Exceles memoria | FINAL | Copias usadas por memoria. | Sí, si se trabaja resultados |
| `08_MEMORIA/Excel_benchmark` | Exceles memoria | FINAL | Benchmark usado por memoria. | Sí, si se trabaja benchmark |
| `08_MEMORIA/ENTREGABLE` | Entregables Excel | FINAL | Copias para entrega. | Solo si se empaqueta entrega |
| `08_MEMORIA/TFM.docx` | Word histórico/base | HISTÓRICO/BASE | Documento real anterior usado como base. | No por defecto |
| `10_WAZUH` | Comparación Wazuh | ACTIVO | Reglas y material Wazuh. | Solo si la tarea trata Wazuh |
| `99_ARCHIVE` | Archivo histórico | HISTÓRICO / NO_LEER_POR_DEFECTO | Iteraciones y prompts antiguos. | No |
| `README.md` | README público | ACTIVO | Descripción general del repositorio. | Opcional |
