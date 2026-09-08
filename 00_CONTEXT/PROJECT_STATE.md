# PROJECT_STATE

Actualizado: 2026-09-08 09:58 CEST.

## Limpieza final del repositorio — completada

- Eliminados 5.217 archivos `DELETE` y liberados 1.236.425.453 bytes.
- `KEEP`, `REVIEW` y `.git` preservados.
- Entregables finales reunidos en `08_MEMORIA/ENTREGA_TFM_Julgarrei`.
- Rutas del paquete reproducible reparadas hacia fuentes canónicas validadas.
- Self-test: PASS (17 scripts, 9 hashes, matriz TEC 9/9).
- Archivos actuales sin webhooks, credenciales, tokens ni claves privadas publicables.
- Informe: `CLEANUP_REPORT.md`.

---

Actualizado: 2026-09-07 12:00 CEST.

## Limpieza final del repositorio — en curso

- Inventario previo: 6.889 archivos, 681 directorios y 125,626 GiB fuera de
  `.git`; 127 GiB pertenecen a la VM local ignorada.
- Clasificación propuesta: 342 archivos `KEEP`, 5.217 `DELETE` y 1.330
  `REVIEW`; estos últimos son material local no publicable conservado por
  precaución.
- Duplicados SHA-256: 1.546 grupos antes de la limpieza.
- Se han neutralizado credenciales, webhooks, rutas personales e IP privadas
  en los entregables y fuentes de benchmark afectados, sin cambiar métricas ni
  fórmulas.
- La eliminación de 5.217 archivos está pendiente de confirmación explícita
  tras el bloqueo preventivo de la plataforma.
- No se ha ejecutado `git add`, commit ni push.
- Existe riesgo de secretos en el historial Git; no se ha reescrito ni borrado
  el historial.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260907_1200.md`.

---

Actualizado: 2026-09-02 09:58 CEST.

## Presentación de defensa TFM — 2026-09-02 09:58 CEST

- Nuevo entregable:
  `08_MEMORIA/ENTREGA/DEFENSA_TFM_Julio_Garcia_Amorena.pptx`.
- 12 diapositivas 16:9, diseño académico y visual, con notas del orador en
  12/12 diapositivas y duración prevista próxima a 11 minutos.
- Contenido limitado a `TFM.docx`,
  `Analisis_Tecnicas_TFM_Velociraptor.xlsx` y
  `TFM_BENCHMARK_RENDIMIENTO.xlsx` de `08_MEMORIA`.
- Las rutas equivalentes de `08_MEMORIA/ENTREGA_TFM_Julgarrei` son
  binariamente idénticas; las rutas solicitadas bajo `08_MEMORIA/ENTREGA` no
  existían al iniciar la sesión.
- QA: reapertura PowerPoint sin reparación, revisión visual 12/12 a
  1920x1080, 0 XML inválidos, 0 vínculos externos y 0 mojibake.
- SHA-256:
  `0F4851DCE2783E826C5E008D1421961FD07B430AEA4C1C42E81D631F12208891`.
- Se preserva la semántica: 379 filas `CLIENT_EVENT`, no alertas únicas;
  ETW como telemetría sin detección específica; TrackNetwork como contexto;
  Router/JSONL/Discord como salida; Velociraptor no se presenta como EDR.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260902_0958.md`.

## Excel de visibilidad final consolidado — 2026-08-27 16:01 CEST

- Entregables actualizados:
  `Analisis_Tecnicas_TFM_Velociraptor_2.xlsx` y
  `Analisis_Tecnicas_TFM_Velociraptor_FINAL.xlsx`.
- Estructura final: 17 hojas, 22 tablas y 4 gráficos; los dos ficheros son
  binariamente idénticos.
- Arquitectura reorganizada por metodología, evaluación, resultados públicos,
  resultados custom, FP, alertabilidad, gaps, Wazuh y trazabilidad.
- Las dos validaciones `#REF!` heredadas fueron corregidas o retiradas de forma
  segura, sin modificar resultados experimentales ni inventar categorías.
- QA: 25/25 PASS, 33 hipervínculos internos, 0 vínculos externos, 0 errores de
  celda, 0 XML inválidos y 0 partes de recuperación.
- SHA-256:
  `93F74E4BC7F2EC371C6342E5191F1068E500620505456A7733EA601BF78B6F82`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260827_1601.md`.

## Excel de visibilidad enriquecido — 2026-08-27 10:44 CEST

- Nuevo entregable no destructivo:
  `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_Velociraptor_2.xlsx`.
- Fuente de verdad preservada:
  `Analisis_Tecnicas_TFM_Velociraptor.xlsx`; los otros tres libros se usaron
  solo como contraste de estructura y presentación.
- Estructura final: 30 hojas, 15 tablas y 15 gráficos.
- Añadidos: dashboard ampliado con cuatro gráficos enlazados, guía de lectura,
  índice navegable completo y numeración de las 30 hojas.
- No se reintrodujeron `09_Benchmark_Plan`, `99_Listas`, `FP_HITS` ni
  `DISCREPANCIAS`.
- QA: 49.339 celdas/fórmulas preservadas con 0 diferencias, 0 XML inválidos,
  15/15 gráficos con series y revisión visual de 5/5 páginas de presentación.
- Incidencia heredada: dos validaciones `#REF!` del maestro en
  `04_Control_Publicos` y `05_Matriz_Resultados`; no se generaron errores
  nuevos.
- SHA-256:
  `D7D08CB61FB316B7F15F374C87374D61D31080537D3860FB5752A47BA91C8F1E`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260827_1044.md`.

## Rectificación ETW y control de campañas públicas — 2026-08-26 14:37 CEST

- Criterio vigente: `Windows.ETW.Monitoring` tiene campaña `CERRADA` y
  `CONCLUYENTE`.
- Resultado observado: 0 filas TEC, 0 filas FP, 0 detecciones específicas y
  0 alertas en la configuración ejecutada.
- Interpretación: resultado negativo concluyente limitado a esa configuración;
  no demuestra que ETW ni el artifact fallen con carácter general.
- Las seis campañas de artifacts públicos están cerradas: cinco acreditan
  visibilidad positiva y ETW aporta el resultado negativo concluyente anterior.
- Solo Hayabusa Monitoring CH acredita detección específica: 3/9 técnicas.
- Se actualizaron `08_MEMORIA/Ultima_validacion/Ultima_iteracion.docx`, su PDF y
  `ULTIMAS_CORRECCIONES.docx` para imponer este criterio en REV-002, REV-013,
  REV-027 y el checklist final.
- Los maestros `08_MEMORIA/TFM.docx` y
  `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx` todavía requieren que el
  autor aplique esas sustituciones. El benchmark no contiene una valoración de
  cierre ETW que corregir.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260826_1437.md`.

## Guía de aplicación de correcciones — 2026-08-25 14:57 CEST

- Entregable: `08_MEMORIA/Ultima_validacion/ULTIMAS_CORRECCIONES.docx`.
- Alcance: 66/66 hallazgos de `Ultima_iteracion.pdf`, de `REV-001` a
  `REV-066`, con localización y corrección concreta.
- Formato: 14 páginas, agrupación por prioridades y estilo visual coherente con
  el informe de auditoría.
- QA: secuencia completa, auditoría OOXML, 0 incidencias altas de accesibilidad
  y revisión visual 14/14.
- SHA-256: `CED832C236145E12AAF1B6CF7D26B7251976AC7D57C71A320D8763F0893D9EC8`.
- `08_MEMORIA/TFM.docx` y `Ultima_iteracion.pdf` permanecen sin cambios.
- Estado: guía cerrada; aplicación de correcciones pendiente del autor sobre
  una copia de trabajo.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260825_1457.md`.

## Estado global

Revisión cohesionada del capítulo VII cerrada el 18/08/2026:

- Entregable vigente para revisión:
  `08_MEMORIA/ENTREGA/TFM_CAPITULO_7_BENCHMARK_RENDIMIENTO_CERRADO_7_6_COHESIONADO.docx`.
- El original de 12 páginas se conserva sin cambios.
- 15 páginas, 3.936 palabras, 26 encabezados, 10 tablas y 5 figuras.
- 7.5 sintetiza exclusivamente el coste operativo y 7.6 relaciona detección y
  coste sin mezclar unidades ni experimentos.
- Conclusión: balance favorable observado para Velociraptor custom en el
  laboratorio —9/9 técnicas, CPU media 0,95 % bajo runner y memoria próxima a
  57 MB—, sin afirmar superioridad universal ni eficiencia frente a Wazuh.
- Wazuh se compara por cobertura: base 3/9, custom 4/9 y unión complementaria
  7/9; no existe benchmark homólogo de recursos.
- QA estructural, geometría, campos, accesibilidad y privacidad: `PASS`.
- Limitación de QA: render visual no disponible por falta de `pdf2image` y
  bloqueo de la exportación PDF de Word reproducible también con el original.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260818_1416.md`.

Capítulo VII independiente cerrado el 17/08/2026:

- `08_MEMORIA/ENTREGA/TFM_CAPITULO_7_BENCHMARK_RENDIMIENTO_CERRADO.docx`.
- Estado documental: `APTO` para revisión/integración posterior por el autor.
- Alcance: solo capítulo VII; no sustituye ni modifica `08_MEMORIA/TFM.docx`.
- Extensión: 12 páginas, 9 tablas y 5 figuras.
- QA: accesibilidad 0 incidencias; geometría 9/9 tablas; revisión visual 12/12.
- CPU, memoria y disco ampliados con separación explícita entre métricas de
  proceso y actividad contextual del sistema.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260817_1152.md`.

Estado de los entregables Excel:

- `APTO` - `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx` (328.177 B; SHA-256 `FDA4501CB730058F2645F2A7FF5A80D66665AE933B45474833FFCB3ED2B86775`).
- `APTO` - `08_MEMORIA/ENTREGABLE/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx`.
- Microsoft Excel COM: recálculo completo, guardado y reapertura sin reparación.
- OpenXML: 0 errores XML, 0 referencias rotas, 0 gráficos sin series y 0 errores de fórmula.
- Auditoría cruzada: 39/39 comprobaciones.
- La carpeta específica `08_MEMORIA/ENTREGABLE/ENTREGA_FINAL_EXCELES_TFM` contiene exactamente los dos Excel definitivos. La raíz conserva además la copia histórica `PRE_CUSTOM_REVIEW` y documentos auxiliares.
- Informes: `08_MEMORIA/AUDITORIA_EXCELES_FINAL_20260712`.

Ajuste final quirúrgico del libro de visibilidad cerrado el 13/07/2026:

- Gráfico de capas limitado a Visibilidad, Detección y Alerta RT; salida externa separada de CLIENT_EVENT.
- Detección específica acreditada: Velociraptor custom 9/9, Hayabusa CH 3/9, Wazuh base 3/9 y Wazuh custom 4/9.
- Seis campañas públicas cerradas: 5 con visibilidad positiva y ETW con
  resultado negativo concluyente; CHM no ejecutada.
- TrackNetwork acredita visibilidad de conexión; el HTTP 200 de TEC-009 se atribuye solo al runner independiente.
- Discord, HTTP, Router, SERVER_EVENT y transporte quedan separados de los detectores CLIENT_EVENT.
- Valores exactos `Pendiente`: 0 al abrir el maestro manual vigente y 0 al cerrar la integración Wazuh base/custom.
- Copia auxiliar de visibilidad binariamente idéntica al entregable.
- Benchmark no abierto ni guardado; conserva 255.393 B y SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.

TFM sobre evaluación de Velociraptor como HIDS/DFIR en Windows para observar,
detectar y analizar un subconjunto controlado de técnicas MITRE ATT&CK mediante
artifacts personalizados.

Estado actual de la memoria:

- `APTO PARA REVISIÓN FINAL HUMANA`.
- `NO APTO PARA PDF FINAL TODAVÍA`.
- Documento vigente para continuar:
  `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
- Documento original beta que no debe modificarse sin orden expresa:
  `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx`.

## Alcance real

- Sistema evaluado: Velociraptor en laboratorio Windows.
- Comparación auxiliar: Wazuh como HIDS/SIEM, sin equivalencia directa de
  unidades con Velociraptor.
- Técnicas evaluadas: `TEC-001` a `TEC-009`.
- No extrapolar a toda MITRE ATT&CK Enterprise.
- No extrapolar a producción.
- Codex no valida experimentalmente; la validación experimental corresponde al
  autor en laboratorio.

## Técnicas evaluadas

| TEC | MITRE ATT&CK | Escenario resumido |
|---|---|---|
| TEC-001 | T1059.001 PowerShell | Ejecución PowerShell con logging 4104. |
| TEC-002 | T1059.003 Windows Command Shell | Uso controlado de `cmd.exe`. |
| TEC-003 | T1053.005 Scheduled Task | Creación, ejecución y retirada de tarea temporal. |
| TEC-004 | T1547.001 Registry Run Keys | Valor Run temporal controlado. |
| TEC-005 | T1569.002 Service Execution | Servicio temporal controlado. |
| TEC-006 | T1518.001 Security Software Discovery | Discovery benigno de software/servicios de seguridad. |
| TEC-007 | T1486 Data Encrypted for Impact | Cifrado AES de datos dummy controlados. |
| TEC-008 | T1485 Data Destruction | Borrado controlado dentro del dataset dummy. |
| TEC-009 | T1074.001, T1560.001, T1048.003 | Staging, ZIP y exfiltración HTTP controlada en laboratorio. |

## Resultados finales consolidados

Velociraptor:

- Cobertura real documentada: 9/9 técnicas.
- Filas consolidadas `CLIENT_EVENT`: 379.
- Distribución por perfil:
  - P1 = 19.
  - P2 = 118.
  - P3 = 109.
  - P4 = 133.

Falsos positivos:

- FP-001 a FP-010 presentes.
- 10/10 OK.
- 0 WARN, 0 FAIL, 0 SKIPPED.
- `VelociraptorHits.Total = 0`.
- 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas con `vr_hits`; se
  documentan como discrepancia y no se usan como FP definitivo.

Benchmark:

- 9/9 runs válidos.
- Escenarios: `BASELINE_NO_VR`, `VR_IDLE`, `VR_TEC_RUNNER`.
- 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- CPU media en `VR_TEC_RUNNER`: 0,95 %.
- RAM media aproximada del cliente Velociraptor: 57 MB.
- Pico CPU cliente: 54 %.
- `RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER` es WARN no bloqueante.

Wazuh:

- Baseline sin reglas TFM: 902 archives y 58 alertas; no contiene marcadores TEC y no se usa como numerador de cobertura.
- Campaña TEC con ruleset completo: 18.285 archives y 110 alertas, desglosadas en 54 alertas nativas y 56 alertas 110xxx.
- Wazuh base: 3/9, TEC-002/004/006; reglas 92032, 92302, 91835 y 92077.
- Wazuh custom: 4/9, TEC-001/005/007/008; reglas 110301, 110402, 110203 y 110202.
- La unión base+custom es 7/9 como dato complementario, no una quinta solución independiente.
- Regla `110201=0` conservada como gap TEC-009.
- Wazuh se usa como comparación HIDS/SIEM, no como sustituto ni unidad
  equivalente a `CLIENT_EVENT`.

TEC-009:

- Se mantiene como exfiltración HTTP controlada en laboratorio.
- Evidencia mínima defendible en la memoria: `ReceiverReachable=True`,
  `HTTPStatus=200`, `UploadSucceeded=True`, SHA-256 local y bytes transferidos
  en `ReceiverResponse`.
- Si en el paquete revisado no aparece `receiver_log.jsonl` o ZIP recibido, se
  trata como limitación documental secundaria, no como invalidación de TEC-009.

## Artifacts públicos - campaña 12/07/2026

- Seis carpetas de campaña procesadas.
- Total TEC canónico: 516; 521 filas raw en ventana antes de deduplicación.
- Total de filas en ventanas benignas/FP: 127.
- ProcessCreation, ServiceCreation, SysmonLogForward y TrackNetwork se interpretan como visibilidad/forense según su capa; no se convierten en detección por volumen.
- Hayabusa Monitoring CH: 185 matches canónicos, 7 high, 0 critical y detección específica 3/9.
- Hayabusa FP: 4 high, 3 eventos subyacentes y un caso benigno FP-003.
- Hayabusa CHM: no ejecutada; no existe campaña Medium separada.
- ETW: campaña cerrada y concluyente; 0 filas TEC, 0 filas FP, 0 detecciones
  específicas y 0 alertas en la configuración evaluada. El resultado no se
  extrapola a otras configuraciones.
- TrackNetwork: 128 eventos canónicos; tres filas a 192.168.1.129:8088 con PID=0/ProcInfo vacío; sin atribución fiable.
- TEC-009 independiente: HTTP 200, UploadSucceeded=True, ZIP 6.245 B y SHA-256 coincidente.

## Estado de Velociraptor

Artifacts fuente de verdad:

- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`

Criterio:

- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, normaliza, persiste o notifica.
- JSONL y Discord son salidas externas, no detección primaria.
- `01_ARTIFACTS/candidate` y `01_ARTIFACTS/debug` no deben leerse por defecto.

## Estado de Wazuh

Fuente principal:

- `10_WAZUH/TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz`.
- Ruleset efectivo dentro del paquete: `ENTREGA_WAZUH_VISIBILIDAD_20260622_114027/config/tfm_wazuh_custom_rules_v2.xml`.

Evidencia documental:

- Ventanas y JSONL `base/wazuh_base_alerts_delta.jsonl`, `custom/wazuh_custom_alerts_delta.jsonl` y `custom/tfm_wazuh_custom_detections_110xxx.jsonl` dentro del paquete anterior.
- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`.
- `04_EVIDENCE/Excel_visibilidad_26062026`.

Uso metodológico:

- Comparación de capacidades HIDS/SIEM.
- No equivalencia directa entre `archives/alerts` y `CLIENT_EVENT`.
- Gap TEC-009 mantenido en regla `110201=0`.

## Estado de benchmark

Fuente final:

- `08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx`.

Trazabilidad:

- `04_EVIDENCE/Excel_benchmark_26062026`.
- Fuente canónica regenerada:
  `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS/02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`.

Interpretación:

- Benchmark de coste operativo observado en laboratorio.
- No mide calidad de detección.
- No debe usarse para afirmar rendimiento universal en producción.

## Estado de falsos positivos

Fuente final:

- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`.
- Trazabilidad en `04_EVIDENCE/Excel_visibilidad_26062026`.

Criterio:

- FP definitivo se basa en `summary.json` y `vr_hits.csv` reconciliados.
- Filas fuera de ventana, duplicadas o no reconciliadas se documentan, pero no
  inflan resultados.

## Estado de la memoria

Memoria vigente:

- `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.

Estado:

- Contenido técnico y metodológico avanzado.
- No exportar PDF final todavía.
- Bloqueantes antes de PDF:
  - cerrar Word y confirmar desaparición del temporal
    `~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`;
  - eliminar índice general duplicado;
  - actualizar índice general, tablas, figuras y código;
  - reescribir frase del capítulo VI sobre JSONL/Discord;
  - revisión visual de tablas, captions, anexos y PDF provisional.

## Rutas importantes

| Ruta | Uso |
|---|---|
| `00_CONTEXT/README_CONTEXT.md` | Puerta de entrada al contexto limpio. |
| `00_CONTEXT/OLD_CONTEXT_USELESS` | Contexto histórico/no leer por defecto. |
| `01_ARTIFACTS/validated` | Artifacts finales/validados. |
| `01_ARTIFACTS/candidate` | Candidatos históricos o no finales; no leer por defecto. |
| `02_SCRIPTS/validated` | Scripts validados de apoyo. |
| `03_RUNNERS/validate` | Runners/scripts de validación controlada. |
| `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO` | Runners antiguos y backups. |
| `04_EVIDENCE/Excel_visibilidad_26062026` | Evidencia final visibilidad/detección/FP/Wazuh. |
| `04_EVIDENCE/Excel_benchmark_26062026` | Evidencia final benchmark. |
| `06_CONTROLLED_RERUN` | Paquete reproducible de repetición controlada. |
| `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10` | Iteración vigente de memoria pre-PDF. |
| `10_WAZUH` | Material de comparación Wazuh. |
| `99_ARCHIVE` | Histórico ya archivado. |

## Fuentes de verdad

Leer primero:

1. `00_CONTEXT/README_CONTEXT.md`.
2. `00_CONTEXT/PROJECT_STATE.md`.
3. `00_CONTEXT/CURRENT_TASK.md`.
4. `00_CONTEXT/DECISIONS.md`.

Fuentes secundarias:

- `00_CONTEXT/ARTIFACT_INDEX.md`.
- `00_CONTEXT/TEST_MATRIX.md`.
- `00_CONTEXT/EVIDENCE_INDEX.md`.
- `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`.
- `00_CONTEXT/CODEX_RULES.md`.
- `00_CONTEXT/VM_ACTIVE_PATHS.md`.

No leer por defecto:

- `00_CONTEXT/OLD_CONTEXT_USELESS`.
- Cualquier carpeta `candidate`.
- `01_ARTIFACTS/debug`.
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`.
- `99_ARCHIVE`.

## Cierre Excel de visibilidad — 2026-07-13 20:45 CEST

- Maestro vigente: `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`.
- Estado: `APTO` tras revisión física custom y exclusión del benchmark.
- Tamaño/SHA-256: 311.995 B / `4877A0391822495DEC4EFF6C987964A6EC59B4D538B598860B5E6C7D508A6D0F`.
- Estructura: 30 hojas, 15 tablas y 11 gráficos; `09_Resultados_Custom` sustituye a `09_Benchmark_Plan`.
- Resultado custom: 379 alertas CLIENT_EVENT, P1=19, P2=118, P3=109, P4=133; 9/9 TEC detectadas; 10/10 FP OK y 0 hits.
- Comparativa: Hayabusa CH=3/9, Wazuh base=3/9 y Wazuh custom=4/9.
- Validación OpenXML, Excel COM y reapertura sin reparación: `PASS`.
- Benchmark intacto: 255.393 B; SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260713_2045.md`.

## Paquete mínimo de entrega — 2026-07-14 13:37 CEST

- Carpeta: `08_MEMORIA/ENTREGABLE/ENTREGA_FINAL_EXCELES_TFM`.
- Contenido: exactamente dos archivos `.xlsx`, sin documentos auxiliares ni subcarpetas.
- Benchmark: 255.393 B; SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.
- Visibilidad: 311.995 B; SHA-256 `4877A0391822495DEC4EFF6C987964A6EC59B4D538B598860B5E6C7D508A6D0F`.
- Las copias son binarias idénticas a los maestros definitivos; no se regrabaron los libros.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260714_1337.md`.

## Entrega final completa — 2026-08-04 12:51 CEST

- Estado: `APTO PARA ENTREGA`.
- Carpeta final: `08_MEMORIA/ENTREGA/FINAL_PARA_SUBIR`.
- Contenido: exactamente un DOCX, su PDF exportado desde Word y los dos Excel
  definitivos.
- Word/PDF: 177 páginas; 68 tablas, 52 figuras y 32 códigos con campos SEQ
  consecutivos; cero líneas genéricas `Fuente: elaboración propia`.
- Detección: 379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones
  repetidas; no equivalen a 379 alertas únicas. Se excluyen 82 filas no
  reconciliadas del cómputo definitivo de FP.
- Cobertura contrastada: Velociraptor custom 9/9, Hayabusa 3/9, Wazuh base
  3/9, Wazuh custom 4/9 y unión complementaria Wazuh 7/9.
- Benchmark: 9 runs válidos, tres escenarios y tres repeticiones; SERVER_GUI
  excluido de las métricas principales del cliente.
- Rutas personales, IP privadas y URL de webhook neutralizadas en los
  entregables finales sin modificar métricas.
- Informe: `08_MEMORIA/ENTREGA/INFORME_VALIDACION_FINAL.txt`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260804_1251.md`.

## Corrección independiente del capítulo II — 2026-08-05 16:27 CEST

- Original preservado: `Planificación.docx`.
- Entregables: `Planificación_CORREGIDA.docx`,
  `Planificación_CORREGIDA.pdf` e `INFORME_PLANIFICACION.txt`.
- Tabla 2 y tablas 2.4.1-2.4.11: coincidencia exacta.
- Totales: 150 h estimadas, 290 h reales y 205 h reales en V-VIII.
- Tabla XI reconstruida sin contenido duplicado del capítulo X.
- Imágenes: 9/9 binariamente idénticas; los 8 Gantt no se modificaron.
- DOCX/PDF: 14 páginas.
- Auditoría: 20 PASS, 0 FAIL; revisión visual 14/14.
- Estado: `APTO`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260805_1627.md`.

## Segunda corrección de tablas de planificación — 2026-08-05 19:23 CEST

- Entregables actualizados: `Planificación_CORREGIDA.docx`,
  `Planificación_CORREGIDA.pdf` e `INFORME_PLANIFICACION.txt`.
- Las tablas 2.4.1 a 2.4.11 reproducen los epígrafes inmediatos exactos del
  índice facilitado por el autor.
- No se inventaron apartados 7.6, 10.5 ni 11.x.
- Tabla 2 y tablas detalladas: coincidencia exacta.
- Totales: 150 h estimadas, 290 h reales y 205 h reales en V-VIII.
- Imágenes: 9/9 binarias intactas; Gantt modificados: 0.
- DOCX/PDF: 14 páginas.
- Auditoría: 72 PASS, 0 FAIL; revisión visual 14/14.
- Incidencia: `Planificación.docx` no estaba presente al iniciar esta segunda
  revisión. No se recreó ni sobrescribió.
- Estado de coherencia documental: `APTO`.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260805_1923.md`.

## Revisión final de los originales actuales — 2026-08-24 21:27 CEST

- Fuente revisada: `08_MEMORIA/TFM.docx` y los dos Excel maestros indicados por
  el autor, todos en modo de solo lectura.
- Informe definitivo: `08_MEMORIA/Ultima_validacion/Ultima_iteracion.docx` y
  `Ultima_iteracion.pdf`.
- Resultado: `62/100`; estado `NO ENTREGAR TODAVÍA`.
- Hallazgos: 66 cambios trazables (`REV-001`–`REV-066`), con 16 críticos.
- Control: DOCX/PDF equivalentes, 39 páginas, 66/66 REV y revisión visual 39/39.
- Los hashes de los tres originales permanecen sin cambios.
- Este dictamen se aplica al `08_MEMORIA/TFM.docx` actual de 24/08/2026 y no
  reescribe el estado histórico de paquetes derivados anteriores.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260824_2127.md`.

## Auditoría de prioridad del maestro Excel — 2026-08-28 09:12 CEST

- Libros auditados en modo de solo lectura:
  `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx` y
  `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor_FINAL.xlsx`.
- Las métricas experimentales principales permanecen alineadas con el maestro.
- Estado de trazabilidad: corrección pendiente. Se localizaron 25 celdas en las
  que cuatro nombres/rutas de CSV reales fueron sustituidos indebidamente por
  nombres derivados de hojas; las rutas introducidas no existen.
- Las divergencias ID 26 y 521/516 proceden de contradicciones internas del
  propio maestro; FINAL aplicó las evidencias específicas y definitivas: ID 26
  no observado y 516 filas públicas.
- Los Excel no fueron modificados durante esta auditoría.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260828_0912.md`.

## Última validación independiente del Excel de visibilidad — 2026-08-28 13:28 CEST

- Archivo auditado sin modificar:
  `08_MEMORIA/Ultima_validacion/Analisis_Tecnicas_TFM_Velociraptor.xlsx`.
- Dictamen: `NO-GO — CORREGIR ANTES DE ENTREGAR`; nota `6,2/10`.
- Integridad positiva: 0 errores de celda, 0 XML inválidos, 0 vínculos
  externos, 0 series rotas y 238/238 hashes de evidencia coincidentes.
- Coherencia positiva: ETW 143/16 y POSITIVO, ID 26=0/no observado, TEC-009
  correctamente estratificada, custom 379 filas/9 de 9, FP=0 y Wazuh
  base/custom 3/9 y 4/9.
- Bloqueantes: referencias internas `ChatGPT` en el tema OOXML, 238 fechas
  mostradas como seriales y maquetación de impresión no finalizada con 241
  páginas.
- SHA-256 antes y después:
  `0EA74CBF5FEF395662D4C5B28BDEE4E3E0BFDEB59284271B90E23DFBCA1C99E5`.
- El benchmark no formó parte de esta revisión.
- Detalle: `00_CONTEXT/SESSION_LOGS/session_20260828_1328.md`.
