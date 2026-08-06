# PROJECT_STATE

Actualizado: 2026-08-05 16:27 CEST.

## Estado global

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
- Seis campañas públicas: 5 positivas y 1 no concluyente; CHM no ejecutada.
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
- ETW: no concluyente por falta de runner/log/summary TEC y manifiesto de configuración.
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
