# VR_HIDS-EDR

Repositorio del TFM sobre evaluacion de Velociraptor como HIDS/DFIR en Windows
para observar, detectar y analizar tecnicas MITRE ATT&CK mediante artifacts
custom, evidencia local y comparacion auxiliar con Wazuh.

El proyecto separa de forma estricta cinco niveles:

- Visibilidad de telemetria.
- Deteccion emitida por artifacts `CLIENT_EVENT`.
- Enrutamiento o persistencia mediante artifacts `SERVER_EVENT`.
- Evidencia forense y evidencias externas.
- Notificaciones o salidas auxiliares.

`JSONL` y Discord no se consideran deteccion. El router no detecta; solo
normaliza, persiste o notifica eventos ya emitidos por artifacts de cliente.

## Estado del proyecto

Estado documentado a 2026-06-29:

- Evaluacion principal de Velociraptor HIDS/DFIR: preparada y documentada.
- Artifacts custom organizados por perfiles P1/P2/P3/P4.
- Excel final de visibilidad, deteccion, falsos positivos y comparacion Wazuh generado.
- Excel final de benchmark de rendimiento generado.
- Evidencias y paquetes historicos conservados como trazabilidad.
- Validacion experimental: realizada por el autor en laboratorio; Codex solo
  prepara, revisa estructura y genera derivados documentales.

Entregables principales:

- `08_MEMORIA/TFM.docx`
- `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`
- `08_MEMORIA/ENTREGABLE/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`
- `08_MEMORIA/Excel_visibilidad_FP/`
- `08_MEMORIA/Excel_benchmark/`

## Arquitectura

Flujo metodologico:

```text
Windows / Sysmon / PowerShell 4104
  -> Velociraptor CLIENT_EVENT custom
  -> Velociraptor SERVER_EVENT router
  -> JSONL / salida externa
  -> analisis posterior y memoria
```

Perfiles principales:

| Perfil | Funcion |
|---|---|
| P1 | Senales criticas de alta prioridad. |
| P2 | Deteccion y contexto forense de mayor detalle. |
| P3 | Comportamiento, normalizacion y alertabilidad. |
| P4 | Visibilidad basica de bajo coste. |
| Router SOC | `SERVER_EVENT`; enruta y normaliza, no detecta. |

Artifacts de referencia actuales:

- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
- `01_ARTIFACTS/validated/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`

## Estructura del repositorio

| Ruta | Contenido |
|---|---|
| `00_CONTEXT/` | Estado del proyecto, decisiones, matriz de pruebas, logs de sesion e indices de evidencia. |
| `01_ARTIFACTS/` | Artifacts Velociraptor `candidate`, `debug` y `validated`. |
| `02_SCRIPTS/` | Scripts de apoyo y scripts validados. |
| `03_RUNNERS/` | Runners de campana, falsos positivos, benchmark y validacion. |
| `04_EVIDENCE/` | Evidencias, Exceles, datasets normalizados y reportes legacy. |
| `05_LOGS/` | Logs y salidas brutas locales. Carpeta ignorada para GitHub. |
| `06_CONTROLLED_RERUN/` | Paquete reproducible para repeticion controlada y validacion. |
| `07_DOCS/` | Guias operativas y documentacion tecnica auxiliar. |
| `08_MEMORIA/` | Memoria, entregables finales y Exceles finales. |
| `09_LAB/` | Area reservada para elementos de laboratorio local. |
| `10_WAZUH/` | Material de comparacion con Wazuh. |
| `99_ARCHIVE/` | Iteraciones historicas, prompts antiguos y material desplazado fuera de la raiz. |

Carpetas locales no pensadas para publicacion:

- `lab/`: maquina virtual local y snapshots. No mover sin revisar VirtualBox.
- `Carpeta_Compartida_TFM/`: intercambio con VM/laboratorio.
- `99_REVIEW_CLEANUP_*/`: candidatos locales de limpieza no publicados.

## Reproducibilidad

Antes de modificar o ejecutar nada, leer:

1. `00_CONTEXT/PROJECT_STATE.md`
2. `00_CONTEXT/CURRENT_TASK.md`
3. `00_CONTEXT/DECISIONS.md`
4. `00_CONTEXT/ARTIFACT_INDEX.md`
5. `00_CONTEXT/TEST_MATRIX.md`
6. `00_CONTEXT/EVIDENCE_INDEX.md`
7. `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`
8. `00_CONTEXT/CODEX_RULES.md`
9. `00_CONTEXT/VM_ACTIVE_PATHS.md`

La repeticion controlada se documenta en `06_CONTROLLED_RERUN/README.md`.
Los resultados finales usados para memoria estan trazados en `08_MEMORIA/` y
en las carpetas derivadas de `04_EVIDENCE/`.

## Higiene GitHub

Se excluyen del repositorio publico:

- Imagenes de VM, snapshots, ISOs y paquetes grandes.
- Logs brutos, JSONL, EVTX, PCAP y subidas de laboratorio.
- OneNote, accesos directos y ficheros temporales.
- Salidas locales de `05_LOGS/` y `06_CONTROLLED_RERUN/OUTPUT/`.

La limpieza de raiz del 2026-06-29 esta documentada en
`00_CONTEXT/CLEANUP_20260629.md`.

## Limitaciones

- No ejecutar scripts fuera de un laboratorio controlado y autorizado.
- No interpretar telemetria como deteccion sin fila `CLIENT_EVENT` trazable.
- No interpretar JSONL, Discord o router como fuente primaria de deteccion.
- Sysmon ID 26 se trata como evidencia forense, no como alerta individual.
- Las metricas de benchmark describen el laboratorio usado, no un rendimiento universal.

