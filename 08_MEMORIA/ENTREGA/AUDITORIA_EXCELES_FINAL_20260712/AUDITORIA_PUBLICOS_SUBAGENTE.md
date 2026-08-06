# Auditoría documental de artifacts públicos

Generado: `2026-07-12T15:46:19.959352Z`.

## Alcance y criterio

Se inspeccionaron los contenidos reales de CSV, JSON/JSONL, TXT y LOG. Los formatos CSV y JSONL son serializaciones alternativas del mismo export y no se suman. Los conteos defendibles se filtran por las ventanas UTC documentadas. Los datos crudos no se modificaron.

Separación aplicada: evento crudo = visibilidad; coincidencia Sigma/Hayabusa = detección; `_ts` durante la campaña = evidencia de entrega temporal; HTTP 2xx/`UploadSucceeded`/respuesta del receptor = salida externa independiente; Sysmon ID 26 = forense; FP exige una detección injustificada, no mera telemetría.

## Inventario

- Ficheros: **112**.
- Bytes: **12473645**.
- Imágenes: **0** (no hay PNG/JPG que inventariar).
- Campañas esperadas ausentes: **ninguna**.
- No existe una campaña Hayabusa CHM ni una campaña Medium separada; no se crea ni se infiere.

| Campaña | Artifact | Ventana TEC UTC | Ventana FP UTC | Export REAL | Canónico TEC (raw) | Export FP | Canónico FP (raw) | Veredicto |
|---|---|---|---|---:|---:|---:|---:|---|
| 01_ProcessCreation | `Windows.Events.ProcessCreation` | 2026-07-12T11:50:43.356001Z — 2026-07-12T11:58:00.216547Z | 2026-07-12T12:01:47.092405Z — 2026-07-12T12:02:15.729249Z | 408 | 89 (89) | 464 | 18 (18) | `POSITIVO_VISIBILIDAD` |
| 02_ServiceCreation | `Windows.Events.ServiceCreation` | 2026-07-12T12:26:12.997533Z — 2026-07-12T12:33:05.540017Z | 2026-07-12T12:37:57.099829Z — 2026-07-12T12:38:39.571197Z | 1 | 1 (1) | 1 | 0 (0) | `POSITIVO_VISIBILIDAD_TEC005_7045` |
| 03_SysmonLogForward | `Windows.Sysinternals.SysmonLogForward` | 2026-07-12T12:46:16.670076Z — 2026-07-12T12:52:34.859019Z | 2026-07-12T12:55:41.452035Z — 2026-07-12T12:56:21.063587Z | 123 | 113 (113) | 157 | 24 (24) | `POSITIVO_VISIBILIDAD_Y_FORENSE` |
| 04_HayabusaMonitoring_CH | `Windows.Hayabusa.Monitoring` | 2026-07-12T13:25:03.318484Z — 2026-07-12T13:31:18.640618Z | 2026-07-12T13:41:00.378004Z — 2026-07-12T13:41:29.197245Z | 240 | 185 (187) | 409 | 73 (73) | `POSITIVO_DETECCION_CON_ANOMALIA_DE_NIVELES` |
| 05_ETWMonitoring | `Windows.ETW.Monitoring` | 2026-07-12T14:22:25.782443Z — 2026-07-12T14:30:04.869263Z | 2026-07-12T14:32:20.307013Z — 2026-07-12T14:32:57.518718Z | 0 | 0 (0) | 0 | 0 (0) | `NO_CONCLUYENTE` |
| 06_TrackNetworkConnections | `Generic.Events.TrackNetworkConnections` | 2026-07-12T14:43:12.204651Z — 2026-07-12T14:50:12.514120Z | 2026-07-12T14:52:39.396356Z — 2026-07-12T14:53:04.206073Z | 190 | 128 (131) | 314 | 12 (12) | `POSITIVO_VISIBILIDAD_TEC009_CON_UPLOAD_CONFIRMADO_POR_RUNNER` |

## Resumen cuantitativo defendible

| Artifact | Técnicas visibles | Técnicas detectadas | Alerta RT | Salida externa del artifact | FP confirmados |
|---|---:|---:|---|---|---:|
| Windows.Events.ProcessCreation | 9/9 (proceso correlacionado) | 0/9 | No | No | 0 |
| Windows.Events.ServiceCreation | 1/9 (TEC-005, 7045) | 0/9 | No | No | 0 |
| Windows.Sysinternals.SysmonLogForward | 9/9 | 0/9 | No | No | 0 |
| Windows.Hayabusa.Monitoring CH | N/A (salida de reglas) | 3/9 | Sí, con latencia medida | No | 4 matches / 1 caso |
| Windows.ETW.Monitoring | 0 coincidencias, NO CONCLUYENTE | 0 acreditadas | No acreditada | No | No evaluable |
| Generic.Events.TrackNetworkConnections | 1/9 (TEC-009) | 0/9 | No | No; upload confirmado por runner independiente | 0 |

## Resultados por campaña

### 01_ProcessCreation — `Windows.Events.ProcessCreation`

Tipo: `CLIENT_EVENT`. Rol: telemetría de creación de procesos; visibilidad.

Conteo canónico: REAL 89 (89 filas raw) en ventana de 408 exportadas; FP 18 (18 raw) de 464 exportadas.

Procesos de interés en la ventana TEC: `cmd.exe`=39, `powershell.exe`=21, `reg.exe`=4, `sc.exe`=5, `schtasks.exe`=5, `wmic.exe`=1.

Hay telemetría temporalmente asociada a las nueve técnicas cuando aparece el proceso esperado, pero no se considera ninguna técnica detectada solo por el proceso.

Anomalías/limitaciones:

- El export REAL contiene filas fuera de la ventana TEC; se excluyen del conteo defendible.
- El export FP contiene filas fuera de la ventana FP; se excluyen del conteo defendible.
- El runner FP documenta warning de JSONL custom; JSONL no se usa como detección del artifact público.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

### 02_ServiceCreation — `Windows.Events.ServiceCreation`

Tipo: `CLIENT_EVENT`. Rol: telemetría de creación de servicios (System 7045); visibilidad.

Conteo canónico: REAL 1 (1 filas raw) en ventana de 1 exportadas; FP 0 (0 raw) de 1 exportadas.

Se observa `7045` a `2026-07-12T12:28:47.050509Z`, servicio `TFM TEC005 Test Service`, cuenta `LocalSystem`. `StartExit=1053` es el fallo esperado del wrapper `cmd.exe` y no invalida la creación/7045.

Anomalías/limitaciones:

- El export FP contiene filas fuera de la ventana FP; se excluyen del conteo defendible.
- El JSON FP repite el evento 7045 de la ventana TEC; queda fuera de FP. El CSV FP solo conserva _ts.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

### 03_SysmonLogForward — `Windows.Sysinternals.SysmonLogForward`

Tipo: `CLIENT_EVENT`. Rol: reenvío de telemetría Sysmon; visibilidad y evidencia forense.

Conteo canónico: REAL 113 (113 filas raw) en ventana de 123 exportadas; FP 24 (24 raw) de 157 exportadas.

IDs solicitados en TEC: `1`=75, `3`=12, `11`=21, `12`=1, `13`=4, `14`=0, `26`=0.

Sysmon ID 26: **0** filas; se clasifica exclusivamente como evidencia forense, no como alerta individual.

Anomalías/limitaciones:

- El export REAL contiene filas fuera de la ventana TEC; se excluyen del conteo defendible.
- El export FP contiene filas fuera de la ventana FP; se excluyen del conteo defendible.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

### 04_HayabusaMonitoring_CH — `Windows.Hayabusa.Monitoring`

Tipo: `CLIENT_EVENT`. Rol: coincidencias de reglas Sigma/Hayabusa; detección por regla.

Conteo canónico: REAL 185 (187 filas raw) en ventana de 240 exportadas; FP 73 (73 raw) de 409 exportadas.

Niveles únicos REAL: `critical`=0, `high`=7, `informational`=92, `low`=38, `medium`=48. Coincidencias Critical/High: **7**.

La presencia de `medium/low` dentro del export etiquetado CH es una anomalía del paquete; no se transforma en una campaña Medium. El estado `Stable` y los parámetros exactos no aparecen en las filas.

Detección CH semántica: **3/9** técnicas (`TEC-001, TEC-003, TEC-005`), **6** coincidencias de regla; el total High bruto es 7 y Critical 0.

Entrega temporal CH: 7 coincidencias con latencia `_ts - Timestamp` de 38.538 a 207.008 s (mediana 72.004 s); todas se emitieron dentro de la campaña TEC documentada: `True`.

FP confirmado CH: **4** coincidencias High/Critical, **3** eventos subyacentes y **1** caso afectado (`FP-003`); Critical=0, High=4.

Anomalías/limitaciones:

- El export REAL contiene filas fuera de la ventana TEC; se excluyen del conteo defendible.
- El export FP contiene filas fuera de la ventana FP; se excluyen del conteo defendible.
- CSV/JSON REAL difieren en filas lógicas: CSV=237, JSONL=240.
- El export de la carpeta CH contiene niveles medium/low; no constituye una campaña Medium separada y no debe ocultarse.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

### 05_ETWMonitoring — `Windows.ETW.Monitoring`

Tipo: `CLIENT_EVENT`. Rol: monitorización ETW con reglas Sigma.

Conteo canónico: REAL 0 (0 filas raw) en ventana de 0 exportadas; FP 0 (0 raw) de 0 exportadas.

**Los exports ETW están vacíos durante ventanas documentadas, pero el paquete no conserva runner/log/summary TEC ni manifiesto de parámetros/exclusiones. Al no poder acreditar ejecución TEC correcta y configuración aplicada, el cero no se eleva a negativo válido.**

No se acreditan en el paquete las exclusiones predeterminadas concretas; por tanto no se enumeran como configuración observada.

Anomalías/limitaciones:

- Los cuatro exports ETW son de 0 bytes; el nombre CSV REAL incluye un guion extra. No hay cabeceras ni manifiesto de parámetros.
- No existe TFM_TEC runner/log/summary dentro de la carpeta ETW; solo se conserva runner FP. El resultado TEC queda NO CONCLUYENTE.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

### 06_TrackNetworkConnections — `Generic.Events.TrackNetworkConnections`

Tipo: `CLIENT_EVENT`. Rol: telemetría diferencial de conexiones; visibilidad.

Conteo canónico: REAL 128 (131 filas raw) en ventana de 190 exportadas; FP 12 (12 raw) de 314 exportadas.

Conexiones objetivo 192.168.1.129:8088 observadas: **3** filas add/remove. Las filas tienen `Pid=0`, `ProcInfo` vacío y timestamp crudo 1601; se usa `_ts` y no se atribuyen a PowerShell.

Runner TEC-009 independiente: HTTP `200`, `UploadSucceeded=True`, ZIP `6245` bytes, SHA-256 `247f688837e109355e49d3242f323a6fcb34c90c464db290fefba7c68e6e87dc`, receptor `received_public_final\upload_20260712_144827_450458.zip`.

`Test-NetConnection=True` no se usa como prueba de POST. La salida externa se confirma por HTTP 2xx, `UploadSucceeded` y `ReceiverResponse`, no por el artifact de conexiones.

| `_ts` UTC | Estado | Local | Remoto | PID | Timestamp crudo | En TEC |
|---|---|---|---|---:|---|---|
| 2026-07-12T14:44:14.034000Z | added | 192.168.1.143:51334 | 192.168.1.129:8088 | 0 | 1601-01-01T00:00:00Z | Sí |
| 2026-07-12T14:46:09.305000Z | removed | 192.168.1.143:51334 | 192.168.1.129:8088 | 0 | 1601-01-01T00:00:00Z | Sí |
| 2026-07-12T14:48:20.793000Z | added | 192.168.1.143:51356 | 192.168.1.129:8088 | 0 | 1601-01-01T00:00:00Z | Sí |
| 2026-07-12T14:50:29.852000Z | removed | 192.168.1.143:51356 | 192.168.1.129:8088 | 0 | 1601-01-01T00:00:00Z | No |

En CSV y JSONL las cuatro filas objetivo conservan Timestamp=1601-01-01T00:00:00Z; _ts aporta las observaciones 14:44:14.034 added, 14:46:09.305 removed, 14:48:20.793 added y 14:50:29.852 removed. La última queda 17,338 s fuera de TEC.

Anomalías/limitaciones:

- El export REAL contiene filas fuera de la ventana TEC; se excluyen del conteo defendible.
- El export FP contiene filas fuera de la ventana FP; se excluyen del conteo defendible.
- CSV/JSON FP difieren en filas lógicas: CSV=278, JSONL=314.
- Las conexiones TCP objetivo tienen Timestamp=1601-01-01, Pid=0 y ProcInfo vacío; se usa _ts para tiempo de observación y no se atribuye proceso.
- La auditoría es documental/local y no sustituye la validación experimental del autor.
- La configuración exacta del artifact no se conserva en el paquete; no se infieren parámetros no acreditados.
- Los nombres/rutas/marcadores del laboratorio solo sirven para correlación del ensayo, no como IOC generalizable.

## Fuentes principales y SHA-256

| Campaña | Rol | Fuente exacta | Bytes | SHA-256 |
|---|---|---|---:|---|
| 01_ProcessCreation | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_REAL.json` | 125747 | `4815aa3bd045e08c4a059aa6c6685c27c2604268a9c6c694cef8e1415867aced` |
| 01_ProcessCreation | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_REAL.csv` | 77666 | `0ac92ce526fce9e2817469a7bbd54819bb2bdafb447feadce06aca6e73f7976a` |
| 01_ProcessCreation | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_FP.json` | 143980 | `863a72d901277082157d2e13ff06c417c1fc59eaed02de60bd0cc72456d1b5de` |
| 01_ProcessCreation | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\Windows.Events.ProcessCreation_FP.csv` | 89205 | `d6617cf8372fccebd74239138822f83e3a106c1730661d84c99776714fa017b1` |
| 01_ProcessCreation | runner_tec_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\TFM_TEC_Run_CANDIDATE_v5_20260712_135043_summary.json` | 64160 | `e561e1cecaa0077ae99de008b8cae3e386eae5da90fedb841e676f92db530355` |
| 01_ProcessCreation | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\01_ProcessCreation\TFM_FP_20260712_140147_summary.json` | 9480 | `2fd97a90c24f5f285324ed0982c70ebfaaa2c6aaea5e4985c82a77651639cf70` |
| 02_ServiceCreation | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_REAL.json` | 1114 | `0e88d5a30f839ab41112dcc691ab2244db20af72977bb33f0a3bd1d9bf4b8e88` |
| 02_ServiceCreation | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_REAL.csv` | 1258 | `3ceb993c20a7394f40980c743f34ca81556e71c806c851bf5a513d59c74972f3` |
| 02_ServiceCreation | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_FP.json` | 1114 | `0e88d5a30f839ab41112dcc691ab2244db20af72977bb33f0a3bd1d9bf4b8e88` |
| 02_ServiceCreation | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\Windows.Events.ServiceCreation_FP.csv` | 18 | `cef90bf57b1e9f6df655046cbd63ea10ebded736db09dd1b413ee402714ee0e2` |
| 02_ServiceCreation | runner_tec_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\TFM_TEC_Run_CANDIDATE_v5_20260712_142613_summary.json` | 63789 | `d4635b06f60cf31e9b5994e1c5ce2a4f83dc5e8b06a4a15cd6a6150c4b85c7ac` |
| 02_ServiceCreation | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\02_ServiceCreation\TFM_FP_20260712_143757_summary.json` | 9373 | `0e78afe89b60b09267d78052580956bab627cf71c47077ef4ca64b07595ac04c` |
| 03_SysmonLogForward | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_REAL.json` | 127567 | `74024007455405d7791a081dd20b32babaff3d8b6fddf72d72fb81821c7218ed` |
| 03_SysmonLogForward | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_REAL.csv` | 138172 | `4124a68dd252adbb1e348bd36f669effe989bec190bff96068e7fbf700a4a1aa` |
| 03_SysmonLogForward | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_FP.json` | 161909 | `ef9bd911b884ea6667d2cc63d577bf263cb29f60a2e83142adef517c5ec8d3f9` |
| 03_SysmonLogForward | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\Windows.Sysinternals.SysmonLogForward_FP.csv` | 175331 | `a04627d54e7efcf0dcdb495d78df1ccb1c07d2070a4009e4c0c8f5746ec72990` |
| 03_SysmonLogForward | runner_tec_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\TFM_TEC_Run_CANDIDATE_v5_20260712_144616_summary.json` | 63672 | `a97a921a41cd3af02b8057fb2b828aee7ddf4b25edc506a0e3a56078c312dca5` |
| 03_SysmonLogForward | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\03_SysmonLogForward\TFM_FP_20260712_145541_summary.json` | 9369 | `b8f43c80309b8d88bfab5de4d74db6c64f949c13e126f718ce8363c92fbe3c4d` |
| 04_HayabusaMonitoring_CH | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_REAL.json` | 799285 | `4e49bcdaca5dbe540b71b2d1172500c0ebb79d744aa2cc3c3bbd3484ff6cbc06` |
| 04_HayabusaMonitoring_CH | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_REAL.csv` | 832227 | `b506e51b1f07b9a5d090a769a6ddc320e735d170577c4754836bb4c066f2934e` |
| 04_HayabusaMonitoring_CH | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_FP.json` | 3625078 | `96c1ceedf144266ef319c1cc9a12e8698e27e837d7cb1c2ab2841281fbf40a51` |
| 04_HayabusaMonitoring_CH | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\Windows.Hayabusa.Monitoring_FP.csv` | 3658383 | `e13c49463b1897b61e21e3d4de023fd88e67278bb7a2d33acf591b9b846155a4` |
| 04_HayabusaMonitoring_CH | runner_tec_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\TFM_TEC_Run_CANDIDATE_v5_20260712_152503_summary.json` | 63549 | `6c8e30f528a06ed41d5c91b5a575f2ee6d03123eba29056e6cc17651ad1aeb27` |
| 04_HayabusaMonitoring_CH | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\04_HayabusaMonitoring_CH\TFM_FP_20260712_154100_summary.json` | 9369 | `f2d3669760831d4c8b188d99d1d36ac6326d544fb3ee215998657d6b42fa4534` |
| 05_ETWMonitoring | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\Windows.ETW.Monitoring_REAL.json` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| 05_ETWMonitoring | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\Windows.ETW.Monitoring-_REAL.csv` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| 05_ETWMonitoring | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\Windows.ETW.Monitoring_FP.json` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| 05_ETWMonitoring | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\Windows.ETW.Monitoring_FP.csv` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| 05_ETWMonitoring | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\05_ETWMonitoring\TFM_FP_20260712_163220_summary.json` | 9370 | `26a57d36eb264f94e8fa9c16a072e4439d4f8c1e7a649da0d2f9a740c2ad9819` |
| 06_TrackNetworkConnections | real_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\Generic.Events.TrackNetworkConnections_REAL.json` | 204323 | `93bd02a45ad38cbedbf68e3ba753e1871fc4e9647a6cf3dbdacb8e1640252ab1` |
| 06_TrackNetworkConnections | real_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\Generic.Events.TrackNetworkConnections_REAL.csv` | 217976 | `8c4966bbba4cc0e3a2ccde9efe4b477c93e331cb1b6752f7e6f7c98d50a45c76` |
| 06_TrackNetworkConnections | fp_json | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\Generic.Events.TrackNetworkConnection_FP.json` | 328865 | `dd3f6c4d12d00776f6a17a0468791e96e7496a39ba7e18dafd01b43309a172e8` |
| 06_TrackNetworkConnections | fp_csv | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\Generic.Events.TrackNetworkConnections_FP.csv` | 312837 | `03b9310bf98b0c9bf80e9c9eb9b20f5d01701e0e72fbb1d38d6aa23fa7f026f8` |
| 06_TrackNetworkConnections | runner_tec_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\TFM_TEC_Run_CANDIDATE_v5_20260712_164312_summary.json` | 63548 | `b108d08ab174888caa637e3e890955fb010cbd4efa80812b92261eb3e5f1fb8b` |
| 06_TrackNetworkConnections | runner_fp_summary | `04_EVIDENCE\PRUEBAS_CON_ARTIFACTS_PUBLICOS\06_TrackNetworkConnections\TFM_FP_20260712_165239_summary.json` | 9372 | `a4c4347d6c83e26ddbd23facaed59603deb804bfb92165c8fa17f8fc2342ea40` |

El JSON de auditoría contiene los 112 hashes, incluidos TXT, LOG, ventanas y ficheros auxiliares.

## Hayabusa: detalle Critical/High por técnica

| TEC | Coincidencias semánticas todos los niveles | Critical/High | Detectada CH | Títulos CH (máx. 5) |
|---|---:|---:|---|---|
| TEC-001 | 6 | 2 | Sí | Base64 Encoded PowerShell Command Detected 2 (2); Non Interactive PowerShell Process Spawned 2 (2); Change PowerShell Policies to an Insecure Level 2 (2) |
| TEC-002 | 0 | 0 | No |  |
| TEC-003 | 5 | 1 | Sí | Schedule Task Creation From Env Variable Or Potentially Suspicious Path Via Schtasks.EXE 2 (2); Scheduled Task Created - FileCreation (1); Scheduled Task Creation Via Schtasks.EXE 2 (1); Suspicious Schtasks Schedule Types 2 (1) |
| TEC-004 | 5 | 0 | No | Autorun Keys Modification 2 (2); Potential Persistence Attempt Via Run Keys Using Reg.EXE 2 (1); Direct Autorun Keys Modification 2 (1); CurrentVersion Autorun Keys Modification 2 (1) |
| TEC-005 | 6 | 3 | Sí | New Service Creation 2 (1); New Service Creation Using Sc.EXE 2 (1); Suspicious New Service Creation 2 (1); Service Binary in Suspicious Folder 2 (1); Suspicious Service Path (1) |
| TEC-006 | 3 | 0 | No | Potential Product Reconnaissance Via Wmic.EXE 2 (1); Potential Product Class Reconnaissance Via Wmic.EXE 2 (1); System Information Discovery Via Wmic.EXE 2 (1) |
| TEC-007 | 0 | 0 | No |  |
| TEC-008 | 0 | 0 | No |  |
| TEC-009 | 4 | 0 | No | PowerShell Script With File Upload Capabilities (1); Usage Of Web Request Commands And Cmdlets - ScriptBlock (1); Compress-Archive Cmdlet Execution (1); Network Connection Initiated By PowerShell Process 2 (1) |

## TEC-009: ejecuciones documentadas

| Campaña | Run UTC | HTTP | UploadSucceeded | ZIP bytes | SHA-256 | Receptor |
|---|---|---:|---|---:|---|---|
| 01_ProcessCreation | 2026-07-12T11:55:35.937805Z — 2026-07-12T11:55:49.582799Z | 200 | True | 6245 | `cd7dff71cd0726ba5c04ebf9dba264a1f4d34ce2710d7cf70617860cd29611b6` | received_public_final\upload_20260712_115601_726038.zip |
| 02_ServiceCreation | 2026-07-12T12:30:56.958775Z — 2026-07-12T12:31:21.687373Z | 200 | True | 6245 | `e67bd7943a0d0435db8b7e345dcb20af00e5cc49ca5ce66a8f4899ef9cc36422` | received_public_final\upload_20260712_123136_272943.zip |
| 03_SysmonLogForward | 2026-07-12T12:50:47.886833Z — 2026-07-12T12:51:14.012321Z | 200 | True | 6245 | `c6fd50dd508e56c7a1351919c8375d28f88a1fe2eda1da4e4bb19cbbdc404d17` | received_public_final\upload_20260712_125109_313993.zip |
| 04_HayabusaMonitoring_CH | 2026-07-12T13:29:18.169259Z — 2026-07-12T13:29:36.928013Z | 200 | True | 6245 | `d9ab4747658450604855eaf8209d801f1231a627694fcf150eff0a33497b725c` | received_public_final\upload_20260712_132948_189878.zip |
| 06_TrackNetworkConnections | 2026-07-12T14:48:08.791122Z — 2026-07-12T14:48:24.288296Z | 200 | True | 6245 | `247f688837e109355e49d3242f323a6fcb34c90c464db290fefba7c68e6e87dc` | received_public_final\upload_20260712_144827_450458.zip |

Las cinco ejecuciones TEC-009 documentadas son independientes; ETW no conserva runner TEC. Para la campaña de conexiones se usa la fila correspondiente a `06_TrackNetworkConnections`; las demás no se suman como detecciones del artifact.

## Integridad y trazabilidad

Se calcularon SHA-256 para los **112** ficheros. El inventario completo con ruta absoluta, tamaño, timestamp y hash está en `AUDITORIA_PUBLICOS_DETALLE.json`.

Todos los CSV/JSON/TXT/LOG se abrieron y escanearon. Los JSON con objetos por línea se validaron como JSONL; los CSV se leyeron con parser CSV real, incluida multilinealidad, y no por recuento físico de líneas. Las discrepancias CSV/JSON y exports fuera de ventana constan por campaña.

## Estado

**APTO CON OBSERVACIONES** para integración documental: las seis carpetas existen y son trazables; Hayabusa exige conservar la anomalía de niveles/configuración; ETW queda **NO CONCLUYENTE** por ausencia de runner TEC y manifiesto de parámetros; TrackNetworkConnections no permite atribución de PID/proceso en las conexiones objetivo.
