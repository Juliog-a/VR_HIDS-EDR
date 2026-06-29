# DECISIONS

## D001 - Arquitectura

Se mantiene arquitectura P4/P3/P2/P1 por perfiles de profundidad.

## D002 - Router

El SERVER_EVENT no detecta. Solo enruta, persiste o notifica.

## D003 - JSONL

alerts.jsonl es salida externa/persistencia, no evidencia de deteccion por si sola.

## D004 - TEC-009

TEC-009 se interpreta como staging, archivado y exfiltracion HTTP controlada de datos dummy en laboratorio. No implica datos sensibles reales.

## D005 - Sysmon ID 26

Sysmon ID 26 se usa como evidencia forense de borrado, no como alerta individual.

## D006 - IA

Codex/ChatGPT se usan como apoyo de desarrollo y documentacion. No validan experimentalmente.

## D007 - Separacion crypto/hash

En P3_v8 el hashing de integridad (`Get-FileHash`, `SHA256`, `SHA1`, `MD5`, `LocalSHA256`, `FileHash`) se separa del cifrado real. SHA256 aislado no es indicador de ransomware-like encryption.

## D008 - CU-009 fuerte

CU-009 fuerte se basa en HTTP upload de ZIP con herramienta HTTP, `InFile`, `POST` y `application/zip` o combinacion equivalente. Permite hashing de integridad y no depende de `/upload`, rutas, nombres de scripts, datos dummy ni etiquetas TEC.

## D009 - CU-009 command shape

La deteccion fuerte CU-009 no debe basarse solo en tokens sueltos. Debe exigir forma real de comando de upload HTTP ZIP y excluir ruido de validacion/detection engineering y texto de documentacion o prompt.

## D010 - CU-009 shape por coexistencia estructural

Para diagnostico v4, la forma real de CU-009 se evalua por coexistencia estructural en el mismo `ScriptText`: herramienta HTTP, URI/destino, POST, `InFile` y `application/zip`. No se usa una unica regex ordenada como criterio final, porque los bloques PowerShell multilinea con backticks pueden romper esa aproximacion. `Remove-Item` de limpieza local se registra como contexto destructivo, pero no bloquea CU-009 si el upload fuerte esta presente.

## D011 - P3_v9 candidate

Tras validacion experimental de `CU009Shape_v4` comunicada por el investigador, se crea `P3_v9` como candidate. `P3_v9` integra la forma validada en `CU009_Source_TEC009_HTTP_Upload_Strong_4104`, mantiene P3_v8 intacto y no se mueve a `validated` hasta completar campana de regresion.

## D012 - P3_v10 live watch_evtx directo

Tras fallo experimental de P3_v9, se crea `P3_v10` como candidate. La source fuerte CU-009 4104 deja de envolver `watch_evtx` en `foreach` y vuelve a un patron live directo sobre `EventData.ScriptBlockText`, igual que las sources 4104 que si emiten. Se mantiene la semantica validada por `CU009Shape_v4` y no se modifica el script de prueba.

## D013 - Separacion P3/P2 para tecnicas criticas

Se abandona la integracion de CU-009 fuerte dentro de P3. P3 queda como capa media de comportamiento general y contexto, sin responsabilidad de validar ransomware-like, sabotaje/destruccion o exfiltracion fuerte.

Se crea `P3_v11` candidate para degradar CU-007, CU-008 y CU-009 fuertes a contexto y mantener CU-001/CU-006 como senales medias.

Se crea `P2_v1` candidate como artifact especializado para CU-007, CU-008 y CU-009. La logica validada de `CU009Shape_v4` se integra en P2, no en P3.

P1 queda reservado para correlacion critica final y no se implementa todavia.

## D014 - Runner v4 para campana completa y TEC-009 real controlada

Se crea `TFM_Run_All_TEC_Tests_v4.ps1` como nueva version de runner, sin sobrescribir runners previos.

La ruta activa de ejecucion pasa a `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas`.

TEC-009 deja de usar staging inline en el runner. La ejecucion valida de TEC-009 se delega en `TEC-009_Exfiltracion\exfiltracion_v3.ps1`, lanzado desde su directorio real con `Push-Location`, `ReceiverUrl`, `EnableUpload`, `TimeoutSec` y `KeepArtifacts`.

El runner registra precheck del receiver con `Test-NetConnection`, salida de `exfiltracion_v3.ps1`, HTTP status, `Upload succeeded` y resumen de eventos PowerShell 4104 desde `runStart`.

No se modifica `exfiltracion_v3.ps1`, artifacts, router, Discord ni P1/P2/P3.

## D015 - Runner v4 con ruta absoluta y resumen estructurado

Tras revisar `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_115408.log`, se confirma que TEC-009 fallaba porque `powershell.exe -File ".\exfiltracion_v3.ps1"` se resolvia mal aunque el runner hubiese hecho `Push-Location`.

Se corrige TEC-009 para invocar siempre `exfiltracion_v3.ps1` por ruta absoluta:

`C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion\exfiltracion_v3.ps1`

El runner v4 mantiene log completo y genera resumen final TXT, JSON y CSV por tecnica en `Logs_Pruebas_TFM`.

Se mantiene `EnableExfilUpload` y `KeepExfilArtifacts` como `[bool]` para compatibilidad con los comandos ya usados (`-EnableExfilUpload $true`). No se convierten a `[switch]` para no romper ejecuciones previas.

## D016 - Script de visibilidad Sysmon/Windows

Se crea `TFM_Check_Sysmon_Visibility.ps1` como script de solo lectura. No ejecuta tecnicas, no modifica el laboratorio y no genera actividad ofensiva.

Su objetivo es documentar disponibilidad de fuentes para memoria y validacion: Sysmon, PowerShell 4104/4103, Security, System y Application. Exporta log, resumen TXT, resumen JSON y CSV de muestras si se usa `-IncludeSamples`.

## D017 - Ultima iteracion runner/scripts de campana v4

La ruta activa queda centralizada mediante `TFM_BASEPATH`.

El runner v4 establece `$env:TFM_BASEPATH = $BasePath`. Los scripts hijos deben resolver `$BasePath` desde `$env:TFM_BASEPATH` y, si no existe, usar `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas`.

TEC-007 se corrige con scripts que crean dataset dummy en el `DUMB_LAB` activo, generan `lista_archivos.csv` en la carpeta real de la tecnica y cifran ficheros dummy con AES para producir `.aesCount > 0`.

TEC-008 se corrige garantizando que el restore crea ficheros dummy antes del borrado controlado, de modo que `FilesBefore > 0`.

TEC-009 cambia `exfiltracion_v3.ps1` a parametros `[switch]$EnableUpload` y `[switch]$KeepArtifacts`. El runner invoca la tecnica con `-EnableUpload` y `-KeepArtifacts`, no con booleanos serializados por `cmd.exe`.

La consulta 4104 de TEC-009 registra `QueryError` si falla y trata `StartTime` como `[datetime]` y `Message` como `[string]`.

Se crea paquete entregable `C:\Users\julio\Desktop\TFM\ultima_iteracion` y ZIP final `C:\Users\julio\Desktop\TFM\ultima_iteracion.zip`.

## D018 - Artifacts historicos P1/P2/P3 con parse_evtx

Se crea una nueva iteracion de artifacts candidate para backtesting/hunt historico:

- `Custom.TFM.HIDS.P1.Low.Basic_v3`
- `Custom.TFM.HIDS.P2.High.Forensic_v2`
- `Custom.TFM.HIDS.P3.Medium.Behavioral_v12`

Decision:

- Usar `type: CLIENT`.
- Usar `parse_evtx()` en todas las fuentes.
- No usar `watch_evtx()` para estos artifacts.
- Mantener `watch_evtx()` solo para futuras variantes live si se documentan aparte.

## D019 - TEC-009 por 4104 fuerte y red como contexto

Para TEC-009 en P1/P2/P3 historicos:

- La deteccion principal es PowerShell 4104.
- Condicion fuerte:
  - herramienta HTTP (`Invoke-WebRequest`, `iwr`, `wget` o `curl`);
  - `-InFile`;
  - `POST`;
  - `application/zip` o `.zip` o `upload`.
- Sysmon ID 3 hacia receiver es contexto.
- Sysmon ID 3 no bloquea ni habilita por si solo la alerta fuerte.

## D020 - Cobertura 9/9 en P1/P2/P3

La nueva iteracion exige cobertura explicita para TEC-001 a TEC-009 en los tres artifacts:

- P1: bajo coste, logica simple.
- P2: detalle forense y contexto.
- P3: salida normalizada para alertabilidad.

TEC-007 y TEC-008 pueden emitir HIGH por 4104 fuerte aunque falten Sysmon ID 11 o 23/26.

## D021 - Runner v5 para 4104 fuerte en TEC-009

Se crea `TFM_Run_All_TEC_Tests_v5.ps1` sin sobrescribir v4.

Decision:

- TEC-009 se ejecuta mediante `powershell.exe -NoProfile -ExecutionPolicy Bypass -File`.
- El objetivo es favorecer que PowerShell 4104 capture el cuerpo real de `exfiltracion_v3.ps1`.
- El runner registra hora local/UTC, comando ejecutado, summary JSON, ZIP, copia ZIP, SHA256, HTTP status y `UploadSucceeded`.
- Se anade seccion final `TEC-009 EVIDENCE CHECK`.

Motivo:

- La llamada directa al script puede dejar en 4104 solo el lanzador.
- La deteccion principal de TEC-009 depende de evidencia 4104 fuerte (`Invoke-WebRequest`, `-InFile`, ZIP/upload/ReceiverUrl).

## D022 - P1/P2/P3 EVENT como linea live separada

Se crea una nueva familia de artifacts:

- `Custom.TFM.HIDS.P1.Low.Basic.Event_v1`
- `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`
- `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`

Decision:

- `type: CLIENT_EVENT`.
- Uso de `watch_evtx()`.
- Activacion antes de ejecutar runner v5.
- Cobertura 9/9.
- P3 EVENT queda como artifact principal de alertabilidad.

Justificacion:

- `CLIENT_EVENT` live debe observar eventos mientras ocurren.
- Las versiones historicas `CLIENT` con `parse_evtx()` se mantienen para backtesting.

## D023 - TEC-009 EVENT: 4104 bloquea la alerta fuerte, red no

Para TEC-009 en P1/P2/P3 EVENT:

- HIGH por 4104 si existe herramienta HTTP, carga de fichero y evidencia ZIP/upload/ReceiverUrl.
- `POST` suma evidencia, pero no es condicion obligatoria porque puede aparecer como `-Method POST`.
- Sysmon ID 1/3/11 aporta contexto y puede generar eventos MEDIUM.
- Sysmon ID 3 hacia receiver no es condicion obligatoria.

Esta decision evita que una conexion HTTP breve o perdida por polling bloquee la deteccion principal.

## D024 - Router SOC JSONL/Discord v1

Se crea `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v1` como `SERVER_EVENT`
candidate.

Decision:

- El router escucha fuentes `CLIENT_EVENT` P1/P2/P3 mediante `watch_monitoring()`.
- Normaliza las filas como alerta SOC con CU, tecnica interna, MITRE, IOA, IOC,
  severidad, confianza y accion recomendada.
- Persiste JSONL en `soc_alerts.jsonl`.
- Envia a Discord solo si `EnableDiscord=true` y `DiscordWebhook` se proporciona
  como parametro de despliegue.
- No contiene webhook hardcodeado.
- No usa `parse_evtx()` ni `watch_evtx()` porque no detecta telemetria de host.

Justificacion:

- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, persiste y notifica.
- JSONL y Discord siguen siendo salidas, no evidencia de deteccion por si solos.

## D025 - Router SOC JSONL/Discord v2

Se crea `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v2` como nueva version
candidate, sin sobrescribir `SOC_v1`.

Decision:

- Escuchar P1, P2, P3 y P4 mediante `watch_monitoring()`.
- Mantener `type: SERVER_EVENT`.
- No usar `parse_evtx()` ni `watch_evtx()` en el router.
- Usar los nombres reales observados en GUI:
  - P1 `Custom.TFM.HIDS.P1.Low.Basic.Event_v1`;
  - P2 `Custom.TFM.HIDS.P2.High.Forensic_v1`;
  - P3 `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`;
  - P4 `Custom.TFM.HIDS.P4.Low.Basic_v2`.
- Anadir `EnableP4`.
- Normalizar columnas heterogeneas:
  - P1/P3: `AlertTitle`, `Evidence`, `Source`, `RecommendedAction`;
  - P2/P4: `AlertName`, `CU_Evidence`, `Channel`, `CU_Reason`, `DetectionLogic`.
- JSONL mantiene el esquema SOC solicitado:
  `Timestamp`, `Hostname`, `ClientId`, `Artifact`, `Source`, `Profile`, `CU_ID`,
  `TEC`, `MITRE`, `IOA`, `IOC`, `Severity`, `Confidence`, `DetectionName`,
  `Evidence`, `RecommendedAction`, `RawEventSummary`.

Justificacion:

- El router `SOC_v1` no escuchaba P4.
- `SOC_v1` escuchaba P2 `High.Forensic.Event_v1`, pero la GUI mostraba P2
  desplegado como `High.Forensic_v1`.
- P2/P4 no comparten todas las columnas normalizadas de P1/P3, por lo que v2
  usa adaptadores por perfil.

## D026 - P1 Critical Priority como alias metodologico

Se crea `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` sin sobrescribir
`Custom.TFM.HIDS.P1.Low.Basic.Event_v1`.

Decision:

- Corregir naming/metadatos para que P1 represente prioridad/criticidad alta.
- Mantener los mismos source names y logica que P1 Low.Basic.Event_v1.
- Permitir que el router SOC v2 escuche P1 legacy o P1 critical.
- Evitar activar ambos simultaneamente salvo que se acepten duplicados.

## D027 - Runner v6 Evidence Check

Se crea `TFM_Run_All_TEC_Tests_v6.ps1` sin sobrescribir v5.

Decision:

- Corregir solo el bloque `TEC-009 EVIDENCE CHECK`.
- No cambiar la ejecucion de TEC-009 ni `exfiltracion_v3.ps1`.
- Devolver siempre objetos `PSCustomObject` homogeneos con:
  `CheckName`, `Status`, `Expected`, `Found`, `Evidence`, `Comment`.
- Usar `WARNING` para advertencias del evidence check.
- Mantener salida legible en consola y resumen TXT.

## D028 - Router SOC JSONL/Discord v3 por source independiente

Se crea `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3` como nueva version
candidate, sin sobrescribir `SOC_v2`.

Decision:

- Mantener `type: SERVER_EVENT`.
- Mantener `watch_monitoring()` como unica entrada del router.
- No usar `parse_evtx()` ni `watch_evtx()` en el router.
- Escuchar el P2 actualmente desplegado:
  `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`.
- Escuchar P1 Critical, P2 Event, P3 Event y P4 v2 con una source de router
  independiente por cada source real monitorizada.
- Usar defaults alineados con la validacion actual:
  - `EnableP1Critical=true`;
  - `EnableP1=false`;
  - `EnableP2Event=true`;
  - `EnableP3=true`;
  - `EnableP4=true`;
  - `ExcludeRouterSelfEvents=false`.
- Normalizar el esquema JSONL como `TFM_SOC_ALERT_v3` con:
  `Timestamp`, `DetectionTime`, `Hostname`, `ClientId`, `Artifact`,
  `Source`, `Profile`, `CU_ID`, `TEC`, `MITRE`, `IOA`, `IOC`, `Severity`,
  `Confidence`, `DetectionName`, `Evidence`, `RecommendedAction`,
  `RawEventSummary`, `SignalType`, `EventID`, `RouterArtifact`.

Justificacion:

- La prueba de `soc_alerts.jsonl` mostraba solo `P1_CRITICAL`.
- SOC_v2 usaba una unica source con muchas ramas `watch_monitoring()` dentro
  de `chain()`. En monitorizacion continua, esas ramas no son colecciones
  finitas y pueden impedir que se consuman ramas posteriores.
- SOC_v2 ademas apuntaba a P2 legacy en vez de al P2 Event v1 desplegado.

## D029 - Recalibracion P1 Critical Priority

Se recalibra `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` manteniendo el
mismo artifact name interno, el mismo nombre de fichero y los mismos source
names para no romper `SOC_v3`.

Decision:

- P1 deja de ser alias funcional de `Low.Basic`.
- P1 ya no emite `TEC-001`, `TEC-002`, `TEC-003` ni `TEC-006` genericos.
- P1 queda limitado a senales fuertes:
  - `TEC-004`: Run Key clara por Sysmon ID 1.
  - `TEC-005`: servicio sospechoso por System 7045 o Sysmon ID 1.
  - `TEC-007`: 4104 con AES + CryptoStream/CreateEncryptor + `.aes`.
  - `TEC-008`: 4104 con enumeracion + borrado + `-Force`/`-Recurse`.
  - `TEC-009`: 4104 con herramienta HTTP + POST + InFile + ZIP + upload/ReceiverUrl;
    Sysmon ID 3/11 queda solo como soporte P1 de alta fidelidad.
- Se anade exclusion fuerte de autoeventos router/JSONL/Discord.
- P2/P3/P4/SOC_v3 no se modifican.

Justificacion:

- La prueba con `SOC_v3` mostro 2117 alertas, 1938 de ellas `P1_CRITICAL`.
- `P1_CRITICAL` incluia autoeventos del pipeline SOC y actividad basica.
- La visibilidad no se elimina: queda redistribuida en P4/P3/P2.

## D030 - Runner benigno de falsos positivos

Se crea `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1`.

Decision:

- Medir falsos positivos con acciones benignas/administrativas controladas.
- No tocar runner TEC, artifacts, router, receiver ni scripts existentes.
- Usar workspace controlado:
  `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\FP_WORKSPACE\<RunId>`.
- Generar salidas en:
  `C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM\FPs`.
- Leer `soc_alerts.jsonl` y cruzar hits por ventana temporal y `RunId`.
- Clasificar hits:
  - `P1_CRITICAL`: falso positivo grave.
  - `P2_EVENT`: falso positivo forense relevante.
  - `P3`: falso positivo medio/comportamental.
  - `P4`: falso positivo leve o visibilidad esperada.

Justificacion:

- La campana ofensiva TEC ya esta validada.
- P1 fue recalibrado para reducir ruido critico.
- Hace falta una prueba separada que no simule malware ni modifique persistencia real.

## D031 - Runner FP v1.1: correccion de postproceso

Se corrige `03_RUNNERS/TFM_Run_FP_Tests_v1.ps1` manteniendo el mismo nombre de
archivo para facilitar su copia a la VM.

Decision:

- No modificar FP-001 a FP-010 salvo bugs de postproceso.
- No modificar P1/P2/P3/P4.
- No modificar `SOC_v3`.
- Corregir solo lectura de JSONL, asociacion por `RunId`/ventana temporal,
  resumen y `vr_hits`.
- Tratar JSONL vacio como caso valido con 0 hits.
- Leer JSONL linea a linea.
- Saltar lineas JSON corruptas con `WARN`.
- No abortar por timestamps no parseables.
- Usar `FP_ID=UNKNOWN` cuando no pueda inferirse.
- Forzar arrays y normalizar singletons para evitar errores de Windows PowerShell.

Justificacion:

- La campana FP `TFM_FP_20260618_144027` ejecuto las pruebas, pero fallo despues
  de esperar al JSONL por `System.ArgumentException`.
- El fallo estaba en el postproceso, no en la ejecucion de las pruebas FP.
- La evidencia disponible mantiene `P1_CRITICAL=0`; no hay motivo para tocar
  detecciones sin repetir campana con el runner corregido.

## D032 - Informe de fiabilidad y calidad de asociacion FP

Se genera `PORCENTAJE FIABILIDAD.xlsx` como informe de fiabilidad/deteccion/FP.

Decision:

- Separar datos por calidad:
  - `LOG`: procede de ficheros locales revisados.
  - `CONTEXTO_CONFIRMADO`: procede de resultados confirmados por el autor cuando
    no existe agregado local equivalente.
  - `INFERIDO`: procede de ventana temporal sin `RunId` embebido.
  - `CONTAMINADO`: evidencia asociada a un RunId anterior.
- Mantener todos los hits en `FP_HITS_RAW` para auditoria.
- No usar hits contaminados para porcentajes principales salvo como bloque
  separado.
- Mantener `P1_CRITICAL=0` y `P2_EVENT=0` de FP v1.1 como resultado principal
  LOG.
- No modificar P1/P2/P3/P4 a partir de esta muestra; documentar y ampliar muestra
  si se necesita validez estadistica.

Justificacion:

- `TFM_FP_20260618_151316_vr_hits.csv` incluye evidencias del RunId anterior
  `TFM_FP_20260618_151244`.
- La conclusion critica no cambia porque los hits contaminados observados no son
  `P1_CRITICAL` ni `P2_EVENT`, pero deben quedar separados por rigor documental.

## D033 - Benchmark de rendimiento Velociraptor

Se crea una linea separada de benchmark para medir impacto de Velociraptor en
Windows.

Decision:

- Usar tres escenarios:
  - `BASELINE_NO_VR`;
  - `VR_IDLE`;
  - `VR_TEC_RUNNER`.
- No modificar detecciones, router, receiver, runner TEC, runner FP, scripts TEC
  ni JSONL.
- Medir proceso Velociraptor y sistema por muestras:
  CPU, RAM, working set, private memory, handles, threads, IO, CPU/RAM sistema,
  disco, red y estado del runner.
- Generar una carpeta por `RunId` con:
  `samples.csv`, `process_samples.csv`, `summary.json`, `summary.txt`,
  `process_list_start.txt`, `process_list_end.txt` y `environment.txt`.
- Si no hay datos reales, el Excel queda como plantilla `PENDIENTE_DE_EJECUCION`
  y no se inventan metricas.
- La validez de escenario queda condicionada:
  - `BASELINE_NO_VR` invalido si hay procesos VR;
  - `VR_IDLE` invalido si no hay procesos VR;
  - `VR_TEC_RUNNER` invalido si no arranca runner o no hay procesos VR.

Justificacion:

- El apartado de rendimiento debe ser reproducible y auditable.
- Codex no puede validar experimentalmente el consumo real de la VM.
- Las conclusiones de coste bajo/moderado/alto solo se emitiran tras ejecutar los
  tres escenarios y procesar logs reales.

## D034 - Benchmark rendimiento: SkipServiceControl y roles de proceso

Se corrige el benchmark de rendimiento tras error `NamedParameterNotFound` con
`-SkipServiceControl`.

Decision:

- `SkipServiceControl` se implementa como `[bool]` en el `param()` real.
- Se mantiene compatibilidad con invocacion:
  `-SkipServiceControl $true`.
- Si `SkipServiceControl=true`, el script no ejecuta `Stop-Service` ni
  `Start-Service`; solo lee estado del servicio.
- El servicio Velociraptor observado en VM se trata como cliente HIDS:
  `C:\Program Files\Velociraptor\Velociraptor.exe` con `client.config.yaml`
  y `service run`.
- El proceso local server/GUI con `server.config.yaml gui` queda como
  `SERVER_GUI` y no se mezcla con las metricas del cliente.
- `notepad.exe` se excluye aunque su command line contenga rutas con
  `Velociraptor`.
- Las metricas no disponibles quedan como `NA/null`, no como `0`, y no entran en
  medias.
- Los resultados anteriores al fix quedan marcados como parciales/legacy.

Justificacion:

- El objetivo del apartado de rendimiento es medir el coste del agente cliente
  Velociraptor como HIDS, no el coste del server/GUI local de laboratorio.
- El control manual del servicio es mas trazable en la VM y evita errores por
  permisos al parar/arrancar servicios desde el script.

## D035 - Benchmark rendimiento: invalidacion de datos legacy 20260618

Se revisan los logs existentes en `05_LOGS/BENCHMARKS` y se decide no usarlos
para conclusiones definitivas de rendimiento del agente cliente Velociraptor.

Decision:

- Mantener los datos como evidencia de calidad y trazabilidad.
- Consolidarlos en `05_LOGS/BENCHMARKS/validado`.
- Generar Excel con tablas y graficas de validacion, pero marcarlo como
  `NO_VALIDO_PARA_CONCLUSIONES_DEFINITIVAS_DE_RENDIMIENTO`.
- No usar metricas legacy de CPU/RAM como coste del agente HIDS.
- Excluir `notepad.exe` como falso match.
- Separar `velociraptor-v0.75.6-windows-amd64.exe` como server/laboratorio.
- Exigir `CLIENT_SERVICE` observado en `VR_IDLE` y `VR_TEC_RUNNER` para un
  benchmark definitivo.

Justificacion:

- Los tres `summary.json` son `SchemaVersion=1.0`.
- `VR_IDLE` y `VR_TEC_RUNNER` aparecen con `ServiceStatus=Stopped`.
- No se observo `C:\Program Files\Velociraptor\Velociraptor.exe` con
  `client.config.yaml service run`.
- El detector legacy incluyo `notepad.exe` por command line con rutas que
  contienen la palabra `Velociraptor`.
- Algunas metricas no disponibles fueron rellenadas con `0` por el script
  anterior, lo que impediria una interpretacion rigurosa.

## D036 - Carpeta de Exceles definitivos para memoria

Se crea `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS` como paquete unico de
Exceles para entregar en la memoria.

Decision:

- Usar como Excel principal de deteccion/FP/fiabilidad:
  `01_TECNICAS_FP_FIABILIDAD_DEFINITIVO.xlsx`.
- Usar como Excel de benchmark:
  `02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS.xlsx`.
- Mantener como referencias/anexos:
  - `03_ANALISIS_TECNICAS_TFM_V4_REFERENCIA.xlsx`;
  - `04_ANALISIS_ARTIFACTS_PUBLICOS_REFERENCIA.xlsx`.
- No recomendar mas Exceles principales para no fragmentar la evidencia.
- Mantener el benchmark con advertencia explicita:
  no mide consumo definitivo del agente cliente hasta repetirlo con
  `CLIENT_SERVICE` observado.

Justificacion:

- El Excel TEC+FP+fiabilidad ya integra la campana ofensiva validada, FP v1.1,
  calidad de asociacion y conclusiones.
- El Excel benchmark actual documenta correctamente la invalidacion de los logs
  legacy y evita presentar conclusiones no soportadas.
- `Analisis_Tecnicas_TFM_V.4.xlsx` aporta continuidad historica.
- `Analisis_Artifacts_Publicos_Velociraptor_TFM_10_10.xlsx` aporta soporte
  metodologico si la memoria compara artifacts publicos y custom.

## D037 - Dictamen de auditoria de Exceles definitivos

Tras auditoria tecnica de la carpeta
`04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_DEFINITIVOS`, se decide:

- Estado global `NO APTO` hasta resolver los blockers.
- No modificar ni generar copias corregidas automaticas mientras no se elija
  una campaña ofensiva canonica y no exista benchmark cliente valido.
- El libro 02 se conserva solo como control de calidad, no como benchmark
  definitivo.
- El libro 01 debe reconstruir recuentos SOC desde un unico JSONL y separar
  ejecucion, visibilidad, deteccion, alerta RT y salida externa.
- Los KPI FP principales deben excluir las tres filas contaminadas del RunId
  anterior; 15 queda como bruto y 12 como conjunto limpio.
- TEC-009 tiene transferencia real confirmada por receiver log y ZIP; el Excel
  debe enlazar esa evidencia original.
- Los libros 03 y 04 solo pueden mantenerse como referencia/anexo si quedan
  claramente rotulados y no se confunden con resultados finales.

Justificacion:

- El benchmark cliente definitivo falta.
- Hay recuentos no reconciliados y trazabilidad insuficiente en el libro 01.
- Existen rangos incorrectos en un grafico y dos tablas.

## D038 - Regeneración basada en campañas canónicas separadas

Antes de reconstruir los Excel definitivos se decide exigir tres conjuntos de
fuentes independientes:

1. Campaña ofensiva canónica TEC-001..TEC-009.
2. Campaña de falsos positivos con JSONL propio y sin contaminación.
3. Benchmark v1.1 válido de baseline, idle y runner.

No se deben concatenar JSONL de campañas distintas ni inferir una detección a
partir de Discord, persistencia JSONL o visibilidad genérica. TEC-009 requiere
además evidencia real de transferencia y concordancia de hash.

El autor realizará la ejecución experimental. Codex preparará nuevas versiones
solo en `candidate`, validará las salidas y regenerará los derivados en
`REVIEWED`, manteniendo intactos originales y artifacts validados.

Justificación:

- El benchmark legacy no observa correctamente `CLIENT_SERVICE`.
- La campaña FP existente contiene tres filas de otro RunId.
- Los recuentos SOC del libro 01 mezclan fuentes y no tienen trazabilidad
  inequívoca a una única campaña.

## D039 - Paquete de repetición controlada

Se crea `06_CONTROLLED_RERUN` como única vía operativa para obtener las nuevas
fuentes antes de reconstruir Excel.

Decisiones:

- JSONL exclusivo para la campaña TEC.
- JSONL independiente para cada una de tres repeticiones FP.
- Listener benigno local controlado para FP-009; `SKIPPED` no es aceptable.
- Benchmark con tres repeticiones de cada escenario y salida separada.
- La carga TEC del benchmark usa `EnableExfilUpload=false`; no se mezcla con la
  transferencia TEC-009 canónica.
- TEC-009 requiere HTTP 2xx, receiver POST, ZIP recibido y SHA-256 coincidente.
- El validador exige exports originales de `CLIENT_EVENT`; JSONL solo acredita
  persistencia/enrutamiento.
- No se reconstruyen CSV/Excel hasta `VALIDACION PASS`.
- No se sobrescriben artifacts `validated`.

Justificación:

- Evitar repetir contaminación temporal y mezcla de campañas.
- Separar evidencia de detección, salida SOC y evidencia externa.
- Obtener benchmark cliente SchemaVersion 1.1 reproducible y auditable.

## D040 - Patch Windows PowerShell 5.1 para TEC-009 4104

`TEC_20260619_FINAL01` se clasifica `NO APTA`: la transferencia existió, pero
no se generó el summary global.

Se decide:

- conservar v5 y v6 intactos;
- crear `TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1`;
- sustituir conversiones `@($genericList)` por `.ToArray()`;
- mantener un error de consulta 4104 como `WARNING`, nunca transformarlo en OK;
- impedir que ese error suprima el summary global;
- escribir logging completo de bloque, técnica, conteos, posición, stack y
  excepción;
- exigir selftest PASS antes de repetir `TEC_20260619_FINAL02`.

Justificación:

- `PSToObjectArrayBinder` de Windows PowerShell 5.1 puede fallar al convertir
  `System.Collections.Generic.List[object]` mediante `@(...)`.
- La evidencia parcial de FINAL01 no compensa la falta de trazabilidad del
  summary global.

## D041 - Cronología original y tratamiento P2 en FP

Se decide:

- usar `DetectionTime` del CLIENT_EVENT como tiempo de la detección;
- tratar `Timestamp` del router como tiempo posterior de persistencia;
- conservar los JSONL raw aunque contengan eventos tardíos de otra repetición;
- no usar sus recuentos brutos para Excel hasta reconciliarlos con CLIENT_EVENT;
- aceptar P2 en FP como WARN/adjudicación, no como fallo de ejecución por sí
  solo, siempre que P1 sea cero y exista evidencia original trazable;
- no modificar artifacts para ocultar P2 antes de cerrar la medición;
- filtrar mecánicamente exports amplios por las ventanas UTC exactas, sin
  editar filas ni timestamps.

Justificación:

- El router es asíncrono y puede persistir una fila después de limpiar el
  JSONL; eso no cambia cuándo ocurrió el CLIENT_EVENT.
- Los P2 observados corresponden a acciones dual-use legítimas: módulos
  Defender 4104, `cmd.exe`, `schtasks.exe` y discovery administrativo.
- P2 aporta visibilidad/forense de baja especificidad; no equivale a P1 ni a
  detección maliciosa fuerte.

## D042 - Criterio benchmark RunnerStillRunningAtEnd

`RunnerStillRunningAtEnd=True` se registra como observación de limpieza y no
como invalidez del benchmark cuando se cumplen simultáneamente:

- `SchemaVersion=1.1`;
- `ScenarioValidity=VALID`;
- cliente Velociraptor observado;
- `RunnerObservedSamples > 0`;
- SERVER_GUI excluido;
- notepad.exe no clasificado como runner.

El objetivo del escenario es medir carga mientras el runner está activo. Exigir
que termine antes del fin de la ventana contradice ese objetivo y no forma
parte del contrato definido en TEST_MATRIX.

## D043 - Fuente única del Excel de benchmark regenerado

El Excel de benchmark para entrega se reconstruye exclusivamente desde los
nueve runs de `BENCH_20260619_FINAL01`. El libro antiguo
`VR_RESOURCE_BENCHMARK.xlsx` queda excluido como fuente porque contiene
`ClientRows=0`, `ValidForClientBenchmark=0` y gráficos vacíos.

La validez derivada se calcula, no se fuerza: exige 9/9 con
`SchemaVersion=1.1` y `ScenarioValidity=VALID`, baseline sin cliente,
idle/runner con cliente, runner observado en los tres runs de carga,
`SERVER_GUI` excluido y cero clasificaciones de `notepad.exe` como runner.
# Decision 20260626_1143 - Excel V.5 visibilidad/FP/Wazuh

- Para el Excel V.5/final, la campana REAL se consolida solo con filas
  `CLIENT_EVENT` de `05_LOGS/Validas/Prueba_24_06_2026` dentro de la ventana
  `2026-06-24T12:50:22.335Z..2026-06-24T12:55:53.996Z`.
- La campana FP se consolida por `summary.json` y `vr_hits.csv` del RunId
  `TFM_FP_20260626_101948`; `TotalHits=0` se conserva sin inventar filas.
- Las filas `CLIENT_EVENT` observadas en ventana FP pero no reconciliadas con
  `vr_hits` se documentan en `DISCREPANCIAS` y no se usan como FP definitivos.
- Los CSV con fechas mezcladas se auditan y se filtran por ventana; no se
  agregan filas fuera de ventana a resultados definitivos.
- TEC-009 no se presenta como exfiltracion contextual completa sin matiz,
  porque no se localizo `receiver_log.jsonl` ni ZIP recibido en las rutas
  revisadas. Se conserva como upload HTTP controlado documentado por summary
  con HTTP 200, `UploadSucceeded=True`, SHA-256 y bytes en `ReceiverResponse`.
- Wazuh `110201=0` se conserva como gap TEC-009.
- Se uso Excel COM local para construir el XLSX porque `@oai/artifact-tool` no
  estaba disponible en el entorno de ejecucion.

## D044 - Benchmark final 26062026 para memoria

Se decide generar y aceptar el Excel final de benchmark porque el precheck es
APTO:

- 9/9 runs validos.
- 3 escenarios: `BASELINE_NO_VR`, `VR_IDLE`, `VR_TEC_RUNNER`.
- 3 repeticiones por escenario.
- `SchemaVersion=1.1` en todos los runs.
- `SERVER_GUI` excluido del calculo principal.
- `notepad.exe` runner igual a 0.
- graficas pobladas con datos reales.
- fuentes con SHA-256.

La fuente canonica es:

- `04_EVIDENCE/ENTREGA_MEMORIA_EXCELES_REGENERADOS/02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx`.

`Analisis_Benchmark.xlsx` queda limitado a referencia historica/formato.

`RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER` se conserva como WARN no
bloqueante, segun D042. No invalida el benchmark porque se cumplio la ventana
de muestreo y el objetivo era medir carga mientras el runner estaba activo.

El Excel benchmark no mide calidad de deteccion, alertas, evidencia ni falsos
positivos. Solo mide coste operativo observado en laboratorio.
