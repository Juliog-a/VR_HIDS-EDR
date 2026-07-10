# TEST_MATRIX

## Propósito

Matriz canónica de repetición controlada para la evaluación de Velociraptor
como HIDS/DFIR. Esta matriz sustituye la ausencia previa del fichero, pero no
declara validación experimental: los resultados deben proceder de ejecuciones
realizadas por el autor.

## Separación obligatoria de niveles

| Nivel | Qué demuestra | Qué no demuestra |
|---|---|---|
| Ejecución del runner | El escenario se ejecutó y produjo su evidencia local esperada | Que Velociraptor lo detectó |
| `CLIENT_EVENT` | El artifact de host emitió una detección o señal | Que el router persistió o notificó |
| `SERVER_EVENT` SOC_v3 | Enrutamiento y normalización | Detección original |
| JSONL | Persistencia externa de la salida SOC | Detección por sí sola |
| Receiver TEC-009 | Transferencia HTTP real de ZIP dummy | Detección HIDS por sí sola |
| Discord | Notificación | Detección |

## Configuración live controlada

| Perfil | Artifact | Tipo | Cobertura esperada |
|---|---|---|---|
| P1 crítico | `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` recalibrado | `CLIENT_EVENT` | TEC-004, TEC-005, TEC-007, TEC-008 y TEC-009 fuertes |
| P2 forense | `Custom.TFM.HIDS.P2.High.Forensic.Event_v1` | `CLIENT_EVENT` | TEC-001 a TEC-009; señal fuerte y contexto forense |
| P3 comportamiento | `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1` | `CLIENT_EVENT` | TEC-001 a TEC-009; comportamiento y alertabilidad |
| P4 básico | `Custom.TFM.HIDS.P4.Low.Basic_v2` | `CLIENT_EVENT` | TEC-001 a TEC-005; visibilidad básica |
| Router | `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3` | `SERVER_EVENT` | Enrutamiento P1/P2/P3/P4; no detecta |

No se deben activar simultáneamente P1 legacy y P1 crítico. Para la repetición:
`EnableJSONL=true`, `EnableDiscord=false`, SOC_v1/v2 desactivados.

## Matriz de técnicas TEC

| TEC | Técnica / MITRE ATT&CK | Escenario ejecutado | Criterio local del runner | `CLIENT_EVENT` fuerte/principal esperado | Perfiles esperados | Evidencia adicional |
|---|---|---|---|---|---|---|
| TEC-001 | PowerShell / T1059.001 | PowerShell con Bypass/NoProfile y bloque 4104 controlado | Comandos terminan y existen eventos 4104/proceso asociados | P3 `TEC001_PowerShell_4104`; P2 `Forensic_EVENT_PowerShell4104_Strong_By_Technique` | P2, P3, P4 | Sysmon ID 1 es apoyo |
| TEC-002 | Windows Command Shell / T1059.003 | `cmd.exe` con encadenamiento e hijo PowerShell | Exit code correcto y proceso observado | P3 `TEC002_CMD_Sysmon`; P2 `Forensic_EVENT_Sysmon_ID1_Process_By_Technique` | P2, P3, P4 | Sysmon ID 1 |
| TEC-003 | Scheduled Task / T1053.005 | Crear, consultar, ejecutar y borrar tarea temporal | Operaciones controladas completadas; sin persistencia residual | P3 `TEC003_ScheduledTask_Sysmon`; P2 proceso por técnica | P2, P3, P4 | Security 4698/4699 opcional, no bloqueante |
| TEC-004 | Registry Run Keys / T1547.001 | Añadir, consultar y borrar valor Run temporal | Clave creada, consultada y retirada | P1 `P1_EVENT_Sysmon_ID1_Process_Basic`; P3 `TEC004_RunKey_Sysmon`; P2 proceso/contexto registry | P1, P2, P3, P4 | Sysmon 12/13/14 como contexto |
| TEC-005 | Service Execution / T1569.002 | Crear, consultar, intentar iniciar y borrar servicio temporal | Servicio creado y eliminado; System 7045 si está disponible | P1 `P1_EVENT_System7045_ServiceCreation`; P3 `TEC005_Service_System7045`; P2 `Forensic_EVENT_System_7045_ServiceCreation` | P1, P2, P3, P4 | Sysmon ID 1 `sc.exe` |
| TEC-006 | Security Software Discovery / T1518.001 | Consultas benignas de SecurityCenter2 y servicios de seguridad | Comandos completados y 4104/proceso asociado | P3 `TEC006_SecuritySoftwareDiscovery_4104`; P2 PowerShell fuerte | P2, P3 | Sysmon ID 1 como fallback/contexto |
| TEC-007 | Data Encrypted for Impact / T1486 | Cifrado AES de datos dummy controlados | Se crea al menos un `.aes`; restore posterior | P1/P2 PowerShell 4104 fuerte; P3 `TEC007_DataEncryptedForImpact_4104` | P1, P2, P3 | Sysmon ID 11 `.aes` es apoyo, no único criterio |
| TEC-008 | Data Destruction / T1485 | Borrado controlado solo dentro del dataset dummy | `FilesBefore > 0` y borrado controlado completado | P1/P2 PowerShell 4104 fuerte; P3 `TEC008_DataDestruction_4104` | P1, P2, P3 | Sysmon 23/26 solo evidencia forense; ID 26 no alerta individual |
| TEC-009 | Staging T1074.001, Archive T1560.001 y transferencia controlada T1048.003 | Crear ZIP dummy y enviarlo por POST al receiver del investigador | ZIP generado; SHA-256 local; `ReceiverReachable=True`; HTTP 2xx; `UploadSucceeded=True`; bytes transferidos en `ReceiverResponse` | P1/P2 4104 fuerte; P3 `TEC009_HTTP_ZIP_Upload_4104` | P1, P2, P3 | Receiver POST, `application/zip`, ZIP recibido y SHA-256 origen=destino si se conservan; su ausencia en el paquete revisado es limitación documental secundaria, no invalidación automática |

### Criterio global TEC

- Runner: nueve técnicas, `9 OK`, `0 WARN`, `0 FAIL`.
- JSONL exclusivo: cero errores de parseo y cobertura TEC-001..TEC-009.
- Cada fila SOC debe enlazar con Artifact, Source, perfil, timestamp y fila
  original `CLIENT_EVENT`.
- Los perfiles indican profundidad/prioridad, no una obligación de que todos
  emitan en cada técnica.
- Una señal genérica P4/P3 no se presentará como detección fuerte P1/P2.
- TEC-009 se formula como staging, archivado ZIP y exfiltración HTTP controlada en laboratorio cuando el summary documenta receptor alcanzable, HTTP 2xx, `UploadSucceeded=True`, SHA-256 local y bytes transferidos. Si no se conserva `receiver_log.jsonl` o el ZIP recibido en el paquete revisado, se documenta como limitación secundaria sobre conservación del receptor final.

## Matriz de falsos positivos

| FP | Actividad benigna | Resultado de ejecución esperado |
|---|---|---|
| FP-001 | PowerShell administrativo legítimo | OK |
| FP-002 | CMD legítimo | OK |
| FP-003 | Scheduled Task benigno temporal | OK y limpieza posterior |
| FP-004 | Registro benigno no persistente | OK |
| FP-005 | Consulta de servicios | OK |
| FP-006 | Discovery legítimo de seguridad | OK |
| FP-007 | ZIP benigno local | OK |
| FP-008 | Borrado benigno limitado al workspace | OK |
| FP-009 | GET HTTP a listener benigno local `127.0.0.1:80` | OK, HTTP 200; nunca `SKIPPED` |
| FP-010 | Creación, copia y renombrado normal de ficheros | OK |

Cada una de las tres repeticiones debe tener JSONL propio y ventana temporal
propia. Criterios: diez pruebas, cero `SKIPPED`, cero `FAIL`, cero P1 crítico.
Un P2 requiere adjudicación manual antes de aceptar la campaña. P3/P4 se
registran como falsos positivos potenciales o visibilidad, sin inflar criticidad.

## Matriz de benchmark

| Escenario | Estado servicio | Duración | Repeticiones | Criterio de validez |
|---|---|---:|---:|---|
| BASELINE_NO_VR | Stopped | 180 s | 3 | SchemaVersion 1.1; `VALID`; `VR_Client_ProcessCountMax=0` |
| VR_IDLE | Running | 180 s + 30 s warmup | 3 | SchemaVersion 1.1; `VALID`; `CLIENT_SERVICE > 0` |
| VR_TEC_RUNNER | Running | 300 s | 3 | SchemaVersion 1.1; `VALID`; `CLIENT_SERVICE > 0`; `RunnerObservedSamples > 0` |

Para los nueve runs:

- `SERVER_GUI` puede observarse, pero queda excluido de métricas de cliente.
- `IncludeServerGuiInTotal=false`.
- `notepad.exe` no puede clasificarse como runner ni proceso Velociraptor.
- Deben existir `samples.csv`, `process_samples.csv`, `summary.json`,
  `summary.txt`, `process_list_start.txt`, `process_list_end.txt` y
  `environment.txt`.
- Métricas no disponibles: `NA/null`, nunca cero sintético.

## Resultado de campaña

Solo se podrán reconstruir los CSV y Excel definitivos cuando
`04_Validate_Controlled_Rerun.ps1` produzca `OverallStatus=PASS`. Un `PASS` del
validador confirma consistencia de las evidencias entregadas, no sustituye la
declaración del autor sobre la ejecución experimental.
