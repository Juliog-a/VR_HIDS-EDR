# PLAN_ARTIFACTS

Plan preliminar de artifacts Velociraptor para el TFM. Este documento no define artifacts YAML; solo establece la trazabilidad entre tecnica MITRE ATT&CK, script de prueba, comportamiento detectable y artifact necesario.

## Criterios de diseno

- Mantener separadas visibilidad, deteccion, alerta en tiempo real, evidencia forense y salida externa.
- Evitar detecciones basadas en rutas de laboratorio, nombres de scripts, comentarios o etiquetas TEC.
- Usar los Casos de Uso como unidad metodologica y los artifacts como unidad tecnica de despliegue.
- Usar Sysmon ID 26 como evidencia forense de borrado, no como alerta individual.
- Tratar TEC-009 como staging, archivado y red contextual; no como exfiltracion real si no hay transferencia demostrada de datos.

## Matriz de artifacts

| CU / TEC | Tecnica MITRE ATT&CK | Script de prueba asociado | Comportamiento detectable | Artifact Velociraptor necesario | Tipo de artifact | Prioridad | Complejidad | Justificacion tecnica breve |
|---|---|---|---|---|---|---|---|---|
| CU-001 / TEC-001 | T1059.001 - Command and Scripting Interpreter: PowerShell | `TFM_Run_All_TEC_Tests.ps1` | Ejecucion de `powershell.exe` con `-NoProfile`, `ExecutionPolicy Bypass`, posible `EncodedCommand`, `IEX` o `DownloadString`. | `Custom.TFM.HIDS.P3.Medium.Behavioral` | `CLIENT_EVENT` | P3 | Media | PowerShell aislado genera mucho ruido; requiere condiciones sobre command-line y/o 4104 para diferenciar visibilidad de deteccion. |
| CU-002 / TEC-002 | T1059.003 - Command and Scripting Interpreter: Windows Command Shell | `TFM_Run_All_TEC_Tests.ps1` | Uso de `cmd.exe /c`, encadenamiento de comandos y llamada indirecta a PowerShell u otras herramientas administrativas. | `Custom.TFM.HIDS.P4.Low.Basic` y `Custom.TFM.HIDS.P3.Medium.Behavioral` | `CLIENT_EVENT` | P3 | Media | `cmd.exe` es comun en administracion; la deteccion debe centrarse en chaining, ejecucion indirecta y patrones de abuso. |
| CU-003 / TEC-003 | T1053.005 - Scheduled Task/Job: Scheduled Task | `TFM_Run_All_TEC_Tests.ps1`, `tec003_task.cmd` | Creacion, ejecucion o borrado de tarea con `schtasks.exe`; persistencia o ejecucion programada. | `Custom.TFM.HIDS.P4.Low.Basic` y `Custom.TFM.HIDS.P2.High.Forensic` | `CLIENT_EVENT` | P2 | Media | La tarea programada deja senales claras en proceso y eventos; conviene correlacionar accion, usuario, comando y contexto para reducir falsos positivos. |
| CU-004 / TEC-004 | T1547.001 - Boot or Logon Autostart Execution: Registry Run Keys / Startup Folder | `TFM_T1547_001.bat`, `TFM_Run_All_TEC_Tests.ps1` | Escritura o borrado en `Run`/`RunOnce`, uso de `reg.exe`, valores apuntando a binarios o interpretes. | `Custom.TFM.HIDS.P4.Low.Basic` y `Custom.TFM.HIDS.P2.High.Forensic` | `CLIENT_EVENT` | P2 | Media | Las Run Keys son persistencia frecuente; debe priorizarse el contenido del valor, proceso origen y rutas sospechosas frente a cambios legitimos de instaladores. |
| CU-005 / TEC-005 | T1569.002 - System Services: Service Execution | `scTECN005.ps1`, `sc.txt`, `TFM_Run_All_TEC_Tests.ps1` | Creacion de servicio Windows, Event ID 7045, `sc.exe create/start/delete` o `New-Service`. | `Custom.TFM.HIDS.P4.Low.Basic` | `CLIENT_EVENT` | P2 | Baja | Event ID 7045 proporciona evidencia precisa y es el mejor caso para comparar con `Windows.Events.ServiceCreation`. |
| CU-006 / TEC-006 | T1518.001 - Software Discovery: Security Software Discovery | `TEC006.ps1`, `TFM_Run_All_TEC_Tests.ps1` | Consulta de `root\\SecurityCenter2`, `AntivirusProduct`, `Get-CimInstance`, `wmic`, servicios/procesos AV/EDR. | `Custom.TFM.HIDS.P3.Medium.Behavioral` | `CLIENT_EVENT` | P3/P4 | Media | El discovery de seguridad puede ser administracion legitima; debe tratarse como senal contextual y elevarse si aparece con ejecucion sospechosa. |
| CU-007 / TEC-007 | T1486 - Data Encrypted for Impact | `Scriptransom.ps1`, `encrypt.ps1`, `enumeracion.ps1`, `descifrar.ps1` | Uso de AES, `CryptoStream`, `System.Security.Cryptography`, creacion masiva de `.aes`, borrado/movimiento de originales. | `Custom.TFM.HIDS.P1.Critical.Correlation` y `Custom.TFM.HIDS.P2.High.Forensic` | `CLIENT_EVENT` | P1 | Alta | Requiere combinar 4104, creacion de ficheros y umbrales; si hay senales crypto y extension cifrada debe priorizarse sobre borrado generico. |
| CU-008 / TEC-008 | T1485 - Data Destruction | `sabotaje.ps1`, `TEC-008_sabotaje_controlado.ps1`, `TFM_Run_All_TEC_Tests.ps1` | Enumeracion seguida de borrado masivo con `Remove-Item`, `-Force`, `-Recurse`, `ForEach-Object`; multiples Sysmon ID 26 como evidencia. | `Custom.TFM.HIDS.P1.Critical.Correlation` y `Custom.TFM.HIDS.P2.High.Forensic` | `CLIENT_EVENT` | P1 | Alta | La alerta operativa debe consolidarse por comportamiento destructivo; los eventos ID 26 individuales son evidencia forense, no notificacion directa. |
| CU-009 / TEC-009 | T1074.001 - Local Data Staging; T1560.001 - Archive Collected Data; T1020/T1041 solo contextual | `exfiltracion.ps1`, `TEC-009_exfiltracion_simulada.ps1`, `TFM_Run_All_TEC_Tests.ps1` | `Compress-Archive`, `Copy-Item`, generacion de ZIP, conexion TCP externa controlada con `TcpClient` hacia puerto 80. | `Custom.TFM.HIDS.P3.Medium.Behavioral`, `Custom.TFM.HIDS.P2.High.Forensic` y `Custom.TFM.HIDS.P1.Critical.Correlation` | `CLIENT_EVENT` | P2 | Alta | La deteccion debe correlacionar staging, archivado y red externa. La red solo permite cobertura contextual de exfiltracion si se documenta telemetria, sin afirmar transferencia real. |

## Artifacts auxiliares

| Artifact Velociraptor | Tipo de artifact | Uso previsto | Prioridad | Complejidad | Justificacion tecnica breve |
|---|---|---|---|---|---|
| `Custom.TFM.HIDS.P0.Debug.Raw` | `CLIENT_EVENT` | Validar campos reales de 4104, Sysmon y 7045 antes de activar reglas. | Laboratorio | Baja | Reduce riesgo de errores de campo en VQL y permite ajustar los artifacts finales al entorno real. |
| `Custom.TFM.HIDS.Discord.Router.SERVER_EVENT` | `SERVER_EVENT` | Enrutar alertas ya clasificadas hacia Discord. | P2 | Media | Discord es salida externa, no detector; debe consumir eventos normalizados y evitar spam. |
| `Custom.TFM.HIDS.Benchmark.ProcessSampler` | `CLIENT_EVENT` | Medir consumo por proceso durante B1-B8. | P2 | Media | Permite justificar coste operativo por perfil de artifact y comparar publicos frente a custom. |

## Orden de implementacion recomendado

1. Validar telemetria real con `P0 Debug Raw`.
2. Implementar `P4 Low Basic` para CU-001 a CU-005 con bajo coste.
3. Implementar `P3 Medium Behavioral` para PowerShell 4104 y CU-006 a CU-009.
4. Implementar `P2 High Forensic` para Sysmon ID 3, 11, 12/13/14 y 26.
5. Implementar `P1 Critical Correlation` para ransomware, destruccion y staging con correlacion.
6. Activar `Discord Router` solo cuando los `CLIENT_EVENT` ya emitan alertas normalizadas.
7. Ejecutar benchmark por perfil y documentar coste, volumen de eventos y ratio eventos/alertas.
