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
