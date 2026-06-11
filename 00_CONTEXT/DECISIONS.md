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
