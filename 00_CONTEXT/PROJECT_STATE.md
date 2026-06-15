# PROJECT_STATE

## Estado global actual

- Proyecto: TFM Velociraptor como HIDS/DFIR.

- Arquitectura custom: P4 / P3 / P2 / P1.

- P4_v2: funcional.

- P3_v5: genera mas evidencia, pero CU-009 sigue pendiente.

- P3_v6: revisado; mezclaba deteccion fuerte CU-009 con contexto opcional.

- P3_v7: creado en candidate; intacto; superado por P3_v8 para corregir separacion crypto/hash en CU-009.

- P3_v8: creado en candidate; superado por P3_v9 para CU-009 4104.

- Debug PowerShell 4104 CU009Shape v4: validado experimentalmente por el investigador; confirma forma real de upload HTTP ZIP en 4104.

- P3_v9: creado en candidate; no validado experimentalmente. En prueba live, staging local emitio fila, pero CU-009 fuerte 4104 no emitio.

- P3_v10: creado en candidate; revisa solo integracion CU-009 fuerte 4104 usando `watch_evtx` directo sobre `EventData.ScriptBlockText`, sin wrapper `foreach` live. No validado; superado arquitectonicamente por P3_v11/P2_v1.

- P3_v11: creado en candidate. Reorganiza P3 como capa media de comportamiento general y contexto. Degrada CU-007/CU-008/CU-009 fuertes a contexto no responsable de validacion critica.

- P2_v1: creado en candidate. Artifact especializado, pequeno y auditable para CU-007, CU-008 y CU-009. Integra CU009Shape_v4 dentro de P2.

- Runner v4: creado en `03_RUNNERS` y copiado a `Carpeta_Compartida_TFM/runner_candidate`. Actualiza la campana completa a la ruta activa `01_ACTIVE_TESTS\Pruebas` y ejecuta TEC-009 mediante `exfiltracion_v3.ps1` con upload HTTP controlado.

- Runner v4 revision 20260612_1215: corregido TEC-009 para invocar `exfiltracion_v3.ps1` por ruta absoluta. Anade resumen final TXT/JSON/CSV por tecnica.

- Visibility checker: creado `TFM_Check_Sysmon_Visibility.ps1` como script de solo consulta para evaluar disponibilidad de Sysmon, PowerShell, Security, System y Application.

- Ultima iteracion 20260612_1258: creado paquete `ultima_iteracion` con runner v4 corregido, checker, scripts corregidos TEC-007/TEC-008/TEC-009, logs de referencia, hashes y README. No se ejecuto campana completa desde Codex.

- Router JSONL: funcional.

- Discord: desactivado.

- P2/P1: no validados experimentalmente. P2_v1 queda candidate pendiente de importacion/validacion; P1 reservado sin implementar.

- Iteracion artifacts historicos 20260612_1418:
  - P1_v3 creado en candidate como Low Basic historico con `parse_evtx()`.
  - P2_v2 creado en candidate como High Forensic historico con cobertura 9/9.
  - P3_v12 creado en candidate como Medium Behavioral historico con salida normalizada.
  - Debug RAW 4104 y Sysmon ID 1/3/11 creados en debug.
  - Paquete `ultima_iteracion_artifacts` creado con backups, analisis, README y hashes.
  - ZIP final `ultima_iteracion_artifacts.zip` creado.
  - No se modifica runner, scripts de ataque, receiver, router ni Discord.

- Validacion experimental: manual por el autor.

## Flujo metodologico

Sysmon / PowerShell 4104

-> CLIENT_EVENT custom

-> SERVER_EVENT router

-> JSONL

-> analisis posterior

## Criterio critico

Telemetria no equivale a deteccion.

JSONL no equivale a deteccion.

Discord no equivale a deteccion.

Codex no valida resultados.
