# PROJECT_STATE

## Estado global actual

- Proyecto: TFM Velociraptor como HIDS/DFIR.

- Arquitectura custom: P4 / P3 / P2 / P1.

- P4_v2: funcional.

- P3_v5: genera mas evidencia, pero CU-009 sigue pendiente.

- P3_v6: revisado; mezclaba deteccion fuerte CU-009 con contexto opcional.

- P3_v7: creado en candidate; intacto; superado por P3_v8 para corregir separacion crypto/hash en CU-009.

- P3_v8: creado en candidate; importado segun validacion del investigador, pero CU-009 4104 no esta validado; en diagnostico con Debug PowerShell 4104 CU009Shape v4 por fallo de forma de comando en v3.

- Router JSONL: funcional.

- Discord: desactivado.

- P2/P1: desactivados.

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
