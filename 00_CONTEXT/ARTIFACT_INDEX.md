# ARTIFACT_INDEX

| Artifact | Version | Estado | Nota |
|---|---:|---|---|
| P4 Low Basic | v2 | Funcional | Mantener activo |
| P3 Medium Behavioral | v5 | Parcial | CU-009 no detecta correctamente |
| P3 Medium Behavioral | v6 | Revisado / superado por v7 | Mezclaba deteccion fuerte CU-009 con contexto opcional |
| P3 Medium Behavioral | v7 | Creado / superado por v8 | Separa CU-009 fuerte de staging/contexto opcional, pero mezcla SHA256/Get-FileHash con crypto fuerte |
| P3 Medium Behavioral | v8 | Creado / pendiente de validacion VM | Separa cifrado real de hashing/evidencia; CU-009 permite SHA256/Get-FileHash y excluye cifrado real |
| Debug PowerShell 4104 Raw | v1 | Creado / superado por v2 | Mezclaba canal logico y ruta fisica en pruebas de EVTX |
| Debug PowerShell 4104 Raw | v2 | Creado / superado para Hunt por artifacts separados | Es CLIENT_EVENT; no aparece como collection manual en Hunt Manager |
| Debug PowerShell 4104 ParseCU009 | v1 | Creado / pendiente de ejecucion VM | CLIENT para Hunt/Collection manual con parse_evtx historico |
| Debug PowerShell 4104 WatchCU009 | v1 | Creado / pendiente de ejecucion VM | CLIENT_EVENT para Client Monitoring live con watch_evtx |
| Debug PowerShell 4104 CU009Shape | v3 | Ejecutado / superado por v4 | CLIENT para backtest; lee 4104 real, pero la forma real de comando queda demasiado estricta y no marca el upload multilinea |
| Debug PowerShell 4104 CU009Shape | v4 | Creado / pendiente de ejecucion VM | CLIENT para backtest; usa coexistencia estructural HTTP tool + URI + POST + InFile + application/zip y no bloquea por limpieza local |
| P2 High Forensic | - | No activo | No tocar todavia |
| P1 Critical Correlation | - | No activo | No tocar todavia |
| Router JSONL Discord | v3_fix3 | Funcional | Discord off |
