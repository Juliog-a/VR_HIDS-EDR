# ARTIFACT_INDEX

| Artifact | Version | Estado | Nota |
|---|---:|---|---|
| P4 Low Basic | v2 | Funcional | Mantener activo |
| P3 Medium Behavioral | v5 | Parcial | CU-009 no detecta correctamente |
| P3 Medium Behavioral | v6 | Revisado / superado por v7 | Mezclaba deteccion fuerte CU-009 con contexto opcional |
| P3 Medium Behavioral | v7 | Creado / superado por v8 | Separa CU-009 fuerte de staging/contexto opcional, pero mezcla SHA256/Get-FileHash con crypto fuerte |
| P3 Medium Behavioral | v8 | Creado / superado por v9 | Separa cifrado real de hashing/evidencia; CU-009 permite SHA256/Get-FileHash y excluye cifrado real, pero no integra shape v4 validado |
| P3 Medium Behavioral | v9 | Candidate / no validado | En live emite staging local, pero no emite CU-009 fuerte 4104; probable problema por `foreach` sobre `watch_evtx` |
| P3 Medium Behavioral | v10 | Candidate / superado arquitectonicamente | Reintegra CU-009 fuerte 4104 con `watch_evtx` directo sobre `EventData.ScriptBlockText`; conserva shape v4 y columnas auditables |
| P3 Medium Behavioral | v11 | Candidate / pendiente validacion VM | Reorganiza P3 como capa media/contextual; no valida CU-007/CU-008/CU-009 fuertes |
| P2 High Forensic | v1 | Candidate / pendiente validacion VM | Artifact especializado para CU-007, CU-008 y CU-009; integra CU009Shape_v4 en P2 |
| P1 Low Basic | v3 | Candidate / pendiente validacion Velociraptor | Artifact historico con parse_evtx; cobertura TEC-001 a TEC-009 |
| P2 High Forensic | v2 | Candidate / pendiente validacion Velociraptor | Artifact historico con parse_evtx; detalle forense y contexto 9/9 |
| P3 Medium Behavioral | v12 | Candidate / pendiente validacion Velociraptor | Artifact historico con parse_evtx; salida normalizada alertable 9/9 |
| Debug Raw PowerShell 4104 LastHours | v1 | Candidate / pendiente ejecucion VM | CLIENT parse_evtx RAW para confirmar lectura 4104 |
| Debug Raw Sysmon ID1_3_11 LastHours | v1 | Candidate / pendiente ejecucion VM | CLIENT parse_evtx RAW para confirmar lectura Sysmon ID 1/3/11 |
| Debug PowerShell 4104 Raw | v1 | Creado / superado por v2 | Mezclaba canal logico y ruta fisica en pruebas de EVTX |
| Debug PowerShell 4104 Raw | v2 | Creado / superado para Hunt por artifacts separados | Es CLIENT_EVENT; no aparece como collection manual en Hunt Manager |
| Debug PowerShell 4104 ParseCU009 | v1 | Creado / pendiente de ejecucion VM | CLIENT para Hunt/Collection manual con parse_evtx historico |
| Debug PowerShell 4104 WatchCU009 | v1 | Creado / pendiente de ejecucion VM | CLIENT_EVENT para Client Monitoring live con watch_evtx |
| Debug PowerShell 4104 CU009Shape | v3 | Ejecutado / superado por v4 | CLIENT para backtest; lee 4104 real, pero la forma real de comando queda demasiado estricta y no marca el upload multilinea |
| Debug PowerShell 4104 CU009Shape | v4 | Validado experimentalmente por investigador | CLIENT para backtest; usa coexistencia estructural HTTP tool + URI + POST + InFile + application/zip y no bloquea por limpieza local |
| P1 Critical Correlation | - | No activo | Reservado para correlacion critica final; no implementar todavia |
| Router JSONL Discord | v3_fix3 | Funcional | Discord off |
| Router JSONL Discord SOC | v1 | Candidate / pendiente validacion Velociraptor | SERVER_EVENT para enrutar P1/P2/P3 EVENT a JSONL y Discord opcional con formato SOC |
| Router JSONL Discord SOC | v2 | Candidate / pendiente validacion Velociraptor | SERVER_EVENT para enrutar P1/P2/P3/P4 a JSONL y Discord opcional; corrige nombres P2/P4 y normalizacion SOC |
| Router JSONL Discord SOC | v3 | Candidate / pendiente validacion Velociraptor | SERVER_EVENT con una source independiente por source monitorizada; apunta a P2 Event v1, P1 Critical, P3 Event y P4 v2 |
| P1 Critical Priority Event | v1 | Recalibrado / pendiente validacion Velociraptor | P1 critico estricto; conserva artifact/source names para SOC_v3; emite solo TEC-004/005/007/008/009 fuertes |
