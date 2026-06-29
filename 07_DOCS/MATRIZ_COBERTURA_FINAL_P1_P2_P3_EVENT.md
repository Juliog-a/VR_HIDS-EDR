# MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT

Fuente base: `04_EVIDENCE/Analisis_Tecnicas_TFM_Velociraptor_DEF.xlsx` y matriz historica en `ultima_iteracion_artifacts/analisis/MATRIZ_COBERTURA_9_TECNICAS.md`.

Decision TEC-009: el Excel define staging/archive host-based. La campana final incorpora HTTP ZIP upload controlado; por eso se mantiene `T1074.001 / T1560.001` y se anade `T1048.003` como contexto operativo de exfiltracion controlada.

| TEC | Nombre | MITRE_ID | Tactica | P1_EVENT | P2_EVENT | P3_EVENT | Fuente principal | Fuente secundaria | Confidence esperada | Severity | Observaciones |
|---|---|---|---|---|---|---|---|---|---|---|---|
| TEC-001 | PowerShell | T1059.001 | Execution | `P1_EVENT_PowerShell4104_Basic` | `Forensic_EVENT_PowerShell4104_Strong_By_Technique` | `TEC001_PowerShell_4104` | PowerShell 4104 | Sysmon ID 1 si aparece en P2/P1 process | HIGH | MEDIUM | ExecutionPolicy Bypass, FromBase64String, IEX o Invoke-Expression. |
| TEC-002 | Windows Command Shell | T1059.003 | Execution | `P1_EVENT_Sysmon_ID1_Process_Basic` | `Forensic_EVENT_Sysmon_ID1_Process_By_Technique` | `TEC002_CMD_Sysmon` | Sysmon ID 1 cmd.exe | PowerShell 4104 si hay hijo PowerShell | MEDIUM | MEDIUM | Excluye el ruido observado de Adobe/Chrome native messaging en P3. |
| TEC-003 | Scheduled Task | T1053.005 | Persistence | `P1_EVENT_Sysmon_ID1_Process_Basic` | `Forensic_EVENT_Sysmon_ID1_Process_By_Technique` + `Forensic_EVENT_Security_ScheduledTask_Optional` | `TEC003_ScheduledTask_Sysmon` | Sysmon ID 1 schtasks | Security 4698/4699 si auditing existe | MEDIUM | HIGH | Security no es obligatorio. |
| TEC-004 | Registry Run Keys / Startup Folder | T1547.001 | Persistence | `P1_EVENT_Sysmon_ID1_Process_Basic` | `Forensic_EVENT_Sysmon_ID1_Process_By_Technique` + registry context | `TEC004_RunKey_Sysmon` | Sysmon ID 1 reg.exe | Sysmon 12/13/14 si existe | MEDIUM | HIGH | Run/RunOnce por proceso y contexto registry. |
| TEC-005 | Service Execution | T1569.002 | Execution | `P1_EVENT_System7045_ServiceCreation` | `Forensic_EVENT_System_7045_ServiceCreation` | `TEC005_Service_System7045` | System 7045 | Sysmon ID 1 sc.exe | HIGH | HIGH | Service creation es evidencia primaria. |
| TEC-006 | Security Software Discovery | T1518.001 | Discovery | `P1_EVENT_PowerShell4104_Basic` + process basic | `Forensic_EVENT_PowerShell4104_Strong_By_Technique` + Sysmon ID1 process | `TEC006_SecuritySoftwareDiscovery_4104` + `TEC006_SecuritySoftwareDiscovery_Sysmon` | PowerShell 4104 | Sysmon ID 1 wmic/powershell | HIGH por 4104, MEDIUM por Sysmon | MEDIUM | P3 incluye fallback Sysmon para evitar repetir el fallo de export sin source TEC-006. |
| TEC-007 | Data Encrypted for Impact | T1486 | Impact | `P1_EVENT_PowerShell4104_Basic` | `Forensic_EVENT_PowerShell4104_Strong_By_Technique` + file context | `TEC007_DataEncryptedForImpact_4104` | PowerShell 4104 crypto/AES | Sysmon ID 11 `.aes` | HIGH | HIGH | No depende solo de Sysmon ID 11. |
| TEC-008 | Data Destruction | T1485 | Impact | `P1_EVENT_PowerShell4104_Basic` | `Forensic_EVENT_PowerShell4104_Strong_By_Technique` + file delete context | `TEC008_DataDestruction_4104` | PowerShell 4104 Get-ChildItem + Remove-Item | Sysmon 23/26 solo forense | HIGH | HIGH | Se excluye ruido Windows/Microsoft observado en PRUEBA_15_06_2026. |
| TEC-009 | Staging / Archive / Controlled HTTP ZIP Upload | T1074.001 / T1560.001 / T1048.003 | Collection / Exfiltration | `P1_EVENT_PowerShell4104_Basic` + `P1_EVENT_TEC009_Sysmon_Context` | `Forensic_EVENT_PowerShell4104_Strong_By_Technique` + Sysmon context | `TEC009_HTTP_ZIP_Upload_4104` + context sources | PowerShell 4104 HTTP tool + file upload + ZIP/upload | Sysmon ID 1, ID 3 receiver, ID 11 ZIP, receiver log externo | HIGH por 4104, MEDIUM por contexto | HIGH | Sysmon ID 3 no es condicion obligatoria. `POST` es evidencia positiva pero no bloqueante si el shape fuerte existe. |

## Justificacion de Confidence

- HIGH: evidencia fuerte en PowerShell 4104 o System 7045 para servicio.
- MEDIUM: proceso Sysmon ID 1 compatible o contexto Sysmon ID 3/11/12/13/14/23/26.
- LOW: no se usa como salida principal en la version final; los indicios debiles quedan fuera de P3 y en P1/P2 solo si aportan contexto defendible.

## Justificacion de Severity

- MEDIUM: ejecucion o discovery cuyo impacto depende del payload.
- HIGH: persistencia, servicio, impacto sobre ficheros, cifrado, destruccion o transferencia HTTP ZIP controlada.

## Validacion pendiente

Codex no valida experimentalmente. La validacion final consiste en activar los artifacts `CLIENT_EVENT`, ejecutar runner v5 en la VM y confirmar eventos en Velociraptor.

