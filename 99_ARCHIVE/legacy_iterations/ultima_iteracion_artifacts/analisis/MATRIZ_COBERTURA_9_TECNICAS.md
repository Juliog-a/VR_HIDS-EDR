# Matriz de cobertura 9 tecnicas

Fuente oficial principal: `04_EVIDENCE/Analisis_Tecnicas_TFM_Velociraptor_DEF.xlsx`.

Nota TEC-009: el Excel define staging/archive host-based. La campana v4 actual anade upload HTTP ZIP controlado; se conserva el mapeo base del Excel y se anade `T1048.003` como contexto operativo de la iteracion v4.

| TEC | Nombre | MITRE_ID | Tactica | Evidencia principal | Evidencia secundaria | P1 | P2 | P3 | Confidence esperada | Severity/Criticidad | Observaciones |
|-----|--------|----------|---------|--------------------|---------------------|----|----|----|--------------------|---------------------|---------------|
| TEC-001 | PowerShell | T1059.001 | Execution | 4104 con ExecutionPolicy Bypass, FromBase64String, IEX o Invoke-Expression | Sysmon ID 1 powershell.exe | Si | Si | Si | HIGH por 4104 fuerte | MEDIUM | Criticidad media: tecnica comun de ejecucion, impacto depende del payload. |
| TEC-002 | Windows Command Shell | T1059.003 | Execution | Sysmon ID 1 con cmd.exe /c, chaining o PowerShell hijo | 4104 si hay PowerShell hijo | Si | Si | Si | MEDIUM | MEDIUM | Criticidad media: shell generica, requiere contexto de comando. |
| TEC-003 | Scheduled Task | T1053.005 | Persistence | Sysmon ID 1 con schtasks /create, /run, /delete o /change | Security 4698/4699 si auditing existe | Si | Si | Si | MEDIUM | HIGH | Criticidad alta por persistencia. Security no es obligatorio. |
| TEC-004 | Registry Run Keys / Startup Folder | T1547.001 | Persistence | Sysmon ID 1 reg.exe add/delete sobre Run/RunOnce | Sysmon 12/13/14 si existen | Si | Si | Si | MEDIUM | HIGH | Criticidad alta por autostart. Registry events enriquecen. |
| TEC-005 | Service Execution | T1569.002 | Execution | System 7045 Service Control Manager | Sysmon ID 1 sc.exe create/start/delete | Si | Si | Si | HIGH por 7045 | HIGH | Criticidad alta por ejecucion mediante servicio. |
| TEC-006 | Security Software Discovery | T1518.001 | Discovery | 4104 con SecurityCenter2, AntivirusProduct, Defender o WMI/CIM | Sysmon ID 1 wmic/powershell | Si | Si | Si | HIGH por 4104 | MEDIUM | Criticidad media: discovery defensivo sin desactivacion. |
| TEC-007 | Data Encrypted for Impact | T1486 | Impact | 4104 con System.Security.Cryptography, AES, CryptoStream, CreateEncryptor o .aes | Sysmon ID 11 .aes si existe | Si | Si | Si | HIGH por 4104 fuerte | HIGH | No depende solo de Sysmon ID 11. |
| TEC-008 | Data Destruction | T1485 | Impact | 4104 con Get-ChildItem + Remove-Item + -Force/-Recurse o pipeline | Sysmon ID 23/26 si existe | Si | Si | Si | HIGH por 4104 fuerte | HIGH | Sysmon ID 23/26 es forense, no condicion principal. |
| TEC-009 | Staging / Archive / Controlled HTTP ZIP Upload | T1074.001 / T1560.001 / T1048.003 context | Collection / Exfiltration | 4104 con Invoke-WebRequest + -InFile + POST + application/zip/.zip/upload | Sysmon ID 11 ZIP, Sysmon ID 3 receiver, receiver log externo | Si | Si | Si | HIGH por 4104 fuerte | HIGH | Sysmon ID 3 solo contexto; si no aparece red, la alerta 4104 sigue siendo valida. |

## Justificacion de criticidad

- MEDIUM: ejecucion o discovery frecuente, con riesgo dependiente del payload o contexto.
- HIGH: persistencia, servicio, impacto sobre ficheros o exfiltracion/staging con transferencia controlada.
- Confidence no mide impacto. Confidence mide calidad de la evidencia observada.
