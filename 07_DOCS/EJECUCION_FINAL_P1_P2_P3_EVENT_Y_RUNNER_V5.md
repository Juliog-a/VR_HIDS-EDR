# EJECUCION_FINAL_P1_P2_P3_EVENT_Y_RUNNER_V5

## 1. Archivos que se deben usar

Artifacts `CLIENT_EVENT`:

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml`
- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml`

Runner:

- `03_RUNNERS/TFM_Run_All_TEC_Tests_v5.ps1`

Recogida de evidencia:

- `03_RUNNERS/TFM_Collect_Final_Evidence_v1.ps1`

## 2. Preparar receiver en el host

En el host donde esta el receiver:

```powershell
cd C:\Users\julio\Desktop\TFM\02_SCRIPTS\candidate
.\Start-TFMReceiver_v4.ps1
```

Comprobar desde la VM:

```powershell
Test-NetConnection 192.168.1.129 -Port 8088
```

Debe devolver `TcpTestSucceeded : True`.

## 3. Importar artifacts en Velociraptor

En la GUI de Velociraptor:

1. Ir a `View Artifacts`.
2. Importar los tres YAML `Event_v1`.
3. Confirmar que aparecen como:
   - `Custom.TFM.HIDS.P1.Low.Basic.Event_v1`
   - `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`
   - `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`
4. Confirmar `type: CLIENT_EVENT`.

## 4. Activar Client Event Monitoring

En Velociraptor:

1. Ir a `Server Events` / `Client Event Monitoring`.
2. Anadir los tres artifacts EVENT.
3. Parametros recomendados:

```text
ReceiverIP=192.168.1.129
ReceiverPort=8088
HostnameOverride=
ClientIdOverride=
```

4. Guardar la configuracion.
5. Esperar a que el cliente reciba la nueva configuracion.

Importante: al ser `CLIENT_EVENT`, deben estar activos antes de ejecutar el runner. No recuperan eventos pasados.

## 5. Verificar logging antes de ejecutar

Seguir:

```text
07_DOCS/CHECKLIST_PREVIO_LOGGING_VR.md
```

Comandos minimos:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName='Microsoft-Windows-PowerShell/Operational'
  Id=4104
  StartTime=(Get-Date).AddHours(-2)
} | Select-Object -First 5 TimeCreated,Id,Message

Get-WinEvent -FilterHashtable @{
  LogName='Microsoft-Windows-Sysmon/Operational'
  Id=1,3,11
  StartTime=(Get-Date).AddHours(-2)
} | Group-Object Id | Select-Object Name,Count
```

## 6. Ejecutar runner v5 en la VM

Copiar `TFM_Run_All_TEC_Tests_v5.ps1` a:

```text
C:\Users\seguridad\Desktop\TFM\runner_candidate
```

Ejecutar como administrador:

```powershell
cd C:\Users\seguridad\Desktop\TFM\runner_candidate
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
$env:TFM_BASEPATH = "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas"

.\TFM_Run_All_TEC_Tests_v5.ps1 `
  -ReceiverUrl "http://192.168.1.129:8088/upload" `
  -EnableExfilUpload $true `
  -ExfilTimeoutSec 60 `
  -KeepExfilArtifacts $true
```

El runner debe imprimir al final:

```text
CAMPAIGN FINISHED
Global result: OK/WARN/FAIL
Techniques OK/WARN/FAIL: x/y/z
TEC-009 EVIDENCE CHECK
```

## 7. Evidencias esperadas en runner v5

Para TEC-009:

```text
ReceiverReachable=True
HTTPStatus=200
UploadSucceeded=True
ZIP exists=OK
ZIP copy exists=OK
SHA256 present=OK
summary JSON exists=OK
PowerShell 4104 upload evidence=OK o WARN
```

Si `PowerShell 4104 upload evidence=WARN`, la transferencia puede haber sido correcta, pero no se puede validar P3 EVENT por 4104.

## 8. Tablas/sources que revisar en Velociraptor

P3 principal:

| TEC | Source P3 EVENT |
|---|---|
| TEC-001 | `TEC001_PowerShell_4104` |
| TEC-002 | `TEC002_CMD_Sysmon` |
| TEC-003 | `TEC003_ScheduledTask_Sysmon` |
| TEC-004 | `TEC004_RunKey_Sysmon` |
| TEC-005 | `TEC005_Service_System7045` |
| TEC-006 | `TEC006_SecuritySoftwareDiscovery_4104` y fallback Sysmon |
| TEC-007 | `TEC007_DataEncryptedForImpact_4104` |
| TEC-008 | `TEC008_DataDestruction_4104` |
| TEC-009 | `TEC009_HTTP_ZIP_Upload_4104` |
| TEC-009 contexto | `TEC009_HTTP_Upload_Sysmon_Process_Context`, `TEC009_HTTP_Network_Sysmon_Context`, `TEC009_ZIP_FileCreate_Sysmon_Context` |

P2 forensic:

- `Forensic_EVENT_PowerShell4104_Strong_By_Technique`
- `Forensic_EVENT_Sysmon_ID1_Process_By_Technique`
- `Forensic_EVENT_System_7045_ServiceCreation`
- `Forensic_EVENT_Sysmon_Registry_File_Network_Context`
- `Forensic_EVENT_Security_ScheduledTask_Optional`

P1 low:

- `P1_EVENT_PowerShell4104_Basic`
- `P1_EVENT_Sysmon_ID1_Process_Basic`
- `P1_EVENT_System7045_ServiceCreation`
- `P1_EVENT_TEC009_Sysmon_Context`

## 9. TEC-009: evidencias que deben aparecer

En `TEC009_HTTP_ZIP_Upload_4104`:

```text
Invoke-WebRequest
-Method POST
-InFile
application/zip
ReceiverUrl o upload o .zip
```

Notas:

- `POST` es evidencia positiva, pero el artifact no debe depender solo de una forma exacta.
- Sysmon ID 3 a `192.168.1.129:8088` es contexto.
- Sysmon ID 11 `.zip` es contexto.
- Si no aparece red pero aparece 4104 fuerte, TEC-009 debe considerarse detectada.

## 10. Exportar resultados

Exportar desde Velociraptor:

1. Eventos de P3 EVENT.
2. Eventos de P2 EVENT.
3. Eventos de P1 EVENT.
4. Notebook o CSV/JSON de sources TEC-009.
5. Capturas de la configuracion de Client Event Monitoring.

Guardar en:

```text
C:\Users\julio\Desktop\TFM\05_LOGS\FINAL_RUN_<timestamp>
```

## 11. Recoger evidencia local automatizada

En el host/proyecto:

```powershell
cd C:\Users\julio\Desktop\TFM
.\03_RUNNERS\TFM_Collect_Final_Evidence_v1.ps1 `
  -ProjectRoot "C:\Users\julio\Desktop\TFM" `
  -CampaignLogDir "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\Logs_Pruebas_TFM" `
  -Tec009Dir "C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion" `
  -ReceiverIP "192.168.1.129" `
  -ReceiverPort 8088 `
  -LookbackHours 4
```

Si se ejecuta dentro de la VM, ajustar `ProjectRoot` a una ruta existente o copiar despues la carpeta `FINAL_RUN_<timestamp>` al host.

## 12. Evidencias para memoria del TFM

Guardar:

- Runner v5 usado.
- Artifacts EVENT importados.
- Summary TXT/JSON/CSV del runner.
- `tec009_transfer_summary.json`.
- ZIP y ZIP copy de TEC-009.
- SHA256 del ZIP.
- `receiver_log.jsonl`.
- Export Velociraptor P3 EVENT.
- Export Velociraptor P2 EVENT.
- Capturas de Client Event Monitoring.
- `RESUMEN_EVIDENCIA_FINAL.md`.

