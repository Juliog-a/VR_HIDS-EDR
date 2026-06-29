# Guia P1/P2 Critical Recalibration v1

## Objetivo

Reducir ruido critico sin perder visibilidad global:

- P4 mantiene visibilidad amplia / low fidelity.
- P3 mantiene comportamiento sospechoso / medium.
- P2 mantiene evidencia forense fuerte / high confidence.
- P1 queda reservado para prioridad critica o alta casi confirmada.

## Cambios aplicados

- Se recalibra solo `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1`.
- Se mantiene el mismo artifact name interno.
- Se mantiene el mismo nombre de fichero.
- Se mantienen los cuatro source names que consume `SOC_v3`.
- No se modifica runner, receiver, scripts TEC, P3, P4 ni SOC_v3.
- P2 queda sin cambios funcionales.

## Sources P1 conservadas

- `P1_EVENT_PowerShell4104_Basic`
- `P1_EVENT_Sysmon_ID1_Process_Basic`
- `P1_EVENT_System7045_ServiceCreation`
- `P1_EVENT_TEC009_Sysmon_Context`

## Condiciones que activan P1

### TEC-007

PowerShell 4104 debe contener:

- `System.Security.Cryptography` o `Aes`
- `CryptoStream` o `CreateEncryptor`
- `.aes`

### TEC-008

PowerShell 4104 debe contener:

- `Get-ChildItem` o alias equivalente
- `Remove-Item` o comando equivalente de borrado
- `-Force` o `-Recurse`

### TEC-009

PowerShell 4104 debe contener:

- `Invoke-WebRequest`, `iwr`, `wget`, `curl`, `Invoke-RestMethod`, `WebClient` o `UploadFile`
- `POST` o `-Method POST`
- `-InFile` o forma equivalente de subida de fichero
- `.zip` o `application/zip`
- `upload` o `ReceiverUrl`

Sysmon solo puede aportar P1 en casos de alta fidelidad:

- Sysmon ID 3: PowerShell hacia `ReceiverIP:ReceiverPort`.
- Sysmon ID 11: PowerShell creando un ZIP.

### TEC-004

Sysmon ID 1 debe mostrar:

- `reg.exe add`
- `CurrentVersion\Run` o `RunOnce`
- `/d`
- destino ejecutable o script.

### TEC-005

System 7045 o Sysmon ID 1 debe mostrar creación de servicio con:

- `cmd.exe`, `powershell.exe`, `pwsh.exe`, `.ps1`, `/c` o `-Command`.

## Exclusiones fuertes

P1 excluye autoeventos del pipeline SOC:

```text
(?i)(soc_alerts\.jsonl|alerts\.jsonl|JSONL_WRITE_ATTEMPT|DISCORD_POST_ATTEMPT|DiscordWebhook|RouterArtifact|Custom\.TFM\.HIDS\.Router|ConvertTo-Json|ConvertFrom-Json|Add-Content|TFM Velociraptor SOC|webhook)
```

## Validacion recomendada

1. Importar el P1 recalibrado.
2. Mantener activos P2, P3, P4 y SOC_v3.
3. Limpiar `soc_alerts.jsonl`.
4. Ejecutar runner v6.
5. Comprobar distribucion:

```powershell
$path = "\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl"
$rows = Get-Content $path | Where-Object { $_.Trim() -ne "" } | ConvertFrom-Json

$rows | Group-Object Profile | Sort-Object Count -Descending | Select-Object Count,Name
$rows | Group-Object TEC | Sort-Object Count -Descending | Select-Object Count,Name

$rows |
Where-Object {
  $_.Profile -eq "P1_CRITICAL" -and
  ($_.IOA -match "DiscordWebhook|RouterArtifact|soc_alerts|ConvertFrom-Json|ConvertTo-Json|Add-Content")
} |
Measure-Object
```

Resultado esperado:

- P1_CRITICAL baja de forma drastica.
- P1_CRITICAL ya no contiene TEC-001/TEC-002/TEC-003/TEC-006 genericos.
- P2/P3/P4 mantienen la vision 9/9.
- SOC_v3 no necesita cambios porque los source names se mantienen.
