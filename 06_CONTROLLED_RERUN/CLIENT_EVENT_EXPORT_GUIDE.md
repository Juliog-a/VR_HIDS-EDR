# Exportación final CLIENT_EVENT

## Regla de evidencia

Exportar desde **Client Event Monitoring** del cliente
`C.364907172793f8c7`. No usar SOC_v3, `soc_alerts.jsonl`, Discord ni Server
Events como sustituto. La fila original `CLIENT_EVENT` es la detección; el
router solo la persiste después y puede introducir retraso temporal.

Antes de exportar, confirmar en la GUI que el identificador visible del equipo
LT29 sigue siendo `C.364907172793f8c7`.

## Ventanas UTC

- TEC FINAL02: `2026-06-19T10:32:49.8183942Z` a
  `2026-06-19T10:39:29.7482579Z`.
- FP REP_01: `2026-06-19T10:44:36.1731984Z` a
  `2026-06-19T10:45:37.0757196Z`.
- FP REP_02: `2026-06-19T10:45:57.3156536Z` a
  `2026-06-19T10:46:49.9924916Z`.
- FP REP_03: `2026-06-19T10:47:10.1043434Z` a
  `2026-06-19T10:48:11.1502438Z`.

Para FP se puede hacer una sola exportación amplia de todas las sources entre
`2026-06-19T10:44:00Z` y `2026-06-19T10:49:00Z`. Debe realizarse después de
terminar REP_03. El helper separará las tres repeticiones mediante
`DetectionTime`, no mediante la hora posterior del router.

## Sources que deben exportarse

### P1 Critical

Artifact: `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1`

- `P1_EVENT_PowerShell4104_Basic`
- `P1_EVENT_Sysmon_ID1_Process_Basic`
- `P1_EVENT_System7045_ServiceCreation`
- `P1_EVENT_TEC009_Sysmon_Context`

### P2 Event

Artifact: `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`

- `Forensic_EVENT_PowerShell4104_Strong_By_Technique`
- `Forensic_EVENT_Sysmon_ID1_Process_By_Technique`
- `Forensic_EVENT_System_7045_ServiceCreation`
- `Forensic_EVENT_Sysmon_Registry_File_Network_Context`
- `Forensic_EVENT_Security_ScheduledTask_Optional`

### P3 Event

Artifact: `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`

- `TEC001_PowerShell_4104`
- `TEC002_CMD_Sysmon`
- `TEC003_ScheduledTask_Sysmon`
- `TEC004_RunKey_Sysmon`
- `TEC005_Service_System7045`
- `TEC006_SecuritySoftwareDiscovery_4104`
- `TEC006_SecuritySoftwareDiscovery_Sysmon`
- `TEC007_DataEncryptedForImpact_4104`
- `TEC008_DataDestruction_4104`
- `TEC009_HTTP_ZIP_Upload_4104`
- `TEC009_HTTP_Upload_Sysmon_Process_Context`
- `TEC009_HTTP_Network_Sysmon_Context`
- `TEC009_ZIP_FileCreate_Sysmon_Context`

### P4

Artifact: `Custom.TFM.HIDS.P4.Low.Basic_v2`

- `CU001_Source_TEC001_PowerShell_Basic`
- `CU002_Source_TEC002_CMD_Basic`
- `CU003_Source_TEC003_ScheduledTask_Basic`
- `CU004_Source_TEC004_RunKeys_Basic`
- `CU005_Source_TEC005_ServiceCreation_7045`

## Procedimiento GUI

1. Abrir el cliente LT29 en Velociraptor.
2. Abrir `Client Events` / `Event Monitoring`.
3. Seleccionar cada artifact y cada source de la lista anterior.
4. Fijar la ventana UTC indicada. No usar una ventana relativa como “última
   hora” si se va a conservar como evidencia final.
5. Exportar CSV incluso cuando la source tenga cero filas; debe conservarse su
   cabecera.
6. Para FP, dejar los 27 CSV originales en:

```text
C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN\OUTPUT\FP\FP_20260619_FINAL01\CLIENT_EVENT_RAW_ALL_REPS
```

No renombrar las sources ni editar sus filas.

## Consolidación automática en la VM

TEC FINAL02 ya dispone de exports por source. Crear los cuatro CSV canónicos:

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\10_Build_ClientEvent_Exports_v2.ps1 `
  -CampaignDirectory ".\OUTPUT\TEC\TEC_20260619_FINAL02" `
  -SourceDirectory ".\OUTPUT\TEC\TEC_20260619_FINAL02" `
  -ReplaceDerivedOutputs
```

Para FP, después de exportar las 27 sources:

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
$fp = ".\OUTPUT\FP\FP_20260619_FINAL01"
$raw = Join-Path $fp "CLIENT_EVENT_RAW_ALL_REPS"

1..3 | ForEach-Object {
    $rep = "REP_{0:D2}" -f $_
    .\10_Build_ClientEvent_Exports_v2.ps1 `
      -CampaignDirectory (Join-Path $fp $rep) `
      -SourceDirectory $raw `
      -ReplaceDerivedOutputs
}
```

Verificación independiente obligatoria:

```powershell
.\11_Verify_ClientEvent_Exports.ps1 `
  -FpCampaignDirectory ".\OUTPUT\FP\FP_20260619_FINAL01"
```

Cada repetición debe terminar con cuatro ficheros en su carpeta
`velociraptor_exports`:

```text
P1_CRITICAL_CLIENT_EVENT.csv
P2_EVENT_CLIENT_EVENT.csv
P3_EVENT_CLIENT_EVENT.csv
P4_CLIENT_EVENT.csv
```

El helper no sobrescribe esos ficheros, filtra por `DetectionTime`, añade
`Artifact` y `Source`, y genera `CLIENT_EVENT_EXPORT_BUILD.csv/json` con
recuentos y hashes.

## Transferencia selectiva final

En la VM:

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\07_Export_VM_Results.ps1 `
  -TransferId "RERUN_20260619_FINAL02" `
  -TecCampaignId "TEC_20260619_FINAL02" `
  -FpCampaignId "FP_20260619_FINAL01" `
  -BenchmarkCampaignId "BENCH_20260619_FINAL01"
```

En el host:

```powershell
cd <RAIZ_TFM>\06_CONTROLLED_RERUN
.\08_Import_VM_Results.ps1 -TransferId "RERUN_20260619_FINAL02"
```

La importación verifica hashes, conserva ficheros ya idénticos y solo añade
ficheros nuevos. Cualquier conflicto de hash aborta; no sobrescribe campañas.
