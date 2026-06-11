# CURRENT_TASK

## Tarea activa

Ejecutar backtest historico de CU-009 4104 con artifact `CLIENT` CU009Shape_v4 para confirmar que el 4104 real de `exfiltracion_v3.ps1` produce `CU009StrongCandidate_v4=True`.

## Artifact base

01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v8.yaml

## Artifact nuevo esperado

01_ARTIFACTS/debug/Custom.TFM.Debug.PowerShell4104.CU009Shape_v4.yaml

## Restricciones

- No tocar P1.

- No tocar P2.

- No tocar router.

- No activar Discord.

- No crear P3_v9 sin resultado del artifact debug.

- Mantener CU-006, CU-007 y CU-008.

- No usar rutas/scripts/lab markers como IOC.

## Criterio de exito

Ejecutar `Custom.TFM.Debug.PowerShell4104.CU009Shape_v4` como Hunt/Collection manual con `SinceMinutes=240`.

Resultado esperado sobre el 4104 real de TEC-009:

- `UploadHttpRegex=True`
- `UploadFileRegex=True`
- `UploadPostRegex=True`
- `UploadZipContentRegex=True`
- `HashEvidenceRegex=True`
- `CryptoEncryptionRegex=False`
- `DetectionEngineeringNoiseRegex=False`
- `DocumentationNoiseRegex=False`
- `ActualInvokeWebRequestShape=True`
- `ActualUploadCommandShape=True`
- `CU009StrongCandidate_v4=True`

`DestructionStrongRegex=True` puede aparecer por limpieza local con `Remove-Item`, pero no debe bloquear CU-009 en v4.

P3_v9 solo se creara cuando v4 confirme el resultado en laboratorio.
