# Plan de validacion P4/P3 v2

## Objetivo

Validar una nueva campana limpia P4/P3 sin modificar artifacts validados, sin activar Discord y sin pasar a P2/P1.

## Preparacion

1. Copiar los archivos candidate al entorno de la VM manteniendo rutas equivalentes.
2. Importar artifacts candidate en Velociraptor.
3. Mantener Discord desactivado.
4. Limpiar o rotar `alerts.jsonl` antes de la campana.
5. Confirmar que Sysmon y PowerShell 4104 siguen activos.

## Orden de importacion

1. `artifacts_candidate/Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
2. `artifacts_candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v2.yaml`
3. `artifacts_candidate/Custom.TFM.HIDS.Router.JSONL.Discord.Final_v3.yaml`, solo si se decide enrutar artifacts `_v2`.

## Comandos de prueba en la VM

Desde PowerShell elevado:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM
.\runner_candidate\TFM_Run_All_TEC_Tests_v2.ps1
```

## Criterios de aceptacion

- El YAML importa correctamente.
- P4 v2 genera alertas esperadas para CU-001 a CU-005.
- P3 v2 genera alerta propia para CU-006.
- TEC-007 genera `.aes` antes de restaurar.
- CU-007 aparece por evidencia crypto/AES/CryptoStream/.aes.
- CU-008 aparece solo por patron destructivo real.
- CU-009 aparece solo por staging/archive y red contextual si se ejecuta `TcpClient`.
- CU-008 y CU-009 no se emiten sobre el ScriptBlock de ransomware.
- CU-001 reduce self-events del router.
- CU-002 no vuelve a alertar por Chrome/Adobe Native Messaging.
- JSONL se interpreta como salida/persistencia, no como deteccion.
- Router se interpreta como salida, no como detector.
- Discord permanece desactivado.

## Comprobaciones posteriores

```powershell
$alerts = Get-Content C:\Users\seguridad\Desktop\TFM\Carpeta_Compartida_TFM\inbox\alerts.jsonl | Where-Object { $_.Trim() } | ConvertFrom-Json
$alerts | Group-Object CU_ID | Sort-Object Name | Select-Object Name,Count
$alerts | Group-Object SourceArtifact | Sort-Object Count -Descending | Select-Object Name,Count
$alerts | Where-Object { $_.CU_ID -in @('CU-006','CU-007','CU-008','CU-009') } | Select-Object CU_ID,MITRE_ID,SourceArtifact,AlertName,CU_Reason
```

## Que no activar todavia

- Discord.
- P2.
- P1.
- Sysmon ID 26 como alerta individual.
- Conclusiones finales de deteccion.
- Cualquier comparacion operacional con EDR comercial.

## Frase para memoria

La iteracion P4/P3 v2 separa la logica de deteccion host-based de las capas de enrutado y persistencia, corrigiendo ruido operativo y priorizacion semantica sin depender de rutas, nombres de scripts ni etiquetas de laboratorio como indicadores.
