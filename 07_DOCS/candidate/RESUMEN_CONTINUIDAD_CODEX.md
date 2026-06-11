# Resumen continuidad Codex - TFM Velociraptor P4/P3

## Estado actual

Proyecto base: `C:\Users\julio\Desktop\TFM`.

Objetivo de la iteracion: preparar una nueva campana limpia P4/P3 para el TFM Velociraptor como HIDS/DFIR, corrigiendo ruido y fallos de validacion sin sobrescribir originales.

No se han modificado artifacts validados ni scripts originales. No se ha activado Discord. No se ha pasado a P2/P1. No se ha ejecutado la nueva campana.

## Decisiones tomadas

- JSONL es salida/persistencia, no deteccion.
- El router es salida/enrutado, no detector.
- Discord debe seguir desactivado.
- Sysmon ID 26 no debe tratarse como alerta operativa individual; solo evidencia forense.
- TEC-009 se mantiene como staging/archivado con red contextual, no exfiltracion real sin transferencia demostrada.
- No se usan rutas de laboratorio, nombres de scripts, `DUMB_LAB` ni etiquetas TEC como IOC en los artifacts.
- P4/P3 siguen siendo la fase actual; P2/P1 quedan pendientes.

## Archivos leidos

- `Prompt\Prompt_Contexto_TFM_VR_v6.txt`
- `Prompt\Prompt_Codex_Siguiente_Iteracion_P4_P3_TFM_Velociraptor_v2.txt`
- `Prompt\Checklist_Uso_Codex_TFM_Velociraptor.txt`
- `Prompt\PLAN_ARTIFACTS.md`
- `Carpeta_Compartida_TFM\inbox\alerts.jsonl`
- `Evidencias\Logs_Pruebas_TFM\TFM_TEC_Run_FIXED_20260609_152425.log`
- `Carpeta_Compartida_TFM\TFM_Run_All_TEC_Tests.ps1`
- `Carpeta_Compartida_TFM\TEC-007_Ransomware\enumeracion.ps1`
- `Carpeta_Compartida_TFM\TEC-007_Ransomware\Scriptransom.ps1`
- `Carpeta_Compartida_TFM\TEC-008_Sabotaje\sabotaje.ps1`
- `Carpeta_Compartida_TFM\TEC-009_Exfiltracion\exfiltracion.ps1`
- `Artifacts\velociraptor_tfm_artifacts_finales_IMPORTABLE_YAML_ROOT\Custom.TFM.HIDS.P4.Low.Basic.yaml`
- `Artifacts\velociraptor_tfm_artifacts_finales_IMPORTABLE_YAML_ROOT\Custom.TFM.HIDS.P3.Medium.Behavioral.yaml`
- `Artifacts\Custom.TFM.HIDS.Router.JSONL.Discord.Final_v2.yaml`
- router candidate/validated existentes.

## Diagnostico clave

- `alerts.jsonl` tenia 78 alertas:
  - P4: 75
  - P3: 3
  - `CU-001=31`, `CU-002=36`, `CU-003=4`, `CU-004=3`, `CU-005=1`, `CU-006=0`, `CU-007=1`, `CU-008=1`, `CU-009=1`.
- TEC-007 no fue validable: `enumeracion.ps1` y `Scriptransom.ps1` usaban `Pruebas\Ransomware`, pero el runner ejecutaba `Pruebas\TEC-007_Ransomware`.
- No hubo CSV valido ni cifrado efectivo, aunque P3 detecto el ScriptBlock con crypto.
- CU-008 y CU-009 se clasificaron incorrectamente sobre el ScriptBlock de ransomware.
- CU-006 no aparecio como alerta propia.
- CU-001 tenia self-events del router por `Add-Content`, `FromBase64String`, `alerts.jsonl`.
- CU-002 tenia FP Chrome/Adobe Native Messaging:
  `chrome.exe -> cmd.exe -> WCChromeNativeMessagingHost.exe`.

## Archivos creados

- `artifacts_candidate\Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
- `artifacts_candidate\Custom.TFM.HIDS.P3.Medium.Behavioral_v2.yaml`
- `artifacts_candidate\Custom.TFM.HIDS.Router.JSONL.Discord.Final_v3.yaml`
- `scripts_candidate\enumeracion_v2.ps1`
- `scripts_candidate\Scriptransom_v2.ps1`
- `scripts_candidate\exfiltracion_v2.ps1`
- `runner_candidate\TFM_Run_All_TEC_Tests_v2.ps1`
- `docs_candidate\artifact_feedback_v2.md`
- `docs_candidate\PLAN_VALIDACION_P4_P3_v2.md`
- `docs_candidate\RESUMEN_CONTINUIDAD_CODEX.md`

## Cambios implementados

- P4 v2:
  - Exclusion acotada de self-events del router en CU-001.
  - Exclusion acotada del FP Chrome/Adobe en CU-002.
  - CU-003, CU-004 y CU-005 se mantienen.

- P3 v2:
  - CU-006 expuesto como deteccion contextual propia.
  - CU-007 exige evidencia crypto/AES/CryptoStream/.aes.
  - CU-008 exige patron destructivo real y excluye bloques crypto/staging.
  - CU-009 exige staging/archivado y trata red solo como contexto.
  - Prioridad semantica: CU-007 > CU-009 > CU-008 > CU-006 > CU-001.

- Scripts candidate:
  - `enumeracion_v2.ps1` genera CSV/TXT junto al script o en ruta parametrizada.
  - `Scriptransom_v2.ps1` usa parametros y `$PSScriptRoot`, valida CSV/password y genera `.aes`.
  - `exfiltracion_v2.ps1` hace `Compress-Archive + Copy-Item + TcpClient`, sin transferir ZIP.

- Runner candidate:
  - Usa scripts candidate.
  - Cuenta `.aes` antes de restaurar TEC-007.
  - Ejecuta TEC-009 con red contextual controlada.

- Router v3:
  - Solo necesario si se importan artifacts `_v2`.
  - Escucha sources versionadas.
  - Usa patron Base64 + `execve`.
  - Discord permanece `false`.

## Verificaciones realizadas

- Parseo PowerShell OK:
  - `scripts_candidate\enumeracion_v2.ps1`
  - `scripts_candidate\Scriptransom_v2.ps1`
  - `scripts_candidate\exfiltracion_v2.ps1`
  - `runner_candidate\TFM_Run_All_TEC_Tests_v2.ps1`
- No se ejecuto la campana.
- No se importaron YAML.
- `git status` no pudo ejecutarse porque `git` no esta disponible en este PowerShell.

## Pendiente antes de validar

- Copiar los candidate al entorno de la VM si no estan ya accesibles desde la ruta esperada.
- Importar YAML candidate en Velociraptor.
- Limpiar/rotar `alerts.jsonl`.
- Ejecutar runner candidate.
- Analizar nuevo JSONL.

## Orden recomendado de importacion

1. `artifacts_candidate\Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
2. `artifacts_candidate\Custom.TFM.HIDS.P3.Medium.Behavioral_v2.yaml`
3. `artifacts_candidate\Custom.TFM.HIDS.Router.JSONL.Discord.Final_v3.yaml`, solo si se quiere enrutar los artifacts `_v2`.

## Comando de campana propuesto

Desde PowerShell elevado en la VM:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM
.\runner_candidate\TFM_Run_All_TEC_Tests_v2.ps1
```

## Criterios de aceptacion

- CU-006 aparece como alerta propia.
- TEC-007 genera `.aes` antes de restaurar.
- CU-007 aparece por evidencia crypto.
- CU-008 aparece solo por destruccion real.
- CU-009 aparece solo por staging/archivado/red contextual.
- CU-008 y CU-009 no se emiten sobre el ScriptBlock de ransomware.
- Baja el ruido CU-001/CU-002.
- No aparece el FP Chrome/Adobe Native Messaging.
- Router y JSONL se documentan como salida, no deteccion.

## No activar todavia

- Discord.
- P2.
- P1.
- Sysmon ID 26 como alerta individual.
- Conclusiones finales de deteccion.
- Comparacion experimental con EDR comercial.

