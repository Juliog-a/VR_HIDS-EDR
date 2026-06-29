# Guia Router SOC Discord v2

Fecha: 2026-06-16

## Artifact

- `01_ARTIFACTS/candidate/last_version/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v2.yaml`

## Objetivo

Router `SERVER_EVENT` para consumir eventos ya generados por artifacts
`CLIENT_EVENT` P1/P2/P3/P4 mediante `watch_monitoring()`, normalizarlos como
alerta SOC, escribir JSONL y enviar opcionalmente a Discord.

El router no detecta telemetria del host. La deteccion la hacen P1/P2/P3/P4.
JSONL y Discord son salida/persistencia/notificacion.

## Correcciones frente a SOC_v1

- Anade `EnableP4`.
- Escucha P4 v2.
- Escucha P2 desplegado como `Custom.TFM.HIDS.P2.High.Forensic_v1`, no solo
  `P2.High.Forensic.Event_v1`.
- Normaliza columnas reales de P2/P4:
  - `AlertName` -> `DetectionName`;
  - `CU_Evidence` -> `Evidence` / `IOA`;
  - `Channel` -> `Source`;
  - `CU_Reason` / `DetectionLogic` -> `RawEventSummary`.
- Mantiene soporte para P1 legacy `Low.Basic.Event_v1`.
- Permite P1 renombrado `Critical.Priority.Event_v1` mediante
  `EnableP1Critical=true`.
- Filtros de severidad/confianza no restringen por defecto.

## Parametros recomendados primera prueba

```text
EnableP1=true
EnableP1Critical=false
EnableP2=true
EnableP3=true
EnableP4=true
EnableJSONL=true
JsonlPath=\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl
EnableDiscord=false
SeverityRegex=(?i).*
ConfidenceRegex=(?i).*
IncludeContextOnly=true
ExcludeRouterSelfEvents=true
```

Solo activar Discord despues de validar JSONL:

```text
EnableDiscord=true
DiscordWebhook=<pegar_webhook_en_parametro>
DiscordUsername=TFM Velociraptor SOC
```

No guardar el webhook dentro del YAML.

## Checklist de validacion

1. Importar artifacts CLIENT_EVENT:
   - `Custom.TFM.HIDS.P1.Low.Basic.Event_v1` o `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1`;
   - `Custom.TFM.HIDS.P2.High.Forensic_v1`;
   - `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`;
   - `Custom.TFM.HIDS.P4.Low.Basic_v2`.
2. Activarlos en Client Event Monitoring para `Label: All`.
3. Importar `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v2`.
4. Activarlo en Server Event Monitoring.
5. Probar primero con `EnableDiscord=false`.
6. Limpiar JSONL antes de la prueba:

```powershell
$p='\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl'
if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
```

7. Ejecutar runner v6 en la VM:

```powershell
cd C:\Users\seguridad\Desktop\TFM
.\03_RUNNERS\TFM_Run_All_TEC_Tests_v6.ps1
```

8. Comprobar JSONL:

```powershell
Get-Content -LiteralPath '\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl' -Tail 20
```

9. Confirmar que aparecen campos:
   - `Timestamp`
   - `Hostname`
   - `ClientId`
   - `Artifact`
   - `Source`
   - `Profile`
   - `CU_ID`
   - `TEC`
   - `MITRE`
   - `IOA`
   - `IOC`
   - `Severity`
   - `Confidence`
   - `DetectionName`
   - `Evidence`
   - `RecommendedAction`
   - `RawEventSummary`

10. Activar Discord solo cuando JSONL funcione.

## Interpretacion de fallos

- Router con barra verde pero sin filas:
  - revisar que P1/P2/P3/P4 hayan emitido eventos despues de activar el router;
  - revisar nombres exactos artifact/source en Event Monitoring;
  - revisar si `EnableP2`, `EnableP3` o `EnableP4` estan desactivados.

- Filas del router con error `JSONL_WRITE_ATTEMPT`:
  - revisar permisos del usuario/servicio Velociraptor sobre `\\VBOXSVR`;
  - confirmar que el server Velociraptor puede ejecutar `powershell.exe`;
  - confirmar que la ruta UNC existe desde el servidor, no solo desde el cliente.

- Discord no envia:
  - validar primero JSONL;
  - comprobar `EnableDiscord=true`;
  - comprobar que `DiscordWebhook` empieza por `https://`;
  - revisar `DISCORD_POST_ATTEMPT` y `Stderr`.

## Nota P1

`Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` es un alias metodologico de
P1 con nombre corregido. Mantiene sources y logica para no romper compatibilidad.
No activar P1 legacy y P1 critical a la vez salvo que se acepten alertas duplicadas.
