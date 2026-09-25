# GUIA ROUTER SOC DISCORD v3

## Objetivo

`Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3` es un `SERVER_EVENT` de routing. No detecta telemetria host. Solo consume eventos ya emitidos por artifacts `CLIENT_EVENT` mediante `watch_monitoring()`, normaliza las filas y las persiste en JSONL o las envia a Discord si se habilita.

## Diagnostico de SOC_v2

SOC_v2 escribia JSONL y Discord funcionaba, pero en la prueba solo aparecia `P1_CRITICAL`.

Causas probables:

- `SOC_v2` escuchaba P2 legacy `Custom.TFM.HIDS.P2.High.Forensic_v1`, mientras el despliegue actual usa `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`.
- `SOC_v2` agrupaba muchos `watch_monitoring()` vivos dentro de una sola source con `chain()`. En monitorizacion continua, la primera rama activa puede quedar viva y bloquear el avance hacia ramas posteriores.
- Los perfiles P2/P4 no usan siempre las mismas columnas que P1/P3, por lo que la normalizacion debe ser especifica y tolerante.

SOC_v3 corrige esto con una source independiente por cada source real monitorizada:

- P1 Critical: 4 sources.
- P2 Event: 5 sources.
- P3 Event: 13 sources.
- P4 v2: 5 sources.

## Artifacts que deben estar activos como Client Event Monitoring

En Label All:

```text
Custom.TFM.HIDS.P1.Critical.Priority.Event_v1
Custom.TFM.HIDS.P2.High.Forensic.Event_v1
Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1
Custom.TFM.HIDS.P4.Low.Basic_v2
```

## Parametros recomendados para primera validacion

```text
EnableP1Critical=true
EnableP1=false
EnableP2Event=true
EnableP2Legacy=false
EnableP3=true
EnableP4=true
EnableJSONL=true
JsonlPath=\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl
EnableDiscord=false
SeverityRegex=(?i).*
ConfidenceRegex=(?i).*
IncludeContextOnly=true
ExcludeRouterSelfEvents=false
```

Activa Discord solo despues de comprobar JSONL:

```text
EnableDiscord=true
DiscordWebhook=<pegar_webhook_al_desplegar>
```

## Checklist

1. Importar `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml`.
2. Desactivar `SOC_v2` en Server Event Monitoring.
3. Confirmar que los Client Events anteriores estan activos en Label All.
4. Activar `SOC_v3` como Server Event Monitoring.
5. Limpiar o renombrar `soc_alerts.jsonl`.
6. Ejecutar `03_RUNNERS/validate/TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1`.
7. Esperar 1-2 minutos tras terminar la campana.
8. Comprobar perfiles, artifacts y tecnicas.
9. Guardar JSONL, hash y capturas.
10. Activar Discord solo cuando JSONL sea correcto.

## Comandos de validacion

```powershell
$path = "\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl"
$rows = Get-Content $path | Where-Object { $_.Trim() -ne "" } | ConvertFrom-Json

$rows | Group-Object Profile | Select-Object Name,Count

$rows |
  Group-Object Artifact |
  Sort-Object Count -Descending |
  Select-Object Count,Name

1..9 | ForEach-Object {
  $id = "{0:D3}" -f $_
  $hits = @(Select-String -Path $path -Pattern "TEC-$id|TEC$id|CU-$id|CU$id")
  [pscustomobject]@{
    ID = $id
    Hits = $hits.Count
  }
}

$rows |
  Select-Object Timestamp, CU_ID, TEC, Profile, Severity, Confidence, DetectionName, Artifact |
  Sort-Object TEC, Profile |
  Format-Table -AutoSize
```

## Resultado esperado

`Group-Object Profile` debe mostrar mas de un perfil. En la validacion amplia deberian aparecer, si los CLIENT_EVENT generaron filas:

```text
P1_CRITICAL
P2_EVENT
P3
P4
```

Para visibilidad 9/9, el conteo por tecnica debe devolver al menos una fila para `001` a `009`. Si una tecnica no aparece en JSONL pero aparece en la GUI del Client Event correspondiente, el problema sigue estando en el router. Si tampoco aparece en la GUI, el problema esta en el artifact detector o en la telemetria.

## Interpretacion

- JSONL con `ReturnCode=0` valida persistencia del router.
- Discord con `DISCORD_POST_ATTEMPT ReturnCode=0` valida notificacion externa.
- La deteccion sigue siendo la fila original del `CLIENT_EVENT`.
- `SOC_v3` no debe usarse para afirmar deteccion si no existen filas P1/P2/P3/P4 aguas arriba.

