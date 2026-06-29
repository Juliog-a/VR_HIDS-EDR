# Guia Router SOC Discord v1

Fecha: 2026-06-16

Artifact:

- `01_ARTIFACTS/candidate/Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v1.yaml`

## Objetivo

Router `SERVER_EVENT` para recoger alertas ya generadas por los artifacts
`CLIENT_EVENT` finales P1/P2/P3, normalizarlas como alerta SOC, persistirlas en
JSONL y enviarlas opcionalmente a Discord.

Este artifact no detecta tecnicas. La deteccion la hacen:

- `Custom.TFM.HIDS.P1.Low.Basic.Event_v1`
- `Custom.TFM.HIDS.P2.High.Forensic.Event_v1`
- `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1`

JSONL y Discord son canales de salida, no evidencia de deteccion por si solos.

## Parametros recomendados

Configuracion base:

```text
EnableP3=true
EnableP1=true
EnableP2=false
EnableJSONL=true
JsonlPath=\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl
EnableDiscord=false
DiscordWebhook=
SeverityRegex=(?i)(Critical|High|Medium)
ConfidenceRegex=(?i)(HIGH|High|MEDIUM|Medium)
IncludeContextOnly=true
ExcludeRouterSelfEvents=true
```

Para enviar a Discord:

```text
EnableDiscord=true
DiscordWebhook=<pegar_webhook_en_parametro>
DiscordUsername=TFM Velociraptor SOC
```

No guardar el webhook dentro del YAML.

## Orden de despliegue

1. Importar los artifacts `CLIENT_EVENT` P1/P2/P3.
2. Activar Client Event Monitoring para P1/P2/P3 antes de la campana.
3. Importar `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v1`.
4. Activarlo como Server Event Monitoring.
5. Configurar `JsonlPath`.
6. Si se quiere Discord, activar `EnableDiscord=true` y pegar `DiscordWebhook`.
7. Ejecutar runner v5.
8. Revisar:
   - Event Monitoring de P1/P2/P3.
   - Salida del router.
   - `soc_alerts.jsonl`.
   - Canal Discord si se ha configurado webhook.

## Formato SOC normalizado

Cada alerta incluye:

- `RouterTime`
- `DetectionTime`
- `SourceArtifact`
- `ClientId`
- `Hostname`
- `CU_ID`
- `ID_Tecnica_Interna`
- `MITRE_ID`
- `MITRE_Tactic`
- `SignalType`
- `AlertTitle`
- `AlertDescription`
- `Evidence`
- `IOA_Type`
- `IOA`
- `IOC`
- `Confidence`
- `ConfidenceScore`
- `Severity`
- `SeverityScore`
- `RecommendedAction`
- `RouterArtifact`

## Discord

El embed de Discord muestra:

- severidad y confianza;
- host y client;
- CU y tecnica;
- MITRE;
- fuente/event ID;
- artifact origen;
- IOA;
- IOC/evidencia;
- accion recomendada.

Los campos largos de IOA/IOC se recortan para evitar limites de Discord.

## Ruido y recomendaciones

`EnableP3=true` debe ser la base para alertabilidad.

`EnableP1=true` se mantiene como fallback util, especialmente si P3 no emite
alguna fila de TEC-009 pero P1 captura la evidencia 4104 fuerte.

`EnableP2=false` por defecto porque P2 es forense y puede emitir mas volumen.
Activarlo solo cuando se quiera contexto adicional para memoria/analisis.

`IncludeContextOnly=false` reduce ruido en Discord si solo se quieren alertas
primarias y no contexto Sysmon.

## Validacion pendiente

Validado localmente:

- `type: SERVER_EVENT`;
- usa `watch_monitoring()`;
- no usa `parse_evtx()` ni `watch_evtx()`;
- no contiene webhook hardcodeado;
- no contiene rutas/IOCs de laboratorio prohibidos;
- escucha sources existentes de P1/P2/P3 EVENT.

Pendiente:

- importar en Velociraptor;
- activar como Server Event Monitoring;
- confirmar que escribe `soc_alerts.jsonl`;
- confirmar envio a Discord con webhook real.
