# ARTIFACT_INDEX

Actualizado: 2026-07-10 10:26 CEST.

## Fuente de verdad

Los artifacts finales se encuentran en:

`01_ARTIFACTS/validated`

No sobrescribir esta carpeta.

| Artifact | Tipo | Estado | Uso |
|---|---|---|---|
| `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml` | `CLIENT_EVENT` | FINAL | Señales críticas de alta prioridad. |
| `Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml` | `CLIENT_EVENT` | FINAL | Detección y contexto forense. |
| `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml` | `CLIENT_EVENT` | FINAL | Comportamiento, normalización y alertabilidad. |
| `Custom.TFM.HIDS.P4.Low.Basic_v2.yaml` | `CLIENT_EVENT` | FINAL | Visibilidad básica. |
| `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml` | `SERVER_EVENT` | FINAL | Enrutamiento, normalización, JSONL y Discord opcional. No detecta. |

## Candidatos

`01_ARTIFACTS/candidate` contiene versiones candidatas, backups y variantes
superadas.

Estado: `NO_LEER_POR_DEFECTO`.

Solo consultar si el usuario pide explícitamente:

- desarrollo de una nueva versión;
- comparación con histórico;
- recuperación de lógica previa;
- depuración de un artifact candidate concreto.

## Debug

`01_ARTIFACTS/debug` contiene artifacts de diagnóstico.

Estado: `NO_LEER_POR_DEFECTO`.

No forman parte de resultados finales.

## Reglas

- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, normaliza, persiste o notifica.
- JSONL y Discord no son detección primaria.
- Sysmon ID 26 es evidencia forense, no alerta individual.
- No usar rutas de laboratorio, nombres de scripts, `DUMB_LAB`, `dump_exfil`,
  `datos_robados` ni etiquetas TEC como IOC.
