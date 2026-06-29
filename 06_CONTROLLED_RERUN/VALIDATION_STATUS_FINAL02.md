# Estado de validación controlada FINAL02

Fecha: 2026-06-19

## Combinación canónica

- TEC: `TEC_20260619_FINAL02`.
- FP: `FP_20260619_FINAL01`.
- Benchmark: `BENCH_20260619_FINAL01`.
- Transferencia final pendiente: `RERUN_20260619_FINAL02`.

`TEC_20260619_FINAL01` se conserva como diagnóstico fallido y no forma parte
de esta combinación.

## Estado actual

| Bloque | Estado | Evidencia |
|---|---|---|
| TEC FINAL02 | APTO | Runner 9/9 OK; JSONL 185/0; transferencia TEC-009 HTTP 200 y SHA-256 origen/destino; CLIENT_EVENT 199 filas, 0 fuera de ventana, cobertura 9/9 |
| Benchmark FINAL01 | APTO CON OBSERVACIONES | 9/9 `summary.json`; SchemaVersion 1.1; todos VALID; cliente correcto; SERVER_GUI excluido; notepad no runner; tres cargas con RunnerObservedSamples > 0 |
| FP FINAL01 | NO APTO todavía | Faltan exports CLIENT_EVENT posteriores a FP; JSONL REP_02 contiene una detección tardía de REP_01 y REP_03 contiene 36/47 filas cuyo DetectionTime pertenece a REP_02; 19 UNKNOWN en REP_01 y REP_03 requieren adjudicación |

Los tres `RunnerStillRunningAtEnd=True` del benchmark son observaciones de
limpieza. No invalidan los runs porque el contrato exige
`RunnerObservedSamples > 0`, `ScenarioValidity=VALID` y cliente observado. Los
tres criterios se cumplen.

## Correcciones aplicadas al paquete

- `04_Validate_Controlled_Rerun.ps1`:
  - acceso seguro a propiedades ausentes, incluida `Name`;
  - excepciones benchmark convertidas en hallazgos, sin abortar el informe;
  - validación completa de los nueve runs;
  - `DetectionTime` prevalece sobre el timestamp posterior del router;
  - P2 FP se registra como WARN y requiere adjudicación, no como FAIL de
    ejecución;
  - exports CLIENT_EVENT validados por CSV, esquema, cronología y cobertura TEC.
- `07_Export_VM_Results.ps1`: transferencia selectiva por los tres CampaignId.
- `08_Import_VM_Results.ps1`: importación idempotente y aditiva; nunca
  sobrescribe un fichero con hash distinto.
- `10_Build_ClientEvent_Exports.ps1`: consolidación cronológica sin alterar las
  fuentes; rechaza un conjunto fuente cuyo último evento sea anterior a la
  campaña.
- `CLIENT_EVENT_EXPORT_GUIDE.md`: lista completa de 27 sources y comandos.

Bundle nuevo para VM:

```text
\\VBOXSVR\Carpeta_Compartida_TFM\controlled_rerun_bundle_validation_patch
```

Verificación local: 28/28 hashes del bundle correctos.

## Interpretación de P2 en FP

Los P2 observados son falsos positivos/visibilidad forense esperables en esta
prueba, no evidencia de ejecución maliciosa:

- `TEC-006` 4104: `Get-MpComputerStatus` carga script blocks generados del
  módulo Defender (`MSFT_Mp*`). La regla detecta discovery real, pero la acción
  es administrativa y benigna.
- `TEC-002`/`TEC-003` Sysmon ID 1: `cmd.exe`, `schtasks.exe` y comandos de
  administración son dual-use. La presencia del proceso aporta visibilidad,
  no intención maliciosa.
- El contexto P2/P3 de FP-009 procede de un GET benigno a `127.0.0.1`; no hay
  POST, `InFile`, ZIP ni transferencia de exfiltración.

No se deben eliminar estas señales antes de cerrar la medición FP. Para la
memoria deben contabilizarse como falsos positivos de baja especificidad o
visibilidad forense. P1=0 sigue siendo el criterio crítico.

Posibles mejoras posteriores, sin aplicarlas a los datos actuales:

- deduplicar 4104 por `ScriptBlockId`/invocación superior para no contar cada
  bloque generado del módulo Defender;
- marcar como `Context` el Sysmon ID 1 genérico y reservar `Detection` para
  combinaciones de comando, padre, acción y destino más específicas.

Modificar ahora las reglas para hacer desaparecer los P2 sesgaría la campaña
ya ejecutada. Primero se exporta y adjudica la evidencia original.

## Pendiente exacto

1. Instalar en la VM el bundle de validación sin tocar artifacts validados.
2. Exportar las 27 sources CLIENT_EVENT FP entre `10:44:00Z` y `10:49:00Z`,
   después de finalizar REP_03.
3. Ejecutar `10_Build_ClientEvent_Exports.ps1` para REP_01..REP_03.
4. Crear la transferencia selectiva `RERUN_20260619_FINAL02`.
5. Importar y ejecutar el validador con la combinación canónica.
6. Revisar la adjudicación de UNKNOWN usando los exports originales.
7. No regenerar Excel hasta cerrar FP como APTO o APTO CON OBSERVACIONES.

