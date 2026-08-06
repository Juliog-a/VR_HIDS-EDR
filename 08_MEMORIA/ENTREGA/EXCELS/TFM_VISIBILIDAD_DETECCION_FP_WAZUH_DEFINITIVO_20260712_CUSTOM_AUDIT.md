# Auditoría custom final del Excel de visibilidad, detección y FP

Fecha de cierre: 2026-07-13 (Europe/Madrid)

## Fichero canónico

- Ruta: `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`
- Copia inicial seleccionada: maestro actual de `ENTREGABLE`, no copias auxiliares ni backups anteriores.
- Tamaño inicial: 316.006 bytes.
- SHA-256 inicial: `311DDEBEA38F523FFB63738017371121838EC4B108F549960E8B52A2484C4761`.
- Tamaño final: 311.995 bytes.
- SHA-256 final: `4877A0391822495DEC4EFF6C987964A6EC59B4D538B598860B5E6C7D508A6D0F`.
- Hojas finales: 30.
- Tablas Excel: 15.
- Gráficos nativos: 11.

La copia previa se conserva como:

`TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_PRE_CUSTOM_REVIEW.xlsx`

Tamaño 316.006 bytes; SHA-256 `311DDEBEA38F523FFB63738017371121838EC4B108F549960E8B52A2484C4761`.

## Artifacts custom auditados

| Perfil / rol | Nombre exacto declarado | SHA-256 | Sources declaradas | Resultado final |
|---|---|---|---:|---|
| P1 | `Custom.TFM.HIDS.P1.Critical.Priority.Event_v1` | `8A1FF2654405AF9E4412D32EF0C2BE08A383924654AA764A04AB264BAA42EDD7` | 4 | 19 alertas; 5/9 TEC |
| P2 | `Custom.TFM.HIDS.P2.High.Forensic.Event_v1` | `255C5F7CA39E362FE45106C662C4C35730EFF57A4063A1E9973FD11A6375FAA2` | 5 | 118 alertas; 7/9 TEC |
| P3 | `Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1` | `C591CB849753194F6F87834C1E0F34EDE6CDFF6B31BC63650A1D281ADA1D410A` | 13 | 109 alertas; 9/9 TEC |
| P4 | `Custom.TFM.HIDS.P4.Low.Basic_v2` | `1EBB29296A14475E38BCF34CA8DD9CC16572F4A8CA00A1AB0C3608E0B3097BC1` | 5 | 133 alertas; 5/9 TEC |
| Transporte | `Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3` | `574E3E8BC03E7C4D655400D5BB5527771E438CA67C413853A0AAAC7A10C5E33F` | 27 routes | SERVER_EVENT; no detector |

Se identificaron 27 sources detectoras P1-P4. Veinticinco generaron filas en la ventana final. No se atribuyeron alertas a `Forensic_EVENT_Security_ScheduledTask_Optional` ni a `TEC009_ZIP_FileCreate_Sysmon_Context`. El resultado por source y técnica contiene 35 grupos trazables. No existe fichero P0 en `01_ARTIFACTS/validated`; no se inventó artifact ni ejecución.

## Reconciliación de resultados

- Campaña TEC: 9 ejecutadas, 9 OK, 9/9 detectadas por Velociraptor custom.
- Alertas CLIENT_EVENT: P1=19, P2=118, P3=109 y P4=133; suma=379.
- Campaña FP custom: 10 ejecutadas, 10 OK y 0 hits.
- Hayabusa Monitoring CH: 3/9 técnicas; FP público separado (4 High, 3 eventos, un caso FP-003).
- Wazuh base: 3/9 técnicas.
- Wazuh custom: 4/9 técnicas.
- Campañas públicas: 5 POSITIVO y 1 NO CONCLUYENTE; ETW no se convirtió en negativo válido.
- TEC-009: TrackNetwork aporta visibilidad de conexión; HTTP 200 se acredita exclusivamente mediante runner independiente.

## Hojas modificadas

- `00_Guia`.
- `01_Dashboard`.
- `03_Evaluacion_VR`.
- `06_Resumen`.
- `07_Catalogo_Artifacts`.
- `09_Benchmark_Plan`, sustituida y renombrada como `09_Resultados_Custom`.
- `10_Gaps_Custom_P1P4`.
- `14_Arquitectura_Custom`.
- `README`.
- `RESUMEN_EJECUTIVO`.
- `VISIBILIDAD_SISTEMA`.
- `COMPARACION_VR_WAZUH`.
- `DISCREPANCIAS`.
- `FUENTES`.
- `GRAFICAS`.

## Sustitución de la hoja 09

`09_Benchmark_Plan` se eliminó como hoja de contenido y se renombró a `09_Resultados_Custom`. Contiene:

1. resumen por perfil;
2. matriz TEC × perfil;
3. resultados por source interna y técnica;
4. campañas TEC y FP custom.

No contiene medidas de CPU, memoria, E/S, consumo por proceso ni escenarios de rendimiento.

## Gráficos

Se reutilizaron los 11 gráficos existentes; no se crearon duplicados.

- Dashboard: cobertura táctica, cobertura TEC por sistema, detección específica por sistema, capas CLIENT_EVENT, FP público y estado de campañas públicas.
- `GRAFICAS`: visibilidad host por técnica, alertas por perfil, campañas TEC/FP separadas, Wazuh base/custom por técnica y alertas custom por técnica.
- La matriz TEC × perfil y las capacidades HIDS se presentan como heatmaps vinculados a datos reales.
- La salida externa no figura como capacidad directa del detector.

## Exclusión del análisis de rendimiento

Se eliminaron o sustituyeron las referencias analíticas de rendimiento en guía, resumen, hoja 09, gaps, arquitectura, catálogo, Dashboard, README, nombres, fórmulas y gráficos. El validador final registra 0 coincidencias prohibidas fuera de las notas explícitas de exclusión y de la evidencia cruda que debía conservarse.

## Comparación con los Excel antiguos

`TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx` y `Analisis_Tecnicas_TFM_V.5.xlsx` son binariamente idénticos: 282.039 bytes y SHA-256 `F9CD06E7E986102028C7FE28523BEAD774ED8E724FD7464382CF2C91A45293C1`.

No se recuperaron métricas: sus estados `Pendiente` y su contenido de rendimiento están superados. La única hoja exclusiva, `99_Listas`, está vacía y no aporta datos, fuentes ni notas metodológicas nuevas.

## Validación final

- ZIP/OpenXML válido: PASS.
- Errores de parseo XML: 0.
- Errores de fórmula o valores almacenados: 0.
- Referencias a hojas inexistentes: 0.
- Enlaces externos: 0.
- Gráficos vacíos: 0.
- Series con referencias rotas: 0.
- Celdas con valor exacto `Pendiente`: 0.
- Referencias analíticas de rendimiento prohibidas: 0.
- Hipervínculos internos del Dashboard: 10.
- Excel COM: cálculo automático, `CalculateFullRebuild`, guardado y cierre correctos.
- Reapertura: solo lectura, `xlNormalLoad`, sin aviso de reparación y `RepairMode=False`.
- Procesos `EXCEL.EXE` tras guardado y reapertura: 0.
- Vista previa PDF: 7 páginas; Dashboard, catálogo, resultados custom, visibilidad y gráficas revisados visualmente.

## Benchmark independiente

No se abrió ni guardó con Excel. Se verificó únicamente por metadatos y hash:

- Tamaño: 255.393 bytes.
- SHA-256: `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.
- Estado: intacto.

## Advertencias pendientes

No quedan defectos bloqueantes en el libro. Se mantienen como limitaciones metodológicas, no como tareas de reparación:

- dos sources declaradas sin filas en la ventana final;
- ETW no concluyente;
- ausencia de P0 en `validated`;
- TrackNetwork sin atribución fiable de proceso;
- salida HTTP de TEC-009 acreditada por runner y no por el artifact detector.
