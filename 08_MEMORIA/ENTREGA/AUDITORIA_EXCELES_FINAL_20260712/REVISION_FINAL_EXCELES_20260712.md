# Revisión final de los Excel del TFM - 12/07/2026

## Estado

**APTO.** Los dos libros definitivos han sido modificados con Microsoft Excel COM, recalculados, guardados, reabiertos sin reparación y auditados mediante OpenXML, openpyxl, Excel COM y revisión visual. La auditoría cruzada final ha superado **39/39 comprobaciones**.

En la raíz de `08_MEMORIA/ENTREGABLE` existen exactamente dos Excel:

1. `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`
2. `TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx`

## Libros maestros y trazabilidad inicial

| Libro | Maestro seleccionado | Tamaño inicial | SHA-256 inicial | Motivo |
|---|---|---:|---|---|
| Visibilidad, detección, FP y Wazuh | `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx` | 280.969 B | `947161E4ECF609EF8CD97AD24BB604EFEDE04FD1251642ECFEAC1BF5AAEEFE09` | 31 hojas, custom P1-P4, Router, FP, Wazuh, comparativa, fuentes y dashboard consolidados. |
| Benchmark | `08_MEMORIA/ENTREGABLE/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx` | 253.219 B | `AD953B6F760DBFF580099B8DA899929CAED3383419C97E8E69C3E6D6F9E52CF4` | 11 hojas, 9/9 runs, tres escenarios, tres repeticiones y seis gráficos. |

Los maestros se seleccionaron antes de su traslado final. Sus copias con los mismos hashes están conservadas en `ENTREGABLE/BACKUP_ANTES_CIERRE_EXCELES_20260712`.

## Evidencias públicas procesadas

Raíz: `04_EVIDENCE/PRUEBAS_CON_ARTIFACTS_PUBLICOS`.

Se acreditaron las seis carpetas previstas y 112 ficheros. No existe una campaña Hayabusa CHM separada; se marca como **NO EJECUTADA** y no se añaden resultados Medium como campaña independiente.

| Artifact / campaña | Ventana TEC UTC | Ventana FP UTC | TEC canónicos | FP / ventana | Resultado y capa |
|---|---|---|---:|---:|---|
| Windows.Events.ProcessCreation | 11:50:43.356-11:58:00.217 | 12:01:47.092-12:02:15.729 | 89 | 18 | POSITIVO: visibilidad; no detección por proceso observado. |
| Windows.Events.ServiceCreation | 12:26:12.998-12:33:05.540 | 12:37:57.100-12:38:39.571 | 1 | 0 | POSITIVO: creación 7045 observada; no detección de malicia. Error 1053 esperado. |
| Windows.Sysinternals.SysmonLogForward | 12:46:16.670-12:52:34.859 | 12:55:41.452-12:56:21.064 | 113 | 24 | POSITIVO: visibilidad, reenvío y evidencia forense; IDs 1, 3, 11, 12 y 13. |
| Windows.Hayabusa.Monitoring CH | 13:25:03.318-13:31:18.641 | 13:41:00.378-13:41:29.197 | 185 matches | 73 filas | POSITIVO CON FP: detección Sigma CH en TEC-001, TEC-003 y TEC-005; 3/9. |
| Windows.ETW.Monitoring | 14:22:25.782-14:30:04.869 | 14:32:20.307-14:32:57.519 | 0 | 0 | **NO CONCLUYENTE**: faltan runner, log/summary TEC y manifiesto de configuración. |
| Generic.Events.TrackNetworkConnections | 14:43:12.205-14:50:12.514 | 14:52:39.396-14:53:04.206 | 128 | 12 | POSITIVO: conexión visible; tres filas destino 192.168.1.129:8088, PID=0 y ProcInfo vacío; sin atribución fiable. |

Totales acreditados: **516 eventos/matches TEC canónicos**, **521 filas raw dentro de ventana antes de deduplicación** y **127 filas en ventanas benignas/FP**. Estos volúmenes no se suman como detecciones.

Hayabusa CH: 185 coincidencias únicas de 187 filas; informational=92, low=38, medium=48, high=7 y critical=0. En FP: 73 filas, 4 high, 3 eventos subyacentes y un caso benigno afectado, FP-003 Scheduled Task.

## Tratamiento de ETW y TEC-009

- ETW se clasifica como **NO CONCLUYENTE**, no como negativo válido y sin afirmar que el artifact no funcione.
- TrackNetwork acredita una conexión a `192.168.1.129:8088` pero no una atribución de proceso fiable.
- El runner TEC-009 independiente acredita `ReceiverReachable=True`, `HTTPStatus=200`, `UploadSucceeded=True`, ZIP de 6.245 B y SHA-256 `247F688837E109355E49D3242F323A6FCB34C90C464DB290FEFBA7C68E6E87DC`, confirmado también por la respuesta del receptor.
- El éxito HTTP se mantiene separado del resultado del artifact TrackNetwork.

## Resultados custom y Wazuh conservados

- Velociraptor custom P1/P2/P3/P4: 379 filas consolidadas; P1=19, P2=118, P3=109 y P4=133.
- Wazuh base acredita **3/9** técnicas: TEC-002, TEC-004 y TEC-006; reglas nativas principales 92032, 92302, 91835 y 92077.
- Wazuh custom acredita **4/9** técnicas: TEC-001, TEC-005, TEC-007 y TEC-008; reglas TFM principales 110301, 110402, 110203 y 110202.
- Las 110 alertas de la campaña TEC se desglosan en 54 alertas de reglas nativas y 56 alertas 110xxx; son trazabilidad y no el numerador de cobertura. La unión base+custom alcanza 7/9 solo como dato complementario, no como quinta solución.
- La ventana base separada contiene 58 alertas pero no marcadores TEC; se conserva como baseline de ruido/visibilidad. La clasificación técnica se realiza con alertas nativas y custom dentro de la ventana TEC trazable.

## Cambios por hoja - libro de visibilidad

- `01_Dashboard`: el gráfico de capas queda limitado a Visibilidad, Detección y Alerta RT (9/9/9 custom; 9/3/3 públicos); la detección específica muestra cuatro barras trazables: VR custom 9/9, públicos CH 3/9, Wazuh base 3/9 y Wazuh custom 4/9; `/14` queda separado de las 9 TEC.
- `README`, `00_Guia`, `06_Resumen` y `RESUMEN_EJECUTIVO`: estado final, limitaciones, totales y conclusiones cuantitativas; Wazuh base/custom queda separado y la unidad de comparación es la técnica TEC única.
- `04_Control_Publicos`: `FP_Risk`, `Custom_Necesario`, `Estado` y `Resultado_Campaña` normalizados con valores semánticos; no quedan valores `Pendiente`.
- `15_Publicos_Definitivo`: seis campañas CLIENT_EVENT separadas de componentes históricos, SERVER_EVENT, Router, transporte y salida externa; Discord y HTTP quedan en columnas distintas; total 5 positivas y 1 no concluyente.
- `05_Matriz_Resultados`: CH histórico/no RT; CHM no ejecutada y excluida de gráficos, totales y cobertura; columnas separadas de VR custom, públicos CH, Wazuh base y Wazuh custom, con estados `DETECTADA` o `VISIBLE SIN DETECCIÓN` y evidencia por TEC.
- `08_Benignas_FP`: OB-001 a OB-008 quedan identificadas como sustituidas por la campaña FP final y no evaluadas individualmente; los resultados PUB-FP consolidados se conservan.
- `13_Hayabusa_Resumen`: FP públicos, niveles, detección CH 3/9 y un caso FP confirmado, sin mezclarlo con el runner FP custom.
- `COMPARACION_VR_WAZUH`: tabla real de cuatro variantes, desglose base/custom, reglas, alertas fuente, ventanas y limitaciones; 7/9 queda identificado como unión complementaria.
- `WAZUH_DETALLE`: reglas nativas y 110xxx separadas, veredicto por técnica, conteos de alertas y fuente/ventana.
- `DISCREPANCIAS` y `17_Incoherencias_Cerradas`: ETW, deduplicación, TrackNetwork, separación del HTTP 200 y criterio conservador de clasificación Wazuh.
- `FUENTES`: trazabilidad previa de evidencia pública y diez entradas concretas de `10_WAZUH`, con ruta interna, tamaño, timestamp, SHA-256, rol y grupo base/custom.
- `GRAFICAS`: estados públicos 5 positivos/1 no concluyente conservados; detección por técnica y heatmap con VR custom, públicos CH, Wazuh base y Wazuh custom separados; salida externa excluida del bloque HIDS.

## Ajuste final quirúrgico - 13/07/2026

- Se creó exclusivamente el backup mínimo del libro de visibilidad en `ENTREGABLE/BACKUP_ANTES_AJUSTE_FINAL_DASHBOARD_20260712`, con 316.711 B y SHA-256 `A4BD6E0752E74B7948022A958FD7A3BFB37ACD7F95B8B84F0FC087B7E008B2E0`.
- La edición se realizó directamente con Microsoft Excel COM y afectó solo a celdas, rangos fuente, series, títulos, notas y formato puntual. Se conservaron 31 hojas y 11 gráficos.
- Los valores exactos `Pendiente` pasaron de 129 a 0 mediante normalización por columna y contexto, sin sustitución masiva genérica.
- El libro de benchmark no se abrió ni se guardó; conserva 255.393 B y SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.
- Auditor OpenXML final: ZIP válido, 0 XML no parseables, 0 referencias rotas, 0 series rotas, 0 gráficos vacíos, 0 errores de fórmula, 0 referencias a hojas inexistentes y 0 enlaces externos.
- Excel COM final: cálculo automático, `CalculateFullRebuild`, guardado, 0 errores, reapertura de solo lectura con `RepairMode=False` y 0 procesos `EXCEL.EXE` restantes.
- La copia auxiliar de visibilidad se sustituyó por una copia binaria idéntica al entregable.

## Integración final de Wazuh base y Wazuh custom - 13/07/2026

- Se utilizó exclusivamente el maestro manual vigente de `ENTREGABLE`. Antes de editar se verificó el backup mínimo `ENTREGABLE/BACKUP_ANTES_WAZUH_BASE_20260713`: 317.634 B, SHA-256 `31A4EFC9DC0A412B97C7ACDF95F8ADDA55370C134247BEB41C9F7A836508F7FE`, 31 hojas y 11 gráficos.
- La inspección se limitó a `10_WAZUH/TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz`, sus ventanas y JSONL de alertas, el ruleset efectivo `tfm_wazuh_custom_rules_v2.xml` y la memoria metodológica `WAZUH.docx` como contraste narrativo.
- Resultado final: Wazuh base 3/9 (TEC-002/004/006) y Wazuh custom 4/9 (TEC-001/005/007/008). El valor anterior 4/9 correspondía exclusivamente a reglas custom; no era un total combinado.
- Reglas base principales: 92032, 92302, 91835 y 92077. Reglas custom principales: 110301, 110402, 110203 y 110202. No se contabilizan 92307/T1543.003 como TEC-005/T1569.002, las alertas genéricas de `schtasks`, 110501 ni reglas con cero alertas.
- `01_Dashboard`, `GRAFICAS`, `COMPARACION_VR_WAZUH`, `05_Matriz_Resultados`, `WAZUH_DETALLE`, `RESUMEN_EJECUTIVO`, `README`, `FUENTES`, `DISCREPANCIAS` y `17_Incoherencias_Cerradas` se actualizaron mediante Excel COM sin reconstruir el libro ni crear gráficos redundantes.
- Revisión visual focalizada mediante exportación PDF nativa de Excel y render PNG: cuatro barras visibles, dos series Wazuh diferenciadas, escala binaria 0/1 legible, heatmap sin salida externa y cocientes `/9` sin conversión a fechas.
- Valores exactos `Pendiente`: 0 antes y 0 después.
- Auditor OpenXML `AUDITORIA_WAZUH_BASE_CUSTOM_FINAL_20260713.json`: ZIP/OpenXML válido; 0 XML no parseables, referencias rotas, series rotas, gráficos vacíos, errores de fórmula, referencias a hojas inexistentes y enlaces externos.
- Validación COM `EXCEL_COM_VALIDACION_WAZUH_BASE_CUSTOM_20260713.json`: cálculo automático, `CalculateFullRebuild`, guardado, 31 hojas, 11 gráficos, 0 errores, reapertura de solo lectura, `RepairMode=False`, 0 aserciones fallidas y 0 procesos `EXCEL.EXE` restantes.
- Libro final de visibilidad: 328.177 B; SHA-256 `FDA4501CB730058F2645F2A7FF5A80D66665AE933B45474833FFCB3ED2B86775`.
- Copia auxiliar: 328.177 B y el mismo SHA-256; identidad binaria confirmada.
- Benchmark: no abierto, no guardado y no modificado; 255.393 B; SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.

## Cambios por hoja - benchmark

- `README`: fecha y estado final.
- `RESUMEN_EJECUTIVO`: `C3="9/9"` y `B5="3/3"` como texto.
- `CALIDAD_DATOS`: comprobaciones de unidades y población de gráficos.
- `DISCREPANCIAS`: registro del formato de KPI y sustitución del gráfico mixto.
- `GRAFICAS`: seis gráficos redimensionados y ordenados; el sexto representa duración por REP_01/02/03 y escenario. No mezcla CPU, RAM y tiempo.

Se conservan exactamente tres escenarios, tres repeticiones por escenario, 9/9 ejecuciones válidas, `SERVER_GUI` excluido y `notepad.exe=0`. Las hojas fuente `RUNS_VALIDOS`, `RESUMEN_ESCENARIO`, `RAW_SAMPLES`, `PROCESS_SAMPLES`, `TRAZABILIDAD` y `LIMITACIONES` no presentan cambios de valores frente al maestro.

## Gráficos

Libro de visibilidad: **11 gráficos** en total, seis en el dashboard. La detección específica muestra VR custom, públicos CH, Wazuh base y Wazuh custom como variantes separadas; la cobertura por TEC usa cuatro series y la salida externa se evalúa aparte.

Benchmark: **6 gráficos**; CPU media/máxima, RAM media/máxima, duración media y duración por repetición/escenario. Todos tienen título, leyenda, unidades y series no vacías.

## Validaciones

| Control | Visibilidad | Benchmark |
|---|---:|---:|
| Hojas | 31 | 11 |
| Tablas estructuradas | 10 | 0 |
| Gráficos | 11 | 6 |
| Fórmulas | 56 | 9 |
| Errores de fórmula almacenados | 0 | 0 |
| Referencias a hojas inexistentes | 0 | 0 |
| Referencias de gráfico rotas | 0 | 0 |
| Gráficos sin series | 0 | 0 |
| Enlaces externos | 0 | 0 |
| XML no parseable | 0 | 0 |
| Excel COM | Válido | Válido |
| Reapertura solo lectura | Sin reparación | Sin reparación |

El aviso de openpyxl sobre extensiones de validación se produjo en lectura; los libros no fueron guardados con openpyxl y la extensión se conservó mediante Excel COM.

## Archivos finales

| Archivo | Tamaño | SHA-256 | Estado |
|---|---:|---|---|
| `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx` | 328.177 B | `FDA4501CB730058F2645F2A7FF5A80D66665AE933B45474833FFCB3ED2B86775` | APTO |
| `TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx` | 255.393 B | `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0` | APTO |

Las copias auxiliares tienen el mismo tamaño y SHA-256 que sus originales en `ENTREGABLE`.

## Observaciones reales no bloqueantes

1. ETW permanece no concluyente por falta de runner/log/summary TEC y manifiesto de configuración.
2. La configuración efectiva completa de Hayabusa no quedó preservada; el veredicto CH utiliza exclusivamente high/critical realmente observados.
3. TrackNetwork no permite atribuir las filas objetivo a PowerShell por `Timestamp=1601`, `PID=0` y `ProcInfo` vacío.
4. El benchmark conserva el WARN no bloqueante `RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER`.
5. Wazuh mantiene gaps específicos en TEC-003 y TEC-009; 92307 es una alerta relacionada con creación de servicio pero no mapea la técnica evaluada T1569.002.
6. La regla custom 110202 acredita TEC-008, pero incluye coincidencias fuera de esa ventana; el ruido queda documentado y no incrementa el número de técnicas.
