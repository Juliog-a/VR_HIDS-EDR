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
- Wazuh permanece integrado en el primer libro: 58 alertas base, 110 alertas custom, 5 reglas 110xxx activadas y un gap asociado a TEC-009.
- Las hojas críticas de datos previos, incluida `WAZUH_DETALLE`, permanecen sin cambios de valores frente al maestro.

## Cambios por hoja - libro de visibilidad

- `01_Dashboard`: KPI públicos, capas metodológicas, 516/127, CH 3/9, custom 9/9, Wazuh 4/9 y TEC-009 HTTP 200; seis gráficos visibles.
- `README`, `00_Guia`, `06_Resumen` y `RESUMEN_EJECUTIVO`: estado final, limitaciones, totales y conclusiones cuantitativas.
- `04_Control_Publicos` y `15_Publicos_Definitivo`: cierre de seis campañas, fuentes, ventanas, conteos, resultado y limitación.
- `05_Matriz_Resultados` y `03_Evaluacion_VR`: visibilidad separada de detección CH; CHM no ejecutada.
- `08_Benignas_FP` y `13_Hayabusa_Resumen`: FP públicos, niveles, 3/9 y un caso FP confirmado.
- `COMPARACION_VR_WAZUH`: fila comparativa de artifacts públicos sin mezclar volúmenes heterogéneos.
- `DISCREPANCIAS` y `17_Incoherencias_Cerradas`: ETW, deduplicación, TrackNetwork y separación del HTTP 200.
- `FUENTES`: trazabilidad individual de los 112 ficheros con ruta, tamaño, timestamp, SHA-256, rol y campaña.
- `GRAFICAS`: rangos helper, reparación Wazuh y fuentes de los nuevos gráficos.

## Cambios por hoja - benchmark

- `README`: fecha y estado final.
- `RESUMEN_EJECUTIVO`: `C3="9/9"` y `B5="3/3"` como texto.
- `CALIDAD_DATOS`: comprobaciones de unidades y población de gráficos.
- `DISCREPANCIAS`: registro del formato de KPI y sustitución del gráfico mixto.
- `GRAFICAS`: seis gráficos redimensionados y ordenados; el sexto representa duración por REP_01/02/03 y escenario. No mezcla CPU, RAM y tiempo.

Se conservan exactamente tres escenarios, tres repeticiones por escenario, 9/9 ejecuciones válidas, `SERVER_GUI` excluido y `notepad.exe=0`. Las hojas fuente `RUNS_VALIDOS`, `RESUMEN_ESCENARIO`, `RAW_SAMPLES`, `PROCESS_SAMPLES`, `TRAZABILIDAD` y `LIMITACIONES` no presentan cambios de valores frente al maestro.

## Gráficos

Libro de visibilidad: **11 gráficos** en total, seis en el dashboard. Se incluyen visibilidad/detección por TEC, custom frente a públicos y Wazuh, capas, FP y estados de campaña. El gráfico Wazuh usa una serie, categorías `A35:A41` y valores `B35:B41`.

Benchmark: **6 gráficos**; CPU media/máxima, RAM media/máxima, duración media y duración por repetición/escenario. Todos tienen título, leyenda, unidades y series no vacías.

## Validaciones

| Control | Visibilidad | Benchmark |
|---|---:|---:|
| Hojas | 31 | 11 |
| Tablas estructuradas | 10 | 0 |
| Gráficos | 11 | 6 |
| Fórmulas | 49 | 9 |
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
| `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx` | 316.711 B | `A4BD6E0752E74B7948022A958FD7A3BFB37ACD7F95B8B84F0FC087B7E008B2E0` | APTO |
| `TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx` | 255.393 B | `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0` | APTO |

Las copias auxiliares tienen el mismo tamaño y SHA-256 que sus originales en `ENTREGABLE`.

## Observaciones reales no bloqueantes

1. ETW permanece no concluyente por falta de runner/log/summary TEC y manifiesto de configuración.
2. La configuración efectiva completa de Hayabusa no quedó preservada; el veredicto CH utiliza exclusivamente high/critical realmente observados.
3. TrackNetwork no permite atribuir las filas objetivo a PowerShell por `Timestamp=1601`, `PID=0` y `ProcInfo` vacío.
4. El benchmark conserva el WARN no bloqueante `RunnerStillRunningAtEnd=True` en `VR_TEC_RUNNER`.

