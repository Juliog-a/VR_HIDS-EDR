# Cambios de los Excel definitivos - 12/07/2026

## Integración quirúrgica Wazuh base/custom - 13/07/2026

- Maestro utilizado: versión manual vigente de `08_MEMORIA/ENTREGABLE/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`; no se restauró ninguna versión anterior.
- Backup mínimo verificado antes de editar: `ENTREGABLE/BACKUP_ANTES_WAZUH_BASE_20260713`, 317.634 B, SHA-256 `31A4EFC9DC0A412B97C7ACDF95F8ADDA55370C134247BEB41C9F7A836508F7FE`, 31 hojas y 11 gráficos.
- Evidencia limitada a `10_WAZUH/TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz`: baseline base (11:10:29..11:22:06Z), campaña TEC (11:26:39..11:37:39Z), JSONL combinado, JSONL 110xxx y ruleset efectivo `tfm_wazuh_custom_rules_v2.xml`.
- Wazuh base: **3/9**, TEC-002, TEC-004 y TEC-006; rule_id principales 92032, 92302, 91835 y 92077.
- Wazuh custom: **4/9**, TEC-001, TEC-005, TEC-007 y TEC-008; rule_id principales 110301, 110402, 110203 y 110202.
- El 4/9 anterior correspondía exclusivamente a Wazuh custom. La campaña TEC contiene 54 alertas nativas y 56 alertas 110xxx; el volumen de alertas no se usa como número de técnicas.
- `01_Dashboard`: gráfico de detección específica con cuatro barras reales (9, 3, 3 y 4), gráfico de cobertura por técnica con cuatro series y nota de separación del ruleset base/custom.
- `GRAFICAS`: fuentes reales A:E y F:G, gráfico base/custom por técnica con dos series, escala 0/1 legible y heatmap de cinco capacidades con cuatro variantes, sin salida externa.
- `05_Matriz_Resultados`: tabla ampliada de forma conservadora a A4:Z13; columnas separadas de VR custom, públicos CH, Wazuh base y Wazuh custom, con estados y evidencia por TEC.
- `COMPARACION_VR_WAZUH`: tabla comparativa de cuatro variantes y desglose base/custom; la unión 7/9 se identifica solo como dato complementario.
- `WAZUH_DETALLE`: rule_id, MITRE, alertas, veredicto, ventana y archivo fuente diferenciados por grupo.
- Textos y trazabilidad actualizados en `RESUMEN_EJECUTIVO`, `README`, `FUENTES`, `DISCREPANCIAS` y `17_Incoherencias_Cerradas`.
- Cambios previos conservados: 0 `Pendiente`, Hayabusa CH 3/9, CHM no ejecutada, ETW no concluyente, TrackNetwork sin proceso fiable, Discord/HTTP y CLIENT_EVENT/SERVER_EVENT/Router separados, cinco campañas positivas y una no concluyente.
- Revisión visual: exportación PDF nativa y render PNG de dashboard, gráficos, comparativa, matriz, resumen y detalle Wazuh; sin solapes relevantes ni cocientes `/9` convertidos en fechas.
- OpenXML: válido; 0 XML no parseables, referencias rotas, series rotas, gráficos vacíos, errores de fórmula, referencias a hojas inexistentes o enlaces externos.
- Excel COM: cálculo automático, `CalculateFullRebuild`, guardado, reapertura de solo lectura con `RepairMode=False`, 0 aserciones fallidas y 0 procesos `EXCEL.EXE` restantes.
- Resultado final: 328.177 B; SHA-256 `FDA4501CB730058F2645F2A7FF5A80D66665AE933B45474833FFCB3ED2B86775`; 31 hojas; 11 gráficos; 56 fórmulas; 0 `Pendiente`.
- Copia auxiliar: mismo tamaño y SHA-256; identidad binaria confirmada.
- Benchmark: no abierto, no guardado y no modificado; 255.393 B; SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.
- Scripts de esta corrección: `Apply-WazuhBaseCustomSurgical.ps1`, `Repair-WazuhBaseCustomDisplay.ps1`, `Validate-WazuhBaseCustomFinal.ps1` y `Export-WazuhBaseCustomPdfQA.ps1`.

## Corrección final quirúrgica del libro de visibilidad - 13/07/2026

- Archivo modificado: `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx`.
- Motor: Microsoft Excel COM; no se reconstruyó el libro ni se repitió el análisis de evidencias.
- Backup mínimo previo: 316.711 B; SHA-256 `A4BD6E0752E74B7948022A958FD7A3BFB37ACD7F95B8B84F0FC087B7E008B2E0`.
- `01_Dashboard`: gráfico de capas reducido a Visibilidad, Detección y Alerta RT; salida externa retirada del rango fuente; valor custom Wazuh corregido a 4 desde celda fuente; notas metodológicas para salida externa y `/14` frente a `/9`.
- `GRAFICAS`: campañas públicas alineadas a 5 positivas y 1 no concluyente; heatmap sin salida externa; runner FP custom separado de Hayabusa Monitoring.
- `04_Control_Publicos`: columnas S, T, U y AE cerradas con vocabularios semánticos válidos.
- `05_Matriz_Resultados`: K7:K9 como no ejecutado en campaña final; CH histórico/no RT; CHM no ejecutada.
- `08_Benignas_FP`: OB-001 a OB-008 marcadas como sustituidas por la campaña FP final y no evaluadas individualmente.
- `15_Publicos_Definitivo`: Discord separado de salida HTTP; TrackNetwork limitado a visibilidad de conexión; HTTP 200 atribuido al runner independiente; SERVER_EVENT, Router e histórico separados de las seis campañas CLIENT_EVENT; formato de lectura e impresión ajustado.
- `09_Benchmark_Plan`, `12_Plan_Alertas` y `99_Listas`: cierre contextual de los valores exactos `Pendiente` restantes, sin introducir resultados experimentales nuevos.
- Textos dependientes normalizados en `RESUMEN_EJECUTIVO`, `11_Alertabilidad`, `13_Hayabusa_Resumen`, `17_Incoherencias_Cerradas`, `COMPARACION_VR_WAZUH` y `README`.
- Valores exactos `Pendiente`: 129 antes; 0 después.
- Resultado intermedio de esa corrección previa, antes de integrar Wazuh base/custom: 318.738 B; SHA-256 `2C8056A8DD9F74D246488D01FEFEC091A95DF1EB00CDE8173897CC5296509E45`; 31 hojas; 11 gráficos.
- OpenXML: válido; 0 XML no parseables, referencias rotas, series rotas, gráficos vacíos, errores de fórmula, referencias a hojas inexistentes o enlaces externos.
- Excel COM: cálculo automático, `CalculateFullRebuild`, guardado y reapertura en solo lectura sin reparación; 0 procesos `EXCEL.EXE` restantes.
- Copia auxiliar: mismo tamaño y SHA-256; identidad binaria confirmada.
- Benchmark: no abierto, no guardado y no modificado; 255.393 B; SHA-256 `C2255C2563AACFA7F8CA56496EAE934843C411D89011539B9D167F8A13BE93D0`.

## Ejecución

- Script de modificación: `Build-DefinitiveExcels.ps1`.
- Motor de edición y guardado: Microsoft Excel COM.
- Auditoría estructural: `audit_xlsx.py` sobre ZIP/OpenXML y lectura openpyxl sin guardado.
- Validación de reapertura: `Validate-DefinitiveExcels.ps1`.
- Auditoría cruzada: `final_cross_audit.py`, 39/39 comprobaciones.
- Revisión visual: exportación PDF nativa de Excel y renderizado PNG; gráficos del benchmark exportados también como objetos individuales.

## Libro de visibilidad, detección, FP y Wazuh

### Datos y trazabilidad

- Integrados 516 eventos/matches TEC canónicos y 127 filas benignas/FP de las seis campañas públicas.
- Incorporadas ventanas TEC/FP, conteos raw/canónicos, resultados, limitaciones y fuentes exactas.
- Añadida trazabilidad de 112 ficheros de evidencia con SHA-256 en `FUENTES`.
- ServiceCreation queda como creación 7045 observada/visibilidad, no como detección de actividad maliciosa.
- ETW queda `NO CONCLUYENTE`; Hayabusa CHM queda `NO EJECUTADA`.
- Hayabusa CH queda en 3/9 TEC específicas, sin campaña Medium inventada.
- TrackNetwork queda sin atribución fiable de proceso; el HTTP 200 de TEC-009 se acredita por el runner independiente.
- Custom P1-P4 y Wazuh se conservan; Wazuh no se separa en otro libro.

### Fórmulas, tablas y gráficos

- Fórmulas finales: 56; errores: 0.
- Tablas estructuradas: 10; rangos ampliados de forma conservadora.
- Gráficos: 11; seis en dashboard.
- Reparado el gráfico Wazuh con una sola serie de alertas y categorías A35:A41.
- Creados/reconstruidos cinco gráficos: visibilidad/detección por TEC, comparación de sistemas, capas, FP y estado de campañas.
- Corregidos formatos heredados que mostraban 516/127 como porcentajes y 3/9 como fecha.
- Mejorados anchos y alturas del resumen ejecutivo y del resumen Hayabusa.

## Benchmark

### Datos conservados

- Tres escenarios: `BASELINE_NO_VR`, `VR_IDLE` y `VR_TEC_RUNNER`.
- Tres repeticiones por escenario y 9/9 runs válidos.
- `SERVER_GUI` excluido; `notepad.exe=0`.
- No se alteraron las mediciones de CPU, RAM, duración, muestras ni totales.
- Las hojas fuente críticas no cambiaron frente al maestro.

### Fórmulas y gráficos

- `RESUMEN_EJECUTIVO!C3="9/9"` y `B5="3/3"` fijados como texto.
- Añadidas nueve fórmulas helper para el gráfico de repeticiones; errores: 0.
- Gráficos 1-5 con títulos, leyendas, unidades y ejes decimales legibles.
- Gráfico 6 sustituido por duración en segundos con REP_01, REP_02 y REP_03 por escenario.
- Los seis gráficos se redimensionaron y ordenaron en una cuadrícula homogénea dentro del área `A1:N112`.

## Validación y cierre

- Excel COM ejecutó cálculo automático, `CalculateFullRebuild`, guardado y reapertura en solo lectura.
- Reparaciones detectadas: 0.
- Errores de fórmula: 0.
- Errores XML: 0.
- Referencias de gráfico rotas: 0.
- Gráficos sin series: 0.
- Enlaces externos: 0.
- Copias auxiliares verificadas por tamaño y SHA-256.
- Las dos versiones antiguas se trasladaron al backup, sin borrado definitivo.
- Número final de Excel en la raíz de `ENTREGABLE`: 2.
