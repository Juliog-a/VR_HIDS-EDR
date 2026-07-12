# Cambios de los Excel definitivos - 12/07/2026

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

- Fórmulas finales: 49; errores: 0.
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

