# Auditoria TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026

Fecha generacion: 2026-06-26 11:40:12 +02:00

## Entradas canonicas

- Excel base: 04_EVIDENCE\Analisis_Tecnicas_TFM_V.4.xlsx
- VR REAL: 05_LOGS\Validas\Prueba_24_06_2026
- VR FP: 05_LOGS\Validas\Prueba_24_06_2026_FP
- Wazuh: 10_WAZUH\TFM_WAZUH_VISIBILIDAD_EVIDENCE

## Criterio de clasificacion

- REAL: filas CLIENT_EVENT de carpeta real dentro de 2026-06-24T12:50:22.335Z..2026-06-24T12:55:53.996Z.
- FP: summary/vr_hits del RunId TFM_FP_20260626_101948 dentro de 2026-06-26T08:19:48.709Z..2026-06-26T08:20:18.059Z.
- Filas fuera de ventana o duplicadas en carpeta FP quedan en DISCREPANCIAS.
- JSONL/Discord no se consideran deteccion.

## Resultados globales

- TEC reales: 9/9 presentes; OK=9.
- Alertas VR reales incluidas: 379.
- Alertas VR por perfil: P1=19; P2=118; P3=109; P4=133.
- FP: 10/10 presentes; hits FP=0.
- Wazuh base: archives=902, alerts=58.
- Wazuh custom: archives=18285, alerts=110, reglas 110xxx=56.
- 110201=0: gap TEC-009 conservado.

## Advertencias

- Discrepancias registradas: 48.
- Filas CLIENT_EVENT en ventana FP no reconciliadas con vr_hits: 82.
- TEC-009 tiene HTTP 200, UploadSucceeded=True, SHA-256 y bytes en ReceiverResponse, pero no se localiza receiver_log.jsonl ni ZIP recibido en las rutas revisadas.
- Se genero el XLSX con Excel COM local porque @oai/artifact-tool no estaba disponible en el entorno.

## Verificacion posterior

- Apertura Excel COM en modo solo lectura: OK.
- Hojas totales: 31.
- Hojas obligatorias presentes: README, RESUMEN_EJECUTIVO, TECNICAS_REAL_VR, VISIBILIDAD_SISTEMA, ALERTAS_VR, FP_RUNNER, FP_HITS, COMPARACION_VR_WAZUH, WAZUH_DETALLE, GRAFICAS, DISCREPANCIAS, FUENTES.
- Graficas nativas en GRAFICAS: 5.
- Heatmaps en GRAFICAS: 2 matrices coloreadas.
- Errores de formula buscados (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A): 0.
- Preview exportado: 04_EVIDENCE\Excel_visibilidad_26062026\GRAFICAS_preview.pdf.
- SHA-256 comun de las cuatro copias XLSX: F9CD06E7E986102028C7FE28523BEAD774ED8E724FD7464382CF2C91A45293C1.

## Salidas

- C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_visibilidad_FP\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx
- C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_visibilidad_FP\Analisis_Tecnicas_TFM_V.5.xlsx
- C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_visibilidad_FP\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_DATA_QUALITY.json
- C:\Users\julio\Desktop\TFM\08_MEMORIA\Excel_visibilidad_FP\TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_AUDIT.md
