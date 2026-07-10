# CONTEXT_ESTADO_VISIBILIDAD_FP_26062026

## Excel generado

- `08_MEMORIA/Excel_visibilidad_FP/TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx`
- `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_V.5.xlsx`
- Copias trazables en `04_EVIDENCE/Excel_visibilidad_26062026`.
- SHA-256 comun de las cuatro copias XLSX:
  `F9CD06E7E986102028C7FE28523BEAD774ED8E724FD7464382CF2C91A45293C1`.

## Salidas auxiliares

- Audit:
  `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_AUDIT.md`.
- Data quality:
  `TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_DATA_QUALITY.json`.
- Normalizados:
  `04_EVIDENCE/Excel_visibilidad_26062026/normalized`.
- Preview:
  `04_EVIDENCE/Excel_visibilidad_26062026/GRAFICAS_preview.pdf`.

## Campanas usadas

- VR REAL:
  `05_LOGS/Validas/Prueba_24_06_2026`.
- VR FP:
  `05_LOGS/Validas/Prueba_24_06_2026_FP`.
- Wazuh:
  `10_WAZUH/TFM_WAZUH_VISIBILIDAD_EVIDENCE/ENTREGA_WAZUH_VISIBILIDAD_20260622_114027`.

## Criterios de clasificacion

- REAL: filas `CLIENT_EVENT` de carpeta real dentro de
  `2026-06-24T12:50:22.335Z..2026-06-24T12:55:53.996Z`.
- FP: `summary.json` y `vr_hits.csv` del RunId `TFM_FP_20260626_101948`
  dentro de `2026-06-26T08:19:48.709Z..2026-06-26T08:20:18.059Z`.
- Filas fuera de ventana o duplicadas en carpeta FP se registran en
  `DISCREPANCIAS` y no se agregan como resultado definitivo.
- JSONL/Discord no se usan como deteccion.
- `CLIENT_EVENT` se mantiene como fuente de deteccion Velociraptor.
- `SERVER_EVENT`/JSONL se tratan como enrutamiento/salida externa.

## Resultados globales VR real

- TEC reales presentes: 9/9.
- Estado runner: 9 OK, 0 WARN, 0 FAIL.
- Alertas/detecciones VR reales incluidas: 379 filas `CLIENT_EVENT`.
- Distribucion por perfil:
  - P1: 19.
  - P2: 118.
  - P3: 109.
  - P4: 133.
- Todas las TEC-001 a TEC-009 tienen visibilidad Velociraptor en la ventana
  canonica.

## Resultados globales FP

- FP presentes: 10/10.
- Estado runner: 10 OK, 0 WARN, 0 FAIL, 0 SKIPPED.
- `VelociraptorHits.Total`: 0.
- `vr_hits.csv`: sin filas de datos.
- Se detectan 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas con
  `vr_hits`; quedan documentadas para revision, no como FP definitivos.

## Resultados globales Wazuh

- Wazuh base:
  - 902 archives.
  - 58 alerts.
  - 34 PowerShell 4104.
  - 457 Sysmon ID 1.
  - 413 Sysmon ID 11.
  - 0 System 7045.
  - 0 reglas custom 110xxx.
- Wazuh custom:
  - 18.285 archives.
  - 110 alerts.
  - 99 PowerShell 4104.
  - 127 Sysmon ID 1.
  - 24 Sysmon ID 11.
  - 1 System 7045.
  - 56 alertas TFM 110xxx.

## Reglas Wazuh 110xxx

- 110201 = 0. Gap TEC-009.
- 110202 = 5.
- 110203 = 4.
- 110301 = 2.
- 110401 = 0.
- 110402 = 1.
- 110501 = 44.

## Gaps y limitaciones

- `receiver_log.jsonl` y ZIP recibido de TEC-009 no se localizaron en las rutas
  revisadas.
- TEC-009 tiene HTTP 200, `UploadSucceeded=True`, SHA-256 y bytes en
  `ReceiverResponse`; se documenta como upload HTTP controlado con limitacion
  de receptor, no como exfiltracion contextual completa sin matiz.
- 48 discrepancias de fechas/ventanas quedan en hoja `DISCREPANCIAS`.
- Se uso Excel COM local para construir el XLSX porque `@oai/artifact-tool` no
  estaba disponible en el entorno.

## Decisiones metodologicas

- No inflar FP: `FP_HITS` se basa en `vr_hits.csv` y `summary.json`.
- No mezclar real y FP: los CSV mixtos se auditan, no se agregan fuera de
  ventana.
- P2 forense se conserva separado de P1/P3/P4 por perfil.
- Sysmon ID 26 se mantiene como evidencia forense, no alerta individual.
- Wazuh 110201 se conserva explicitamente como gap.

## Verificacion del Excel

- Apertura Excel COM en modo solo lectura: OK.
- Hojas totales: 31.
- Hojas obligatorias presentes: 12/12.
- Graficas nativas: 5.
- Heatmaps: 2 matrices coloreadas.
- Errores de formula buscados: 0.

## Proximos pasos Capitulo VI

- Usar `RESUMEN_EJECUTIVO`, `TECNICAS_REAL_VR`, `VISIBILIDAD_SISTEMA` y
  `ALERTAS_VR` para resultados Velociraptor.
- Explicar la separacion entre visibilidad, deteccion `CLIENT_EVENT`, alerta
  casi real y salida externa.
- Tratar TEC-009 con la limitacion de receptor documentada.

## Proximos pasos Capitulo VII

- Usar `COMPARACION_VR_WAZUH` y `WAZUH_DETALLE` para contraste VR/Wazuh.
- Presentar 110201=0 como gap de Wazuh custom para TEC-009.
- Incluir las discrepancias como control de calidad y no como debilidad oculta.
