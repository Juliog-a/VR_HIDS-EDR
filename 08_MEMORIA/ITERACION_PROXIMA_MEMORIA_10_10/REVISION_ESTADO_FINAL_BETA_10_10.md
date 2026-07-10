# REVISION_ESTADO_FINAL_BETA_10_10

Fecha: 2026-07-09

## Nota de vigencia 2026-07-10

Este informe queda como histórico de la revisión pre-PDF.

La versión vigente única tras la fusión es:

`TFM_MEMORIA_BETA_v1.docx`

La copia `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx` fue archivada en:

`_ARCHIVO_OBSOLETO_NO_USAR/FUSION_20260710_1056`

Los bloqueantes de este informe relativos al índice general duplicado, temporal
Word y frase JSONL/Discord fueron tratados durante la fusión. Sigue pendiente
la revisión visual humana y la exportación de PDF provisional.

Documento revisado:

- Original intacto: `TFM_MEMORIA_BETA_v1.docx`
- Copia de revisión con correcciones menores: `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`

Hashes SHA-256:

- Original: `06EA717AA88071F2C785689A4AAFDDA03B5943A85A75F3D99DBB4F788B8A6106`
- Copia revisada: `BB1F1464EC49D9C79DB5AC4E9C88468560E454D20877AEBC982B188C90B13E70`

## Veredicto general

Estado global:

- APTO PARA REVISIÓN FINAL HUMANA

No está apto para PDF final todavía.

Porcentaje estimado de cierre: 88 %.

La memoria está cerca de versión final a nivel de contenido técnico y metodológico. Los resultados principales están alineados con el contexto del proyecto: 9/9 técnicas TEC-001 a TEC-009, 379 filas `CLIENT_EVENT`, FP 10/10 con hits=0, benchmark 9/9 runs válidos y Wazuh como comparación HIDS/SIEM no equivalente. No se detectan contradicciones graves sobre TEC-009, benchmark o Wazuh.

El bloqueo actual no es experimental, sino editorial/metodológico de cierre: índices, numeración de tablas, revisión visual de tablas/captions y una frase del capítulo VI que puede confundirse con que JSONL/Discord son resultado primario.

## Estado por capítulos

| Capítulo | Estado | Riesgo | Acción antes de PDF |
|---|---|---|---|
| I. Introducción | Muy avanzado | Bajo | Revisión final de estilo y alcance. Ya no sobrepromete MITRE completa. |
| II. Planificación | Avanzado | Medio | Revisar tiempos verbales para que no parezca planificación futura. |
| III. Fundamentos teóricos | Avanzado | Bajo/medio | Revisar densidad de citas y formato APA exigido por la Universidad. |
| IV. Entorno, metodología y selección | Avanzado | Medio | Revisión visual de tablas y confirmar que rutas/scripts no aparecen como IOCs. |
| V. Desarrollo experimental | Avanzado | Medio | Mantener clara la separación `CLIENT_EVENT` detecta / `SERVER_EVENT` enruta. |
| VI. Resultados de detección | Requiere intervención puntual | Alto | Reescribir la frase sobre JSONL/Discord como salida externa derivada, no resultado primario. |
| VII. Benchmark | Cerrado a nivel de contenido | Bajo | Verificar visualmente gráficas/tablas. No mezcla benchmark con calidad de detección. |
| VIII. Wazuh | Cerrado a nivel de contenido | Bajo/medio | Revisar tabla comparativa y mantener gap 110201. |
| IX. Trabajo futuro | Avanzado | Bajo | Títulos corregidos en la copia; revisar que no suene a deuda crítica. |
| X. Conclusiones | Avanzado | Bajo | Revisar correspondencia final objetivo/resultado/limitación. |
| Anexos A-L | Completos | Medio | Revisar paginación, citas desde cuerpo e inclusión en índice. |

## Riesgos metodológicos

| Prioridad | Hallazgo | Impacto | Acción |
|---|---|---|---|
| P1 | Capítulo VI: frase “Las salidas JSONL o Discord se usan como resultado final...” | Puede contradecir el criterio fijo de que JSONL/Discord son salida externa, no detección primaria. | Sustituir por: “Las salidas JSONL y Discord se conservan como salida externa trazable derivada de detecciones `CLIENT_EVENT`; no se usan como fuente primaria de detección.” |
| P2 | Índice general duplicado antes del índice de código | Bloquea acabado profesional y puede confundir estructura. | Eliminar o regenerar el segundo TOC general; conservar un único índice general. |
| P2 | Índice de tablas obsoleto | Se ve salto `Tabla 55 -> Tabla 63 -> Tabla 57` en el índice, aunque el cuerpo tiene secuencia 56-64. | Actualizar campos/listas en Word y verificar secuencia completa. |
| P2 | Anexos en índice como “Capítulo XI: Anexo A...” | No invalida, pero queda menos limpio académicamente. | Recomendada forma “Anexo A. ...” si la plantilla lo permite. |
| P2 | Tablas potencialmente anchas | OpenXML detecta cuatro tablas ligeramente por encima del ancho útil estimado. | Revisión visual en Word/PDF. |

No se detectan contradicciones P1 sobre:

- TEC-009.
- `CLIENT_EVENT`.
- `SERVER_EVENT`.
- Benchmark.
- Wazuh.
- Falsos positivos.
- Cobertura MITRE.
- Uso de malware real.
- Extrapolación a producción.

## Problemas de formato

- El documento OpenXML es válido.
- No hay campos marcados como `dirty`.
- No se detectaron referencias cruzadas rotas tipo `Error!`, `Marcador no definido` o `Reference source not found`.
- Hay dos campos `TOC \o "1-3"`, lo que explica el índice general duplicado.
- Existen índices de tablas, figuras y código, pero el índice de tablas está desactualizado.
- Se detectan 104 tablas, 71 imágenes, 76 dibujos y 31 captions de código.
- Tablas a revisar visualmente por ancho estimado: OpenXML table index 29, 50, 52 y 71.
- Tabla grande a revisar por posible partición: OpenXML table index 2, 39 filas.
- En la revisión del 2026-07-09 quedó un temporal Word `~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`. Tras la fusión del 2026-07-10, ese temporal fue archivado en `_ARCHIVO_OBSOLETO_NO_USAR/FUSION_20260710_1056`.

## Problemas de redacción

Corregidos en la copia:

- `campana` -> `campaña`.
- `validos` -> `válidos`.
- `metricas` -> `métricas`.
- `linea` -> `línea`.
- `discord` -> `Discord`.
- `Final del formulario` eliminado.
- `Tabla 47Casos` y `Tabla 51Perfiles` corregidas.
- Títulos del capítulo IX mejorados.

Pendiente:

- Reescritura metodológica de la frase JSONL/Discord del capítulo VI.
- Revisión humana de estilo en capítulos II, IV y V.

## Bibliografía

Estado: suficiente para revisión final humana.

- Existe bloque “Referencias bibliográficas”.
- Hay citas insertadas en el cuerpo.
- No se ha validado contra una variante concreta de APA exigida por la Universidad.
- Las fuentes internas se usan como trazabilidad de resultados, no como sustituto de bibliografía externa.

## Anexos

Estado: completos, pendientes de revisión visual.

- Detectados Anexos A-L dentro del Word.
- Aportan trazabilidad técnica.
- El índice los muestra como capítulos XI-XXII; se recomienda limpiar este formato si la plantilla lo permite.
- No se detecta ausencia de anexos A-L.
- Revisar que cada anexo esté citado desde el cuerpo y que no rompa paginación.

## Estado de TEC-009

Estado: metodológicamente correcto.

- Formulada como exfiltración HTTP controlada en laboratorio.
- Se conserva staging, archivado ZIP y transferencia HTTP controlada hacia host receptor.
- La ausencia de `receiver_log.jsonl` o ZIP recibido en el paquete revisado se trata como limitación documental secundaria.
- No se detectan formulaciones que rebajen TEC-009 a “solo local”, “no válida” o “solo staging”.

## Estado de benchmark

Estado: correcto.

- 9/9 runs válidos.
- 3 escenarios.
- 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- RAM media aproximada: 57 MB.
- CPU media en `VR_TEC_RUNNER`: 0,95 %.
- Pico CPU cliente VR: 54 %.
- No se mezcla benchmark con calidad de detección.

## Estado de Wazuh

Estado: correcto con revisión visual pendiente.

- Wazuh se mantiene como comparación HIDS/SIEM.
- No sustituye a Velociraptor.
- No hay equivalencia numérica directa entre `archives/alerts` y `CLIENT_EVENT`.
- Gap `110201=0` conservado para TEC-009.
- No se afirma superioridad absoluta.

## Decisión PDF

No exportar PDF final todavía.

Siguiente acción exacta:

1. Cerrar Word y eliminar el temporal `~$...` si desaparece tras cerrar.
2. Abrir `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
3. Eliminar/regenerar el índice general duplicado.
4. Actualizar índice general, tablas, figuras y código.
5. Reescribir la frase JSONL/Discord del capítulo VI.
6. Revisar visualmente tablas, captions, saltos de página y anexos.
7. Exportar PDF solo si la revisión visual queda limpia.
