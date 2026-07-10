# Plan Pendiente Por Capítulos

Fecha: 2026-07-08

Nota: en este documento `P1`, `P2`, `P3` y `P4` son prioridades de trabajo. No se refieren a perfiles de artifacts salvo que se indique expresamente.

| Capítulo | Estado | Problemas | Pendiente | Prioridad | Evidencia que falta | Capturas que faltan | Texto a cerrar | Riesgo metodológico | Acción siguiente |
|---|---|---|---|---|---|---|---|---|---|
| I. Introducción | Revisado, no cerrado formalmente | Objetivos aún pueden parecer amplios respecto al subconjunto TEC-001 a TEC-009. | Alinear objetivo general, contribuciones y alcance real. | P2 | Ninguna nueva. | No obligatorias. | Redactar contribución del TFM en 5-7 líneas. | Sobreprometer cobertura MITRE completa. | Revisar objetivo general y limitar alcance a laboratorio. |
| II. Planificación | Revisado, requiere lectura final | Puede conservar menciones históricas de planificación resuelta. | Revisar tiempos verbales y coherencia con decisión Wazuh frente a FortiEDR. | P3 | Ninguna. | No obligatorias. | Ajustar redacción de planificación ya ejecutada. | Bajo; es capítulo de contexto. | Lectura visual y corrección de estilo. |
| III. Fundamentos teóricos | Avanzado con citas APA insertadas | Requiere lectura visual y posible ajuste de estilo según norma específica de la Universidad. | Revisar coherencia de citas y evitar exceso de densidad bibliográfica. | P2 | Ninguna nueva para beta. | No obligatorias. | Reforzar diferencia visibilidad/detección/alerta/evidencia si el tutor lo pide. | Que el tribunal vea herramientas sin marco conceptual. | Lectura académica final del capítulo. |
| IV. Entorno, metodología y selección de técnicas | Corregido en beta | Quedan rutas y scripts como documentación de laboratorio; deben no parecer IOCs. | Revisar captions, arquitectura final y tabla de campañas canónicas. | P1 | Hashes/fuentes ya enlazados en anexos; falta revisión visual. | Arquitectura laboratorio, fuentes Sysmon/PowerShell, receptor HTTP si se quiere. | Mantener criterio TEC-009 ya corregido. | Confundir ejecución, visibilidad, detección y salida externa. | Revisar secciones 4.3-4.6 y referencias a TEC-009. |
| V. Desarrollo experimental | Corregido parcialmente | Puede quedar exceso de detalle operativo frente a detection engineering. | Consolidar tabla final de artifacts custom y rol del router SOC_v3. | P1 | YAML validados y hashes. | Velociraptor Client Monitoring/Client Events por perfiles. | Explicar P4/P3/P2/P1 como perfiles operativos, no prioridades. | Confundir SERVER_EVENT con detector o JSONL/Discord con detección. | Añadir tabla `CLIENT_EVENT detecta / SERVER_EVENT enruta`. |
| VI. Resultados de detección | Cerrado a nivel de contenido beta | Requiere revisión visual de tablas/figuras. | Confirmar que todas las cifras citan Excel/CSV final. | P1 | Ninguna para la beta; receptor final TEC-009 solo si se quiere eliminar limitación documental secundaria. | Client Events filtrado por TEC o perfil si se desea. | Revisar que TEC-009 figure como exfiltración HTTP controlada. | Sobreafirmar más allá de 9/9 en subconjunto experimental. | Abrir Word y revisar tablas de resultados. |
| VII. Benchmark y rendimiento | Cerrado a nivel de contenido beta | Requiere revisar que no se mezcle rendimiento con detección. | Confirmar gráficas y tabla de 9/9 runs. | P1 | Ninguna; Excel benchmark final APTO. | Gráficas CPU/RAM del Excel si se quieren incrustar mejor. | Mantener consumo medible pero bajo. | Interpretar benchmark como calidad de detección. | Revisar capítulo con Excel benchmark final abierto. |
| VIII. Wazuh | Cerrado a nivel de contenido beta | La regla 110201 queda como gap. | Revisar que la comparación no use volúmenes como equivalentes directos. | P1 | Ninguna para beta; revalidar 110201 solo si se decide corregir Wazuh. | Dashboard Wazuh, rule.id 110201=0, resumen base/custom. | Cerrar síntesis comparativa. | Afirmar superioridad absoluta o equivalencia directa Wazuh/Velociraptor. | Revisar tabla comparativa final. |
| IX. Trabajo futuro | Casi cerrado | Debe separar mejoras reales de tareas opcionales. | Priorizar ampliación MITRE, automation, SIEM/SOC, despliegue real y laboratorio. | P2 | Ninguna. | No obligatorias. | Quitar tareas ya cerradas. | Convertir limitaciones en promesas no soportadas. | Lectura final y limpieza de duplicidades. |
| X. Conclusiones | Cerrado a nivel de contenido beta | Pendiente homogeneizar con objetivos iniciales. | Revisar matriz objetivo/resultado/evidencia/limitación. | P1 | Ninguna. | No obligatorias. | Ajustar formulación de contribución real. | Extrapolar a toda MITRE o producción. | Revisar contra capítulos VI-VIII. |

## Tareas P1 Restantes

| Tarea | Motivo | Acción |
|---|---|---|
| Revisar visualmente `TFM_MEMORIA_BETA_v1.docx` en Word | Word abrió y guardó la beta, pero falta revisión humana de paginación y tablas. | Revisar tablas, anexos, captions y saltos de página. |
| Revisar capítulos IV-V | Soportan metodología y artifacts; mayor riesgo conceptual. | Confirmar separación `CLIENT_EVENT`/`SERVER_EVENT`/JSONL/Discord. |
| Revisar capítulo VI | Es el núcleo de resultados. | Confirmar cifras, captions y TEC-009. |
| Revisar capítulo VII | Benchmark no debe mezclarse con detección. | Mantener coste operativo, no calidad de detección. |
| Revisar capítulo VIII | Comparación Wazuh requiere precisión. | Mantener Wazuh como contraste HIDS/SIEM y gap 110201. |
| Revisar capítulo X | Cierre académico. | Responder objetivos sin sobrealcance. |

## Actualización 2026-07-08 12:50

- Referencias APA insertadas en la memoria beta.
- Citas añadidas en capítulos III, IV, V, VII, VIII y X.
- Anexos A-L desarrollados dentro de `TFM_MEMORIA_BETA_v1.docx`.
- Índice/listas actualizados mediante Word en modo invisible.
- SHA-256 beta tras la actualización: `DFB82ECF760E56C9F8E8D78485CF7D9571DCB6C6254D6281C9BF3BB95D9025C3`.

Pendiente P1 tras esta actualización: revisión visual humana de tablas, anexos, captions, saltos de página y coherencia final antes de exportar PDF.
