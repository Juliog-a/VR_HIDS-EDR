# Contexto Memoria Beta v1

Fecha: 2026-07-08

## Estado Real del TFM

- Memoria en beta de trabajo.
- Documento base real más actualizado: `08_MEMORIA/TFM.docx`.
- Nueva beta generada: `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1.docx`.
- SHA-256 beta: `DD5E187FD6CA1D32FF9B314E381EAC347AF26B44D43FCAD93A97B53EF80BAEC9`.
- No se han ejecutado nuevas pruebas experimentales.
- Codex ha realizado revisión documental, edición técnica y validación estructural OpenXML, no validación experimental.

## Documentos Válidos

- `TFM_MEMORIA_BETA_v1.docx`: documento principal de revisión.
- `PLAN_PENDIENTE_POR_CAPITULOS.md`: plan vivo.
- `CAMBIOS_RECOMENDADOS_BETA_v1.md`: changelog técnico y riesgos.
- `CONTEXTO_MEMORIA_BETA_v1.md`: contexto de memoria.
- `README_ITERACION_BETA_v1.md`: guía de uso.
- `00_DOCUMENTOS_ARCHIVADOS_NO_USAR.md`: índice de archivado.

## Documentos Obsoletos

Los documentos auxiliares antiguos y la beta previa se han movido a:

`08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/_ARCHIVO_OBSOLETO_NO_USAR`

No deben usarse para continuar la memoria salvo recuperación puntual de texto histórico.

## Capítulos Cerrados

- Capítulo VI: resultados de detección, FP y calidad de datos.
- Capítulo VII: benchmark y evaluación de rendimiento.
- Capítulo VIII: validación comparativa con Wazuh.
- Capítulo X: conclusiones principales.

Cerrados significa cerrados a nivel de contenido beta, no cierre editorial final.

## Capítulos Pendientes

- Capítulo I: ajustar objetivos y contribución final.
- Capítulo II: revisión de coherencia temporal y planificación ya resuelta.
- Capítulo III: reforzar bibliografía y marco HIDS/DFIR/SIEM.
- Capítulo IV: revisar metodología, captions y anexos de evidencias.
- Capítulo V: reforzar arquitectura final de artifacts y router SOC_v3.
- Capítulo IX: limpiar trabajo futuro y priorizar.
- Bibliografía, anexos, índice, listas de figuras/tablas y revisión visual.

## Criterio Actualizado Sobre TEC-009

TEC-009 debe redactarse siempre en positivo como exfiltración HTTP controlada documentada. Debe evitarse cualquier formulación que la rebaje a mera preparación local o invalide el envío HTTP acreditado por el summary.

Redacción vigente:

- `exfiltración HTTP controlada en laboratorio`.
- `staging, archivado ZIP y exfiltración HTTP controlada hacia host receptor`.

Evidencia defendible:

- ZIP generado.
- Datos dummy en staging.
- Envío por red a otro host.
- Transferencia HTTP controlada.
- Acción equivalente a Invoke-WebRequest / Invoke-RestMethod / upload HTTP.
- `ReceiverReachable=True`.
- `HTTPStatus=200`.
- `UploadSucceeded=True`.
- SHA-256 local.
- Bytes transferidos en `ReceiverResponse`.

Limitación documental permitida:

Si no aparece `receiver_log.jsonl` o ZIP recibido en el paquete revisado, no invalida TEC-009. Se documenta como limitación secundaria sobre conservación del receptor final.

## Criterio Actualizado Sobre Benchmark

Benchmark final válido:

`08_MEMORIA/Excel_benchmark/TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx`

Estado:

- APTO.
- 9/9 runs válidos.
- 3 escenarios.
- 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- `RunnerStillRunningAtEnd=True` como WARN no bloqueante.
- 0 errores de fórmula.
- Gráficas pobladas.
- Fuentes con SHA-256.

Interpretación:

Velociraptor introduce una carga medible. En el laboratorio evaluado, el consumo medio fue bajo y estable, con media aproximada de 57 MB de RAM y 0,95 % de CPU durante el runner TEC.

## Criterio Actualizado Sobre Wazuh

- Wazuh se usa como validación comparativa HIDS/SIEM.
- No sustituye a Velociraptor.
- Wazuh trabaja con logs, reglas, archives y alerts.
- Velociraptor trabaja con artifacts, VQL, CLIENT_EVENT y evidencia host-based.
- Las unidades no son equivalentes directamente.
- Se conserva gap Wazuh TEC-009 / regla 110201.

Redacción vigente:

`Wazuh conserva una limitación en la regla 110201 para TEC-009, lo que evidencia que la cobertura depende de la lógica de reglas y de la fuente de eventos disponible.`

## Criterio Actualizado Sobre Artifacts

- P1/P2/P3/P4 como perfiles de artifacts son perfiles operativos de severidad/visibilidad.
- En documentos de planificación, P1/P2/P3/P4 son prioridades salvo que se indique `perfil artifact`.
- Router SOC_v3 es SERVER_EVENT y no detecta.
- Detección defendible ocurre en CLIENT_EVENT.
- JSONL y Discord son salida externa.
- P4 no es alerta crítica; es visibilidad básica.

## Próximos Pasos

1. Abrir `TFM_MEMORIA_BETA_v1.docx` en Word.
2. Revisar visualmente tablas, captions, saltos de página y anexos.
3. Adaptar bibliografía APA si la Universidad exige una variante concreta.
4. Confirmar capturas finales.
5. Revisar capítulos IV-V por coherencia metodológica.
6. Revisar capítulos VI-VIII contra Exceles finales.
7. Mantener TEC-009 como exfiltración HTTP controlada.
