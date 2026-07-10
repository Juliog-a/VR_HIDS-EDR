# Cambios Recomendados Beta v1

Fecha: 2026-07-08

## Cambios Aplicados

- Generada nueva beta Word desde `08_MEMORIA/TFM.docx`, que era el documento TFM real más actualizado.
- Archivada la beta previa y los documentos auxiliares antiguos en `_ARCHIVO_OBSOLETO_NO_USAR`.
- Actualizado resumen y abstract con resultados consolidados: 9/9 TEC, 379 filas CLIENT_EVENT, FP 10/10 con 0 hits, benchmark 9/9 runs válidos y Wazuh como contraste.
- Corregido TEC-009 en metodología, selección de técnicas, desarrollo y resultados.
- Sustituida formulación antigua de TEC-009 por `exfiltración HTTP controlada en laboratorio`.
- Añadido T1048.003 como parte de TEC-009 junto a T1074.001 y T1560.001.
- Eliminadas formulaciones antiguas que rebajaban TEC-009 a preparación local o invalidaban el envío HTTP documentado.
- Actualizada tabla comparativa Wazuh: eliminados `Pendiente de campaña VR` y `Pendiente`.
- Reforzada comparación Wazuh/Velociraptor como modelos no equivalentes directamente.
- Reescrito benchmark para mantenerlo como coste operativo, no calidad de detección.
- Reforzada la idea de que JSONL y Discord son salida externa, no fuente primaria.
- Añadido índice de archivado y plan pendiente por capítulos.
- Insertadas citas en el cuerpo de la memoria para MITRE, Velociraptor, Sysmon, PowerShell Logging, Wazuh y fuentes internas del TFM.
- Sustituida la bibliografía mínima por referencias APA normalizadas.
- Desarrollados anexos A-L dentro de `TFM_MEMORIA_BETA_v1.docx`.
- Abierto el documento con Word en modo invisible, actualizado el índice/listas y guardado correctamente.

## Cambios Recomendados No Aplicados

- No se exportó PDF final.
- No se revisó visualmente el ajuste de tablas, captions y saltos de página.
- No se localizaron nuevas evidencias de receptor final TEC-009.
- No se revalidó Wazuh 110201.
- No se ejecutaron pruebas nuevas de laboratorio.

## Decisiones Metodológicas

- `CLIENT_EVENT` es la unidad primaria de detección Velociraptor.
- `SERVER_EVENT` enruta, normaliza, persiste o notifica; no detecta telemetría de host.
- JSONL y Discord son salida externa.
- Sysmon ID 26 se mantiene como evidencia forense, no alerta individual.
- Benchmark mide coste operativo observado, no calidad de detección.
- Wazuh se mantiene como comparación HIDS/SIEM basada en logs, reglas, archives y alerts.
- Velociraptor se mantiene como plataforma DFIR extensible con capacidad HIDS cuando se diseña detection engineering mediante artifacts.
- TEC-009 se formula como staging, archivado ZIP y exfiltración HTTP controlada hacia host receptor.

## Mejoras de Redacción

- Sustituido lenguaje provisional por resultados consolidados.
- Eliminadas expresiones ambiguas como `posible exfiltración contextual`.
- Mejorada la redacción del benchmark con consumo medible, medio bajo y no extrapolable.
- Sustituida comparación `equivalente` por comparación `metodológicamente comparable`.
- Corregidas frases coloquiales o poco académicas en capítulos VII y VIII.

## Riesgos Detectados

- Revisión visual pendiente: tablas, captions, saltos de página y anexos deben revisarse manualmente en Word.
- Bibliografía APA insertada; puede requerir adaptación si la Universidad exige una variante formal concreta.
- Capturas finales no cerradas.
- TEC-009 depende de summary para documentar envío HTTP correcto; si no aparece `receiver_log.jsonl` o ZIP recibido en el paquete revisado, debe mantenerse como limitación documental secundaria.
- La muestra FP es controlada y limitada; no permite inferir tasa estadística universal.
- La cobertura 9/9 se limita a TEC-001..TEC-009 en laboratorio.
- Wazuh 110201 sigue siendo gap y no debe ocultarse.

## Dudas Metodológicas Críticas

No queda una duda metodológica crítica bloqueante para generar la beta.

La única duda relevante es documental y no invalida TEC-009: confirmar si existe `receiver_log.jsonl` o ZIP recibido asociado a la campaña final. Si no aparece, conservar la redacción actual:

> exfiltración HTTP controlada documentada por summary, con limitación documental sobre conservación del receptor final.
