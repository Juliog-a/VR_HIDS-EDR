README - ITERACION PROXIMA MEMORIA 10/10

Carpeta:
- 08_MEMORIA\ITERACION_PROXIMA_MEMORIA_10_10

Objetivo:
- Preparar la siguiente iteracion de la memoria sin modificar documentos originales.
- Priorizar trabajo que se puede hacer desde otro ordenador, sin laboratorio.
- Dejar claro que no se han ejecutado nuevas pruebas ni validaciones experimentales.

Archivos generados:
- 00_RESUMEN_EJECUTIVO_ITERACION.txt
- 01_CAPITULO_INTRODUCCION.txt
- 02_CAPITULO_ESTADO_ARTE.txt
- 03_CAPITULO_METODOLOGIA.txt
- 04_CAPITULO_ENTORNO_LABORATORIO.txt
- 05_CAPITULO_DESARROLLO_VELOCIRAPTOR.txt
- 06_CAPITULO_RESULTADOS.txt
- 07_CAPITULO_COMPARATIVA_WAZUH.txt
- 08_CAPITULO_CONCLUSIONES.txt
- 09_TRABAJO_PENDIENTE_LABORATORIO.txt
- 10_CHECKLIST_FINAL_10_10.txt
- 11_FUENTES_Y_REFERENCIAS_PENDIENTES.txt

Resumen de revision:
- La memoria principal `08_MEMORIA\TFM.docx` esta avanzada en capitulos I-V.
- El capitulo VIII Wazuh esta desarrollado, pero debe actualizarse con resultados finales de Velociraptor.
- Los capitulos VI, VII, IX y X son los principales huecos.
- Los datos finales ya existen en Exceles/CSV/JSON de 26/06.
- No es necesario ejecutar laboratorio para redactar resultados actuales.

Orden recomendado de trabajo:

1. ALTA - Capitulo VI Resultados
   - Incorporar 9/9 TEC OK.
   - Incorporar 379 filas CLIENT_EVENT.
   - Incorporar perfiles P1/P2/P3/P4.
   - Incorporar FP 10/10 OK y 0 hits.
   - Redactar limitacion TEC-009.

2. ALTA - Capitulo VII Benchmark
   - Incorporar benchmark APTO.
   - Explicar 9/9 runs validos, 3 escenarios, 3 repeticiones.
   - Explicar CPU/RAM y limitaciones.

3. ALTA - Capitulo VIII Wazuh
   - Sustituir "Velociraptor pendiente".
   - Incorporar comparacion final VR/Wazuh.
   - Mantener 110201=0 como gap.

4. ALTA - Capitulo X Conclusiones
   - Responder objetivos.
   - Evitar sobreafirmaciones.
   - Incluir limitaciones.

5. MEDIA - Capitulo IV/V
   - Ajustar metodologia final.
   - Reforzar arquitectura custom P1/P2/P3/P4.
   - Corregir terminologia Discord/JSONL.

6. MEDIA - Referencias y anexos
   - Meter fuentes oficiales.
   - Preparar anexos de artifacts, datasets y hashes.

Que hacer primero desde otro ordenador:
1. Abrir `06_CAPITULO_RESULTADOS.txt`.
2. Redactar Capitulo VI en la memoria usando:
   - 04_EVIDENCE\Excel_visibilidad_26062026\normalized\dataset_summary.json
   - 04_EVIDENCE\Excel_visibilidad_26062026\normalized\tecnicas_real_vr.csv
   - 04_EVIDENCE\Excel_visibilidad_26062026\normalized\alertas_vr.csv
3. Despues abrir `07_CAPITULO_COMPARATIVA_WAZUH.txt` y actualizar Capitulo VIII.
4. Despues abrir `08_CAPITULO_CONCLUSIONES.txt` y cerrar conclusiones.

Que dejar preparado para volver al laboratorio:
- Lista de capturas finales necesarias.
- Pregunta concreta sobre TEC-009 receiver_log/ZIP recibido.
- Si se decide, plan de revalidacion Wazuh 110201.
- Si se decide, plan de validacion router SOC/Discord.

Reglas metodologicas a mantener:
- CLIENT_EVENT detecta.
- SERVER_EVENT enruta, persiste o notifica.
- JSONL no es deteccion.
- Discord no es deteccion.
- Sysmon ID 26 es evidencia forense.
- No usar rutas/scripts/etiquetas TEC como IOC.
- No afirmar validacion experimental hecha por Codex.

Nota sobre OneNote:
- Se reviso de forma parcial por limitaciones tecnicas del formato `.one`.
- Las notas utiles localizadas coinciden en gran parte con Wazuh, objetivos, benchmark y tareas FP.
- OneNote debe tratarse como apoyo documental, no como evidencia primaria.
