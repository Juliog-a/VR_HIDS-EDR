# CURRENT_TASK

Actualizado: 2026-09-08 09:58 CEST.

## Limpieza final del repositorio — CERRADA

- Eliminación autorizada y validaciones completadas.
- Reproducibilidad estática: PASS.
- Entregables finales e integridad OOXML: PASS.
- Seguridad del árbol actual: PASS; no se ha reescrito el historial.
- Informe: `CLEANUP_REPORT.md`.
- Único paso externo recomendado: revocar/rotar cualquier secreto histórico.

---

Actualizado: 2026-09-07 12:00 CEST.

## Limpieza final del repositorio — ABIERTA Y BLOQUEADA

- Auditoría previa e inventario `KEEP`/`DELETE`/`REVIEW` completados.
- 5.217 archivos y 1.179,15 MiB propuestos para eliminación.
- Redacción de credenciales, webhooks, rutas personales e IP privadas aplicada
  a los archivos de publicación afectados.
- Integridad ZIP/XML de los OOXML modificados comprobada.
- Eliminación recursiva bloqueada por la plataforma hasta recibir confirmación
  explícita del autor tras conocer el alcance.
- `git add`, commit y push no ejecutados.

Riesgo adicional: los webhooks y una credencial ya estaban seguidos por Git y
permanecen en el historial. No se reescribirá ni borrará historial sin una
decisión expresa.

Siguiente paso: confirmar la lista `DELETE`, completar validaciones, generar
`CLEANUP_REPORT.md` y preparar el commit local. El push seguirá condicionado a
resolver el riesgo de secretos históricos.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260907_1200.md`.

---

Actualizado: 2026-09-02 09:58 CEST.

## Presentación PowerPoint para defensa — CERRADA

Entregable:

`08_MEMORIA/ENTREGA/DEFENSA_TFM_Julio_Garcia_Amorena.pptx`

Resultado:

- 12 diapositivas 16:9;
- notas del orador 12/12;
- duración prevista próxima a 11 minutos;
- reapertura PowerPoint y revisión visual 12/12: `PASS`;
- SHA-256:
  `0F4851DCE2783E826C5E008D1421961FD07B430AEA4C1C42E81D631F12208891`.

No queda una tarea técnica abierta para Codex. El siguiente paso corresponde
al autor: revisar el archivo en el equipo de defensa y ensayar el guion oral.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260902_0958.md`.

---

## Última validación independiente del Excel de visibilidad — CERRADA

Archivo auditado, sin modificar:

`08_MEMORIA/Ultima_validacion/Analisis_Tecnicas_TFM_Velociraptor.xlsx`

Resultado: `NO-GO — CORREGIR ANTES DE ENTREGAR` (`6,2/10`).

Bloqueantes principales:

- tema OOXML con referencias internas a `ChatGPT`;
- 238 fechas mostradas como seriales en `07_Fuentes_Evidencias!E17:E254`;
- configuración de impresión inacabada: 241 páginas, páginas residuales y sin
  numeración.

No se modificó el Excel. El benchmark quedó fuera de alcance.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260828_1328.md`.

---

Actualizado: 2026-08-28 09:12 CEST.

## Auditoría de prioridad del maestro — CERRADA CON CORRECCIÓN PENDIENTE

La comparación entre
`08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx` y
`08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor_FINAL.xlsx` confirma que las
métricas principales se conservaron, pero la prioridad del maestro no se
cumple al 100 %.

Pendiente de autorización del autor:

- restaurar 25 celdas de trazabilidad en las que nombres reales de CSV fueron
  sustituidos por nombres derivados de hojas;
- repetir el QA de referencias tras la corrección.

No hay cambios pendientes en resultados experimentales.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260828_0912.md`.

---

Actualizado: 2026-08-27 16:01 CEST.

## Revisión final del Excel de visibilidad — CERRADA

Se han actualizado y validado:

- `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_Velociraptor_2.xlsx`.
- `08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_Velociraptor_FINAL.xlsx`.

Resultado: 17 hojas, 22 tablas, 4 gráficos y QA 25/25 PASS. Las dos
validaciones `#REF!` heredadas quedaron resueltas sin alterar los resultados.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260827_1601.md`.

## Tarea Excel cerrada

Se ha creado, sin modificar los cuatro originales:

`08_MEMORIA/Excel_visibilidad_FP/Analisis_Tecnicas_TFM_Velociraptor_2.xlsx`

Resultado:

- base exacta del maestro indicado por el autor;
- 2 hojas nuevas de dashboard/guía;
- 4 gráficos nuevos enlazados a datos canónicos;
- índice completo navegable;
- numeración y configuración de impresión en 30 hojas;
- QA estructural y visual superada.

Incidencia heredada no corregida: dos validaciones de datos `#REF!` ya
presentes en `04_Control_Publicos` y `05_Matriz_Resultados` del maestro.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260827_1044.md`.

## Tarea documental cerrada

Se han actualizado y verificado los documentos de `08_MEMORIA/Ultima_validacion`
para que Windows.ETW.Monitoring figure como campaña cerrada y concluyente y para
que las seis campañas de artifacts públicos consten como cerradas.

Entregables:

- `08_MEMORIA/Ultima_validacion/Ultima_iteracion.docx`.
- `08_MEMORIA/Ultima_validacion/Ultima_iteracion.pdf`.
- `08_MEMORIA/Ultima_validacion/ULTIMAS_CORRECCIONES.docx`.

Resultado:

- ETW: 0 filas TEC, 0 filas FP, 0 detecciones específicas y 0 alertas en la
  configuración evaluada; resultado negativo concluyente sin extrapolación.
- Seis campañas públicas cerradas; cinco con visibilidad positiva.
- Hayabusa Monitoring CH conserva la única detección pública específica: 3/9.
- REV-002, REV-013 y REV-027 alineadas con el criterio anterior.
- REV-001–REV-066 completas; revisión visual 40/40 y 14/14 páginas.

## Tarea abierta para el autor

Aplicar estas correcciones en copias de trabajo de los maestros:

- `08_MEMORIA/TFM.docx`: síntesis de resultados públicos, conclusión, tabla de
  artifacts públicos y valoración cualitativa de ETW.
- `08_MEMORIA/Analisis_Tecnicas_TFM_Velociraptor.xlsx`: hojas y celdas indicadas
  en REV-027 (`04_Control_Publicos`, `08_Benignas_FP`, `15_Publicos_Definitivo`,
  `17_Incoherencias_Cerradas`, `README`, `RESUMEN_EJECUTIVO`,
  `COMPARACION_VR_WAZUH` y `DISCREPANCIAS`).

`08_MEMORIA/TFM_BENCHMARK_RENDIMIENTO.xlsx` no contiene una clasificación de
cierre de la campaña ETW y no requiere modificación por este criterio.

Codex no ha validado experimentalmente el resultado: ha corregido la coherencia
documental a partir de los resultados consolidados del proyecto. La validación
experimental final corresponde al autor en laboratorio.

Registro: `00_CONTEXT/SESSION_LOGS/session_20260826_1437.md`.
