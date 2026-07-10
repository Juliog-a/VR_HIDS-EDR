# RESUMEN_CONTEXTO_CHATGPT_CODEX_PREPDF

Fecha: 2026-07-09

Uso previsto:

- Documento de contexto rápido para ChatGPT/Codex.
- Leer antes de continuar la memoria beta pre-PDF.
- Resume qué se ha hecho, qué se ha encontrado, qué queda pendiente y qué criterios metodológicos no se deben cambiar.

## Proyecto

Título del TFM:

`Análisis de la capacidad de Velociraptor para la detección de ataques HIDS de la matriz MITRE ATT&CK Enterprise`

Objetivo real del proyecto:

- Evaluar Velociraptor como aproximación HIDS/DFIR en Windows.
- Usar técnicas MITRE ATT&CK Enterprise ejecutadas en laboratorio.
- Limitar resultados al subconjunto experimental `TEC-001` a `TEC-009`.
- Separar ejecución, visibilidad, detección, alerta RT, evidencia forense y salida externa.

## Ruta principal actual

```text
C:\Users\julio\Desktop\TFM\08_MEMORIA\ITERACION_PROXIMA_MEMORIA_10_10
```

Documento original:

```text
C:\Users\julio\Desktop\TFM\08_MEMORIA\ITERACION_PROXIMA_MEMORIA_10_10\TFM_MEMORIA_BETA_v1.docx
```

Copia de revisión pre-PDF:

```text
C:\Users\julio\Desktop\TFM\08_MEMORIA\ITERACION_PROXIMA_MEMORIA_10_10\TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx
```

## Estado global

Veredicto vigente:

- `APTO PARA REVISIÓN FINAL HUMANA`
- `NO APTO PARA PDF FINAL TODAVÍA`

Porcentaje estimado de cierre: 88 %.

La memoria está cerca de versión final a nivel de contenido técnico, resultados y metodología, pero no debe exportarse todavía como PDF final porque quedan bloqueantes editoriales/metodológicos.

## Hashes relevantes

Original `TFM_MEMORIA_BETA_v1.docx`:

```text
06EA717AA88071F2C785689A4AAFDDA03B5943A85A75F3D99DBB4F788B8A6106
```

Copia revisada `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`:

```text
BB1F1464EC49D9C79DB5AC4E9C88468560E454D20877AEBC982B188C90B13E70
```

## Documentos generados en la revisión pre-PDF

En la carpeta de iteración:

- `REVISION_ESTADO_FINAL_BETA_10_10.md`
- `CHECKLIST_P1_ANTES_PDF.md`
- `CAMBIOS_APLICADOS_REVISION_PREPDF.md`
- `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`

En contexto:

- `00_CONTEXT/SESSION_LOGS/session_20260709_1820.md`
- `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`
- `00_CONTEXT/PROJECT_STATE.md` actualizado.
- `00_CONTEXT/CURRENT_TASK.md` actualizado.
- `00_CONTEXT/DECISIONS.md` actualizado con D048.

## Criterios metodológicos fijos

No cambiar salvo orden explícita del autor:

1. TEC-009 debe formularse como:
   - `exfiltración HTTP controlada en laboratorio`; o
   - `staging, archivado ZIP y exfiltración HTTP controlada hacia host receptor`.
2. No rebajar TEC-009 a `solo local`, `no válida`, `no exfiltración` o `solo staging`.
3. Si falta `receiver_log.jsonl` o ZIP recibido, eso es limitación documental secundaria, no invalidación de TEC-009.
4. `CLIENT_EVENT` es la unidad primaria de detección en Velociraptor.
5. `SERVER_EVENT` no detecta. Enruta, normaliza, persiste o notifica.
6. JSONL y Discord son salida externa, no fuente primaria de detección.
7. Benchmark mide coste operativo, no calidad de detección.
8. Wazuh es comparación HIDS/SIEM, no sustituto de Velociraptor.
9. No comparar directamente `archives/alerts` de Wazuh con `CLIENT_EVENT` de Velociraptor como unidades equivalentes.
10. P1/P2/P3/P4 en artifacts son perfiles operativos de severidad/visibilidad. En documentos de planificación pueden ser prioridades si se indica así.
11. Codex no valida experimentalmente. La validación experimental la hace el autor en laboratorio.

## Resultados consolidados que no deben alterarse

Velociraptor:

- Técnicas evaluadas: `TEC-001` a `TEC-009`.
- Cobertura real documentada: 9/9 técnicas.
- Filas `CLIENT_EVENT` consolidadas: 379.
- Perfiles:
  - P1 = 19
  - P2 = 118
  - P3 = 109
  - P4 = 133

Falsos positivos:

- FP-001 a FP-010 presentes.
- 10/10 OK.
- 0 WARN.
- 0 FAIL.
- 0 SKIPPED.
- `VelociraptorHits.Total = 0`.
- 82 filas `CLIENT_EVENT` en ventana FP no reconciliadas se documentan y no se usan como FP definitivo.

Benchmark:

- 9/9 runs válidos.
- 3 escenarios:
  - `BASELINE_NO_VR`
  - `VR_IDLE`
  - `VR_TEC_RUNNER`
- 3 repeticiones por escenario.
- `SERVER_GUI` excluido.
- `notepad.exe` runner = 0.
- RAM media aproximada: 57 MB.
- CPU media en `VR_TEC_RUNNER`: 0,95 %.
- Pico CPU cliente VR: 54 %.
- `RunnerStillRunningAtEnd=True` es WARN no bloqueante.

Wazuh:

- Wazuh base:
  - 902 archives.
  - 58 alerts.
- Wazuh custom:
  - 18.285 archives.
  - 110 alerts.
  - 56 alertas 110xxx.
- Regla `110201=0` se conserva como gap TEC-009.
- No afirmar superioridad absoluta.
- No afirmar equivalencia directa entre Wazuh y Velociraptor.

TEC-009:

- Tiene `HTTPStatus=200`.
- Tiene `UploadSucceeded=True`.
- Tiene SHA-256 local.
- Tiene bytes transferidos en `ReceiverResponse`.
- `receiver_log.jsonl` o ZIP recibido pueden faltar en paquete revisado sin invalidar la técnica.

## Qué se realizó en la revisión pre-PDF

1. Se leyó el contexto obligatorio del proyecto.
2. Se revisaron documentos de iteración:
   - `PLAN_PENDIENTE_POR_CAPITULOS.md`
   - `CAMBIOS_RECOMENDADOS_BETA_v1.md`
   - `CONTEXTO_MEMORIA_BETA_v1.md`
   - `README_ITERACION_BETA_v1.md`
   - `00_DOCUMENTOS_ARCHIVADOS_NO_USAR.md`
3. Se revisaron audits y data quality de Exceles finales:
   - visibilidad/detección/FP/Wazuh;
   - benchmark.
4. Se creó copia de trabajo:
   - `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
5. Se intentó revisión con Word COM.
   - Word quedó bloqueado/abierto y se pasó a revisión OpenXML.
6. Se revisó el paquete `.docx` como OpenXML:
   - documento válido;
   - 104 tablas;
   - 71 imágenes;
   - 76 dibujos;
   - 31 captions de código;
   - 0 campos `dirty`;
   - sin errores de referencia tipo `Error!`, `Marcador no definido` o `Reference source not found`.
7. Se aplicaron correcciones menores seguras en la copia.
8. Se generaron informes finales de revisión.
9. Se actualizó el cierre de sesión y el estado del proyecto.

## Correcciones aplicadas en la copia REVISION_PREPDF

Solo en:

```text
TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx
```

No en el original.

Cambios:

- `campana` -> `campaña`.
- `validos` -> `válidos`.
- `metricas` -> `métricas`.
- `linea` -> `línea`.
- `discord` -> `Discord`.
- Eliminado `Final del formulario`.
- `Tabla 47Casos...` -> `Tabla 47: Casos...`.
- `Tabla 51Perfiles...` -> `Tabla 51: Perfiles...`.
- `Desarrollos nuevos artifacts` -> `Desarrollo de nuevos artifacts`.
- `Automatización detecciones` -> `Automatización de detecciones`.
- `Integración SIEM / SOC...` -> `Integración SIEM/SOC...`.
- Se corrigieron dos sobrecorrecciones antes de cerrar:
  - `alíneada` -> `alineada`;
  - `runs_válidos.csv` -> `runs_validos.csv`.

No se tocaron:

- Criterio TEC-009.
- Cifras.
- Evidencias.
- Benchmark.
- Wazuh.
- Artifacts.
- Scripts.
- Runners.
- Exceles.
- Anexos completos.
- Índices de forma estructural.

## Hallazgos encontrados

### P1 / bloqueantes antes de PDF

1. Índice general duplicado.
   - Hay dos campos `TOC \o "1-3"`.
   - Tras el índice de figuras aparece otra vez el índice general completo.
   - Debe quedar un único índice general.

2. Índice de tablas desactualizado.
   - Se observa salto visible:
     `Tabla 55 -> Tabla 63 -> Tabla 57`.
   - En el cuerpo existen tablas posteriores con secuencia 56-64.
   - Requiere actualizar campos/listas en Word y verificar la secuencia.

3. Frase problemática en capítulo VI:

```text
Las salidas JSONL o Discord se usan como resultado final, ya que para que aparezcan debe haber detectado en cascada sistema-Velociraptor-Server Event-Discord, por lo que todas las partes deben funcionar correctamente.
```

Problema:

- Puede interpretarse como que JSONL/Discord son resultado primario de detección.
- Contradice el criterio fijo: JSONL/Discord son salida externa, no fuente primaria.

Redacción recomendada:

```text
Las salidas JSONL y Discord se conservan como salidas externas trazables derivadas del flujo de monitorización. Su presencia permite comprobar persistencia o notificación posterior, pero la unidad primaria de detección en Velociraptor sigue siendo la fila CLIENT_EVENT emitida por el artifact correspondiente.
```

4. Revisión visual humana pendiente.
   - Tablas.
   - Captions.
   - Saltos de página.
   - Anexos.
   - Páginas en blanco.
   - Exportación PDF provisional.

5. Word/temporal abierto.
   - Quedó temporal:

```text
~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx
```

Acción:

- Cerrar Word.
- Confirmar que desaparece el temporal.
- No exportar PDF mientras el documento esté bloqueado.

### P2 / altos

- Anexos A-L aparecen en el índice como:

```text
Capítulo XI: Anexo A...
Capítulo XII: Anexo B...
...
```

Recomendación:

- Si la plantilla lo permite, dejar formato más limpio:

```text
Anexo A. Técnicas MITRE seleccionadas y trazabilidad
```

- No cambiar automáticamente si puede romper estilos/índices.

### P3 / medios

- Cuatro tablas potencialmente anchas según estimación OpenXML:
  - tabla OpenXML 29;
  - tabla OpenXML 50;
  - tabla OpenXML 52;
  - tabla OpenXML 71.
- Una tabla grande:
  - tabla OpenXML 2, 39 filas.
- Deben revisarse visualmente en Word/PDF.

## Estado por capítulos

| Capítulo | Estado actual | Acción necesaria |
|---|---|---|
| I. Introducción | Avanzado | Revisión final de estilo y alcance. No sobrepromete MITRE completa. |
| II. Planificación | Avanzado | Revisar tiempos verbales para no parecer planificación futura. |
| III. Fundamentos teóricos | Avanzado | Revisar APA y densidad bibliográfica. |
| IV. Entorno/metodología | Avanzado | Revisar tablas, rutas/scripts y que no parezcan IOCs. |
| V. Desarrollo experimental | Avanzado | Mantener clara separación `CLIENT_EVENT`/`SERVER_EVENT`/JSONL/Discord. |
| VI. Resultados | Requiere intervención puntual | Reescribir frase JSONL/Discord. |
| VII. Benchmark | Cerrado a nivel de contenido | Revisión visual de tablas/gráficas. |
| VIII. Wazuh | Cerrado a nivel de contenido | Revisar tabla comparativa y gap 110201. |
| IX. Trabajo futuro | Avanzado | Títulos ya mejorados en copia. Revisar que no suene a tarea crítica pendiente. |
| X. Conclusiones | Avanzado | Verificar correspondencia objetivo/resultado/limitación. |
| Anexos A-L | Completos | Revisar paginación, citas desde cuerpo e índice. |

## Checklist exacto antes de PDF final

1. Cerrar Word.
2. Confirmar que desaparece `~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
3. Abrir `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
4. Reescribir la frase JSONL/Discord del capítulo VI.
5. Eliminar o regenerar el índice general duplicado.
6. Actualizar todos los campos:
   - índice general;
   - índice de tablas;
   - índice de figuras;
   - índice de código.
7. Confirmar que la secuencia de tablas no salta.
8. Revisar visualmente tablas anchas y tablas partidas.
9. Revisar captions de figuras/tablas/código.
10. Revisar anexos A-L.
11. Exportar PDF provisional.
12. Revisar PDF página a página.
13. Solo entonces exportar PDF final.

## Qué no hacer

- No modificar `TFM_MEMORIA_BETA_v1.docx` original sin decisión explícita.
- No tocar `01_ARTIFACTS/validated`.
- No cambiar cifras de resultados.
- No degradar TEC-009.
- No presentar JSONL/Discord como detección.
- No presentar `SERVER_EVENT` como detector.
- No mezclar benchmark con calidad de detección.
- No afirmar equivalencia directa Wazuh/Velociraptor.
- No afirmar cobertura de toda MITRE ni validez en producción.
- No afirmar validación experimental hecha por Codex.

## Ruta recomendada para continuar

Trabajar sobre:

```text
08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx
```

Consultar primero:

```text
08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/CHECKLIST_P1_ANTES_PDF.md
08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/REVISION_ESTADO_FINAL_BETA_10_10.md
08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/CAMBIOS_APLICADOS_REVISION_PREPDF.md
```

Si se necesita contexto completo para una nueva conversación de ChatGPT, pegar desde `Proyecto` hasta `Qué no hacer`.

