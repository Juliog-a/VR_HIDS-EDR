# CURRENT_TASK

Actualizado: 2026-07-10 10:26 CEST.

## Tarea actual

Continuar la revisión final pre-PDF de la memoria sobre:

`08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`

La limpieza de contexto de 2026-07-10 queda documentada en:

- `00_CONTEXT/LIMPIEZA_CONTEXTO_REPORT.md`.
- `00_CONTEXT/INVENTARIO_PROYECTO_LIMPIO.md`.
- `00_CONTEXT/INSTRUCCIONES_CODEX_FUTURAS.md`.

## Prioridad actual

Cerrar bloqueantes P1 antes del PDF final.

## Falta antes del PDF

1. Cerrar Word y confirmar que desaparece:
   `08_MEMORIA/ITERACION_PROXIMA_MEMORIA_10_10/~$M_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
2. Abrir `TFM_MEMORIA_BETA_v1_REVISION_PREPDF.docx`.
3. Eliminar o regenerar el índice general duplicado.
4. Actualizar índice general, índice de tablas, índice de figuras e índice de
   código.
5. Corregir el índice de tablas desactualizado.
6. Reescribir en capítulo VI la frase que trata JSONL/Discord como resultado
   final. Redacción recomendada:

   `Las salidas JSONL y Discord se conservan como salidas externas trazables derivadas del flujo de monitorización. Su presencia permite comprobar persistencia o notificación posterior, pero la unidad primaria de detección en Velociraptor sigue siendo la fila CLIENT_EVENT emitida por el artifact correspondiente.`

7. Revisar visualmente tablas anchas, captions, saltos de página y anexos A-L.
8. Exportar PDF provisional y revisarlo página a página.
9. Exportar PDF final solo si la revisión visual queda limpia.

## Qué no debe tocarse

- No modificar `TFM_MEMORIA_BETA_v1.docx` original salvo orden expresa.
- No tocar `01_ARTIFACTS/validated`.
- No cambiar cifras finales sin evidencia nueva explícita.
- No degradar TEC-009.
- No presentar JSONL, Discord ni `SERVER_EVENT` como detección primaria.
- No mezclar benchmark con calidad de detección.
- No afirmar equivalencia directa Wazuh/Velociraptor.
- No extrapolar a toda MITRE ni a producción.
- No leer `00_CONTEXT/OLD_CONTEXT_USELESS` salvo petición expresa.
- No leer carpetas `candidate` salvo petición expresa.

## Siguiente paso recomendado

Resolver los P1 de `CHECKLIST_P1_ANTES_PDF.md` en Word y actualizar los índices
antes de cualquier exportación final.
