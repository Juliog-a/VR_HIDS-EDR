# CLEANUP REPORT — TFM Velociraptor HIDS/DFIR

Fecha de cierre: 2026-09-08 (CEST).

## Alcance y trazabilidad

- Ruta auditada: `C:\Users\julio\Desktop\TFM`.
- Inventario previo: [`CLEANUP_INVENTORY_PRE.csv`](CLEANUP_INVENTORY_PRE.csv).
- Línea base fuera de `.git`: **6.889 archivos**, **681 directorios** y **134.889.442.118 bytes** (125,626 GiB).
- Clasificación previa: 342 archivos `KEEP`, 5.217 `DELETE` y 1.330 `REVIEW`.
- Se preservaron todos los archivos `KEEP` y `REVIEW`, la VM, los logs locales, el intercambio con la VM y todos los directorios `.git`.

## Eliminaciones

- Eliminados exactamente **5.217 archivos `DELETE`**.
- Espacio liberado: **1.236.425.453 bytes** (1.179,15 MiB).
- Eliminados **408 de 410 directorios `DELETE`**.
- Se conservan `08_MEMORIA/OLD_ENTREGA/EXCELS` y su padre únicamente porque contienen un `.git` anidado que no se ha tocado.
- La lista completa de archivos y directorios, junto con la justificación de cada clasificación, está en `CLEANUP_INVENTORY_PRE.csv` (filtrar `Classification=DELETE`).

Grupos principales retirados:

- `tmp/`, `99_ARCHIVE/`, `99_REVIEW_CLEANUP_20260629/` y `00_CONTEXT/OLD_CONTEXT_USELESS/`;
- artifacts `candidate`/`debug`, scripts candidatos y runners reemplazados;
- `08_MEMORIA/OLD_ENTREGA` salvo el `.git` anidado y sus dos directorios contenedores;
- iteraciones de QA, renders, cachés, exportaciones y documentos intermedios de la memoria;
- Excel antiguos reemplazados por los entregables finales;
- evidencias diagnósticas o legacy duplicadas;
- los archivos accidentales de 0 bytes de la raíz (`actualizar`, `adónde`, `cuánto`, `en`, `o`, `porque`, `término`, `únicamente`) y el nombre residual `tructure and cleanup legacy files…`.

## Duplicados

- Antes: 1.546 grupos SHA-256 y 5.162 archivos incluidos en algún grupo; el volumen estaba dominado por copias de `08_MEMORIA/ENTREGA`/`OLD_ENTREGA` y salidas de QA.
- Después: **11 grupos y 28 archivos** en el conjunto publicable.
- Se conservan porque representan fuentes/resultados equivalentes en ubicaciones con función documental distinta (por ejemplo FP frente a REAL, JSON frente a CSV o fuente frente a copia de análisis). No se eliminaron ante duda.

## Reorganizaciones y reparaciones de referencias

- No se realizó una reestructuración masiva.
- Se mantuvo la estructura numerada existente y se retiraron únicamente ramas obsoletas.
- El autodiagnóstico detectó nueve rutas de reproducibilidad que todavía apuntaban a copias eliminadas. Se redirigieron a los equivalentes canónicos de `01_ARTIFACTS/validated`, `02_SCRIPTS/validated` y `03_RUNNERS/validate`.
- Siete equivalentes conservan el mismo SHA-256; P4 y el router se actualizaron deliberadamente a los hashes de sus versiones finales validadas.
- Se actualizó `README.md` para reflejar estructura, técnicas TEC-001 a TEC-009, arquitectura, scripts, resultados, benchmark, Wazuh y entregables reales.
- Se depuró `.gitignore` y se añadió una excepción explícita para `10_WAZUH/TFM_WAZUH_VISIBILIDAD_EVIDENCE.tar.gz`, evidencia necesaria.

## Seguridad

Se neutralizaron antes del cierre:

- URLs completas de Discord webhook en fuentes de benchmark;
- la credencial del dashboard de Wazuh;
- rutas personales e IP privadas en los entregables y fuentes afectadas.

Marcadores utilizados: `YOUR_WEBHOOK_URL`, `<REDACTED>`, `<RAÍZ_TFM>` y `<HOST_LAB>`. El detalle y los SHA-256 posteriores se conservan en `CLEANUP_SANITIZATION_LOG.csv`.

Validación del conjunto publicable:

- webhooks Discord reales: 0;
- claves privadas: 0;
- tokens GitHub, AWS o Slack reconocibles: 0;
- los cinco avisos genéricos restantes son falsos positivos de TEC-007: parámetros `PasswordFile` y una clave AES ficticia y declarada para datos dummy;
- el TAR de evidencia Wazuh (41 entradas) se inspeccionó por separado: 0 coincidencias de webhook, clave privada o tokens conocidos.

No se ha reescrito ni borrado el historial Git. Los valores neutralizados existieron en commits anteriores; su revocación/rotación sigue siendo una medida externa recomendable.

## Entregables e integridad

Existen y no fueron eliminados:

- `08_MEMORIA/ENTREGA_TFM_Julgarrei/TFM.docx`;
- `08_MEMORIA/ENTREGA_TFM_Julgarrei/DEFENSA_TFM_Julio_Garcia_Amorena.pptx`;
- `08_MEMORIA/ENTREGA_TFM_Julgarrei/Analisis_Tecnicas_TFM_Velociraptor.xlsx`;
- `08_MEMORIA/ENTREGA_TFM_Julgarrei/TFM_BENCHMARK_RENDIMIENTO.xlsx`;
- los cinco artifacts de `01_ARTIFACTS/validated`;
- los Excel fuente finales de visibilidad y benchmark;
- la evidencia Wazuh empaquetada.

Validación OOXML:

- los cuatro entregables abren como ZIP OOXML y sus **198 partes XML/RELS** son válidas;
- Excel de técnicas: 9 hojas, 0 errores de fórmula serializados y 0 vínculos externos;
- Excel de benchmark: 11 hojas, 0 errores de fórmula serializados y 0 vínculos externos;
- presentación: 12 diapositivas;
- memoria: 3.565 párrafos, 105 tablas y 3 secciones.

La exportación visual adicional de la memoria no pudo completarse en esta sesión: LibreOffice no está instalado y Word COM quedó bloqueado en la exportación, por lo que se cerró el proceso sin modificar el DOCX. La revisión visual final previa ya estaba documentada; en esta limpieza se verificó de nuevo la integridad estructural del archivo sanitizado.

## Validaciones finales

- `06_CONTROLLED_RERUN/99_Static_SelfTest.ps1`: **PASS** — 17 scripts parseables, 9 hashes fuente y `TEST_MATRIX` 9/9.
- Parseo PowerShell del conjunto publicable: 38 archivos, 0 errores antes de las reparaciones; el self-test posterior confirma el paquete final.
- Archivos accidentales de 0 bytes en la raíz: 0.
- Archivos de 0 bytes publicables: 4, todos evidencia negativa ETW intencional (`REAL`/`FP`, CSV/JSON).
- Archivos mayores de 100 MB: 7, todos dentro de `lab/` y excluidos por `.gitignore`.
- Enlaces del nuevo `README.md`: comprobados contra rutas existentes.
- Remote: `https://github.com/Juliog-a/VR_HIDS-EDR.git`.

## Estructura final

```text
00_CONTEXT/
01_ARTIFACTS/validated/
02_SCRIPTS/validated/
03_RUNNERS/
04_EVIDENCE/
05_LOGS/                         # local, ignorado
06_CONTROLLED_RERUN/
07_DOCS/
08_MEMORIA/ENTREGA_TFM_Julgarrei/
10_WAZUH/
Carpeta_Compartida_TFM/          # local, ignorado
lab/                             # local, ignorado
```

Estado final fuera de `.git`: **1.677 archivos**, **273 directorios** y
**133.654.568.419 bytes** (124,476 GiB). La diferencia neta respecto a la
línea base incluye los nuevos inventario, informe y registro de sesión, además
de las sanitizaciones y actualizaciones documentales. El commit previsto es
`Final TFM repository cleanup and documentation`; su hash final se comunica
tras el push, ya que un commit no puede incluir autorreferencialmente su propio
hash.
