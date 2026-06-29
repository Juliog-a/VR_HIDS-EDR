# CLEANUP_20260629

Fecha: 2026-06-29.

## Objetivo

Preparar el repositorio para GitHub y limpiar la raiz sin perder trazabilidad
util para el TFM.

## Cambios de estructura

- `README.md` sustituido por README de proyecto para GitHub.
- Informes legacy de fiabilidad/FP movidos a:
  `04_EVIDENCE/legacy_reports/20260618_fiabilidad_fp/`.
- Benchmark legacy movido a:
  `04_EVIDENCE/legacy_reports/20260618_benchmark_legacy/`.
- Paquetes historicos `ultima_iteracion*` movidos a:
  `99_ARCHIVE/legacy_iterations/`.
- Prompts antiguos movidos a:
  `99_ARCHIVE/legacy_prompts_20260610/`.
- `README_TFM.md` vacio y `tatus` accidental movidos localmente a:
  `99_REVIEW_CLEANUP_20260629/delete_candidates/`.

## Limpieza GitHub

Retirados del indice Git sin borrar del disco:

- `04_EVIDENCE/diag_p3v3/02_alerts.jsonl`
- `08_MEMORIA/Planificacion/GanttProject.lnk`
- `08_MEMORIA/Planificacion/Toggl Track.lnk`
- `08_MEMORIA/TFM/Master.one`
- `08_MEMORIA/TFM/OneNote_RecycleBin/OneNote_DeletedPages.one`
- `08_MEMORIA/TFM/TFM.one`

Motivo: son ficheros locales/brutos ya cubiertos por `.gitignore`.

Limpieza interna Git:

- Ejecutado `git gc --prune=now`.
- Resultado posterior: `garbage=0`, `size-garbage=0 bytes`.
- El pack historico queda en `3.08 GiB`; no se reescribio historial.

## Elementos no movidos por riesgo operativo

- `lab/`: contiene VM local y snapshots. Moverla puede romper la configuracion
  de VirtualBox.
- `Carpeta_Compartida_TFM/`: intercambio con VM/laboratorio.
- `05_LOGS/`: evidencias brutas locales, ignoradas para GitHub.
- `06_CONTROLLED_RERUN/OUTPUT/`: salidas de campanas, ignoradas para GitHub.
- `10_WAZUH/*.iso` y paquetes comprimidos: material local pesado, ignorado.

## Reglas respetadas

- No se modifico `01_ARTIFACTS/validated`.
- No se reinterpreto JSONL como deteccion.
- No se ejecuto validacion experimental.
- No se activo Discord.
- No se eliminaron evidencias utiles; los elementos dudosos quedaron en carpeta
  local de revision.

## Bugs o residuos detectados

- `README.md` tenia solo el titulo `VR_HIDS-EDR`.
- `README_TFM.md` estaba vacio.
- `tatus` contenia una salida accidental de diffstat y no documentacion real.
- El repositorio mantiene material local muy pesado ignorado por Git y un pack
  historico de 3.08 GiB; revisar antes de publicar si se necesita reescribir
  historial Git.
