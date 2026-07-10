# CODEX_RULES

Actualizado: 2026-07-10 10:26 CEST.

## Leer al iniciar

1. `AGENTS.md`.
2. `00_CONTEXT/README_CONTEXT.md`.
3. `00_CONTEXT/PROJECT_STATE.md`.
4. `00_CONTEXT/CURRENT_TASK.md`.
5. `00_CONTEXT/DECISIONS.md`.
6. Resto de ficheros obligatorios solo si la tarea lo requiere o `AGENTS.md`
   lo exige expresamente.

## No leer por defecto

- `00_CONTEXT/OLD_CONTEXT_USELESS`.
- Carpetas `candidate`.
- `01_ARTIFACTS/debug`.
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`.
- `99_ARCHIVE`.

## Reglas fijas

- No sobrescribir `01_ARTIFACTS/validated`.
- Crear artifacts nuevos solo en `01_ARTIFACTS/candidate`.
- No usar rutas de laboratorio, nombres de scripts, `DUMB_LAB`, `dump_exfil`,
  `datos_robados` ni etiquetas TEC como IOC.
- Separar visibilidad, detección, alerta RT, evidencia forense y salida externa.
- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, persiste, normaliza o notifica.
- JSONL no es detección.
- Discord no es detección.
- Sysmon ID 26 es evidencia forense, no alerta individual.
- Codex no valida experimentalmente.
- La validación final la hace el autor en laboratorio.

## Cierre de sesión

Crear o actualizar:

1. `00_CONTEXT/SESSION_LOGS/session_YYYYMMDD_HHMM.md`.
2. `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`.
3. `00_CONTEXT/PROJECT_STATE.md` si cambia estado global.
4. `00_CONTEXT/CURRENT_TASK.md` si queda tarea abierta.
5. `00_CONTEXT/DECISIONS.md` si se toma una decisión técnica.

Documentar:

- cambios realizados;
- bugs o contradicciones encontradas;
- archivos creados/modificados/movidos;
- pendientes;
- siguiente paso exacto.
