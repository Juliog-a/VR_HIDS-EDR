# AGENTS.md — TFM Velociraptor HIDS/DFIR

## Lectura obligatoria al iniciar cualquier sesión

Lee siempre, en este orden:

1. `00_CONTEXT/README_CONTEXT.md`
2. `00_CONTEXT/PROJECT_STATE.md`
3. `00_CONTEXT/CURRENT_TASK.md`
4. `00_CONTEXT/DECISIONS.md`
5. `00_CONTEXT/ARTIFACT_INDEX.md`
6. `00_CONTEXT/TEST_MATRIX.md`
7. `00_CONTEXT/EVIDENCE_INDEX.md`
8. `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`
9. `00_CONTEXT/CODEX_RULES.md`
10. `00_CONTEXT/VM_ACTIVE_PATHS.md`

No leer por defecto:

- `00_CONTEXT/OLD_CONTEXT_USELESS`
- carpetas `candidate`
- `01_ARTIFACTS/debug`
- `03_RUNNERS/OLD_NO_LEER_POR_DEFECTO`
- `99_ARCHIVE`

## Objetivo del proyecto

TFM sobre evaluación de Velociraptor como HIDS/DFIR en Windows para detectar
técnicas MITRE ATT&CK mediante artifacts personalizados.

## Reglas obligatorias

- No sobrescribir artifacts validados.
- Crear versiones nuevas solo en `01_ARTIFACTS/candidate`.
- No usar rutas de laboratorio, nombres de scripts, `DUMB_LAB`, `dump_exfil`,
  `datos_robados` ni etiquetas TEC como IOC.
- Separar siempre visibilidad, detección, alerta RT, evidencia forense y salida
  externa.
- `CLIENT_EVENT` detecta.
- `SERVER_EVENT` enruta, persiste o notifica.
- JSONL no es detección.
- Discord no es detección.
- Codex no valida experimentalmente.
- La validación final la hace el autor ejecutando campaña en laboratorio.
- Sysmon ID 26 es evidencia forense, no alerta individual.
- Documentar siempre cambios realizados, bugs encontrados y pendientes.

## Cierre obligatorio de sesión

Antes de terminar cualquier sesión, crear o actualizar:

1. `00_CONTEXT/SESSION_LOGS/session_YYYYMMDD_HHMM.md`
2. `00_CONTEXT/SESSION_LOGS/LAST_SESSION.md`
3. `00_CONTEXT/PROJECT_STATE.md` si cambia el estado global
4. `00_CONTEXT/CURRENT_TASK.md` si queda una tarea abierta
5. `00_CONTEXT/DECISIONS.md` si se toma una decisión técnica

## Formato de respuesta

- Español de España.
- Técnico.
- Directo.
- Esquemático.
- Sin conclusiones no soportadas.

