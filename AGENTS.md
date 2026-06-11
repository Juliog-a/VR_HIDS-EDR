\# AGENTS.md — TFM Velociraptor HIDS/DFIR



\## Lectura obligatoria al iniciar cualquier sesión

Lee siempre, en este orden:



1\. 00\_CONTEXT/PROJECT\_STATE.md

2\. 00\_CONTEXT/CURRENT\_TASK.md

3\. 00\_CONTEXT/DECISIONS.md

4\. 00\_CONTEXT/ARTIFACT\_INDEX.md

5\. 00\_CONTEXT/TEST\_MATRIX.md

6\. 00\_CONTEXT/EVIDENCE\_INDEX.md

7\. 00\_CONTEXT/SESSION\_LOGS/LAST\_SESSION.md

8\. 00\_CONTEXT/CODEX\_RULES.md

9\. 00\_CONTEXT/VM\_ACTIVE\_PATHS.md



\## Objetivo del proyecto

TFM sobre evaluación de Velociraptor como HIDS/DFIR en Windows para detectar técnicas MITRE ATT\&CK mediante artifacts personalizados.



\## Reglas obligatorias

\- No sobrescribir artifacts validados.

\- Crear versiones nuevas solo en 01\_ARTIFACTS/candidate.

\- No usar rutas de laboratorio, nombres de scripts, DUMB\_LAB, dump\_exfil, datos\_robados ni etiquetas TEC como IOC.

\- Separar siempre visibilidad, detección, alerta RT, evidencia forense y salida externa.

\- CLIENT\_EVENT detecta.

\- SERVER\_EVENT enruta, persiste o notifica.

\- JSONL no es detección.

\- Discord no es detección.

\- Codex no valida experimentalmente.

\- La validación final la hace el autor ejecutando campaña en laboratorio.

\- Sysmon ID 26 es evidencia forense, no alerta individual.

\- Documentar siempre cambios realizados, bugs encontrados y pendientes.



\## Cierre obligatorio de sesión

Antes de terminar cualquier sesión, crear o actualizar:



1\. 00\_CONTEXT/SESSION\_LOGS/session\_YYYYMMDD\_HHMM.md

2\. 00\_CONTEXT/SESSION\_LOGS/LAST\_SESSION.md

3\. 00\_CONTEXT/PROJECT\_STATE.md si cambia el estado global

4\. 00\_CONTEXT/CURRENT\_TASK.md si queda una tarea abierta

5\. 00\_CONTEXT/DECISIONS.md si se toma una decisión técnica



\## Formato de respuesta

\- Español de España.

\- Técnico.

\- Directo.

\- Esquemático.

\- Sin conclusiones no soportadas.

