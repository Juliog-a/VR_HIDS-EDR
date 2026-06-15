# Validacion realizada

## Validado realmente

- Se leyo el contexto obligatorio de `00_CONTEXT`.
- `00_CONTEXT/TEST_MATRIX.md` no existe; queda registrado como hueco documental.
- Se revisaron artifacts previos:
  - `Custom.TFM.HIDS.P2.High.Forensic_v1.yaml`
  - `Custom.TFM.HIDS.P3.Medium.Behavioral_v11.yaml`
  - `Custom.TFM.HIDS.P4.Low.Basic_v2.yaml`
  - debug PowerShell 4104 previos
- Se reviso el Excel `04_EVIDENCE/Analisis_Tecnicas_TFM_Velociraptor_DEF.xlsx`.
- Se revisaron evidencias locales:
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_131755_summary.txt`
  - `05_LOGS/TFM_TEC_Run_CANDIDATE_v4_20260612_131755_summary.csv`
  - `05_LOGS/TFM_Sysmon_Visibility_20260612_132606_summary.txt`
  - `05_LOGS/CU009Shape_v4_validacion_4104_TEC009.json`
- Se confirmo que los YAML nuevos:
  - son `type: CLIENT`;
  - usan `parse_evtx()`;
  - no usan `watch_evtx()` en consultas;
  - no usan `Generic.Events.TrackNetworkConnections`;
  - no filtran por rutas de laboratorio, nombres de scripts, `DUMB_LAB` ni etiquetas TEC como IOC.
- Se sustituyo `IN (...)` por condiciones `OR` para mayor compatibilidad VQL.
- Se comprobaron conteos estructurales de `parse_evtx()`/`watch_evtx()` y ausencia de tabuladores.
- Validacion estructural local:
  - los cinco YAML tienen `name`, `type: CLIENT`, `parameters`, `sources` y `query`;
  - P1: 4 consultas `parse_evtx()`;
  - P2: 8 consultas `parse_evtx()`;
  - P3: 10 consultas `parse_evtx()`;
  - debug 4104: 1 consulta `parse_evtx()`;
  - debug Sysmon: 1 consulta `parse_evtx()`;
  - ninguna consulta usa `watch_evtx()`;
  - ninguna consulta usa `Generic.Events.TrackNetworkConnections`.

## No validado realmente

- No se pudo ejecutar `velociraptor artifacts verify` ni compilacion VQL local.
- No se encontro binario `velociraptor` en el PATH ni en el arbol del proyecto.
- No habia parser YAML local disponible (`ruby` no instalado; Node sin `yaml`/`js-yaml`).
- `python.exe` no arranco en este entorno, por lo que no se uso Python para validacion.

## Pendiente en GUI de Velociraptor

1. Importar los cinco YAML nuevos.
2. Ejecutar primero:
   - `Custom.TFM.Debug.Raw.PowerShell4104.LastHours`
   - `Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours`
3. Ejecutar P1/P2/P3 con:
   - `LookbackHours=2`
   - `ReceiverIP=192.168.1.129`
   - `ReceiverPort=8088`
4. Confirmar que P3 emite `TEC009_HTTP_ZIP_Upload_4104` aunque no haya fila de Sysmon ID 3.
5. Confirmar que TEC-007 y TEC-008 emiten por 4104 fuerte aunque falten Sysmon ID 11 o 23/26.
6. Revisar que las filas de red y fichero se interpretan como contexto o evidencia forense.

## Riesgo residual

- La sintaxis VQL no se ha compilado con el motor real de Velociraptor en este entorno.
- La validacion experimental final sigue correspondiendo al autor en laboratorio.
