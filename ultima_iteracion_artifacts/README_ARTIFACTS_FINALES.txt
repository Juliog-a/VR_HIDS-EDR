README ARTIFACTS FINALES - TFM Velociraptor/HIDS
=================================================

Contenido
---------

artifacts_finales/
- Custom.TFM.HIDS.P1.Low.Basic_v3.yaml
- Custom.TFM.HIDS.P2.High.Forensic_v2.yaml
- Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml

artifacts_debug/
- Custom.TFM.Debug.Raw.PowerShell4104.LastHours.yaml
- Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours.yaml

backups_artifacts_previos/
- Copias de artifacts previos relevantes antes de esta iteracion.

analisis/
- DIAGNOSTICO_ARTIFACTS_ANTERIORES.md
- MATRIZ_COBERTURA_9_TECNICAS.md
- VALIDACION_REALIZADA.md

Modelo de ejecucion
-------------------

Estos artifacts son para backtesting/hunt historico.

Usan:
- type: CLIENT
- parse_evtx()

No son CLIENT_EVENT live.

Instrucciones de prueba en Velociraptor
---------------------------------------

1. Importar los YAML en Velociraptor.

2. Ejecutar primero los debug RAW:

   - Custom.TFM.Debug.Raw.PowerShell4104.LastHours
     LookbackHours=2

   - Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours
     LookbackHours=2

3. Confirmar que aparecen eventos:

   - PowerShell 4104 con EventData.ScriptBlockText.
   - Sysmon ID 1 con EventData.CommandLine.
   - Sysmon ID 3 si hubo conexion visible.
   - Sysmon ID 11 si hubo creacion ZIP/AES visible.

4. Ejecutar P1/P2/P3:

   LookbackHours=2
   ReceiverIP=192.168.1.129
   ReceiverPort=8088

Ventana temporal de campana OK
------------------------------

Campana v4:
- 2026-06-12 13:17:55 a 13:22:19

TEC-009:
- 2026-06-12 13:21:09 a 13:21:37

Como probar TEC-009
-------------------

1. Ejecutar P3:

   Custom.TFM.HIDS.P3.Medium.Behavioral_v12

2. Revisar la source:

   TEC009_HTTP_ZIP_Upload_4104

3. La alerta esperada debe salir si existe un PowerShell 4104 con:

   Invoke-WebRequest
   AND -InFile
   AND POST
   AND (application/zip OR .zip OR upload)

4. Sysmon ID 3 no es obligatorio.

Interpretacion si no aparece red
--------------------------------

Si TEC009_HTTP_ZIP_Upload_4104 aparece con Confidence=HIGH:
- TEC-009 queda detectada por evidencia fuerte 4104.
- La ausencia de Sysmon ID 3 no invalida la deteccion.
- Revisar receiver_log.jsonl y ZIP recibido como evidencia externa, no como condicion del artifact.

Si solo aparece TEC009_HTTP_Network_Sysmon_Context:
- Hay contexto de red.
- No prueba por si solo transferencia ZIP.
- Buscar 4104 RAW y receiver log.

Si solo aparece ZIP por Sysmon ID 11:
- Hay contexto de staging/archive.
- No prueba exfiltracion.
- Buscar 4104 fuerte.

Notas TEC-007 y TEC-008
-----------------------

TEC-007:
- La deteccion fuerte depende de 4104 con AES/CryptoStream/CreateEncryptor/.aes.
- Sysmon ID 11 .aes es contexto, no condicion obligatoria.

TEC-008:
- La deteccion fuerte depende de 4104 con Get-ChildItem + Remove-Item + -Force/-Recurse.
- Sysmon ID 23/26 es evidencia forense si existe.
- En este entorno Sysmon ID 23/26 no aparece y no debe bloquear la deteccion.

Limitaciones
------------

- La validacion VQL final debe hacerse importando en Velociraptor.
- Codex no ha ejecutado campana experimental.
- JSONL y Discord no forman parte de la deteccion.
- Router/Discord/receiver/runner no se han modificado.
