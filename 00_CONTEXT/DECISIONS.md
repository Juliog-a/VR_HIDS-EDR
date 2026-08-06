# DECISIONS

Actualizado: 2026-07-13 18:27 CEST.

Este fichero contiene solo decisiones metodológicas vigentes. La cronología
histórica previa se ha archivado en `00_CONTEXT/OLD_CONTEXT_USELESS`.

## D-FINAL-001 - Unidad primaria de detección

`CLIENT_EVENT` es la unidad primaria de detección en Velociraptor.

Una detección defendible debe poder trazarse a una fila emitida por un artifact
de cliente, con artifact, source, timestamp y evidencia asociados.

## D-FINAL-002 - Rol de SERVER_EVENT

`SERVER_EVENT` no detecta telemetría de host.

Su función es enrutar, normalizar, persistir o notificar eventos ya emitidos por
artifacts `CLIENT_EVENT`.

## D-FINAL-003 - JSONL y Discord

JSONL y Discord son salidas externas.

No son fuente primaria de detección y no deben usarse para sustituir la
evidencia `CLIENT_EVENT`.

## D-FINAL-004 - TEC-009

TEC-009 se formula como exfiltración HTTP controlada en laboratorio.

Incluye staging, archivado ZIP y transferencia HTTP controlada hacia un receptor
del investigador con datos dummy.

La evidencia mínima defendible para memoria es:

- `ReceiverReachable=True`.
- `HTTPStatus=200`.
- `UploadSucceeded=True`.
- SHA-256 local.
- Bytes transferidos en `ReceiverResponse`.

Si no aparece `receiver_log.jsonl` o ZIP recibido en el paquete revisado, se
documenta como limitación secundaria sobre conservación del receptor final, no
como invalidación de TEC-009.

## D-FINAL-005 - Wazuh

Wazuh se usa como comparación HIDS/SIEM.

No existe equivalencia directa entre:

- `archives` / `alerts` de Wazuh;
- filas `CLIENT_EVENT` de Velociraptor.

La regla Wazuh `110201=0` se conserva como gap TEC-009.

## D-FINAL-006 - Benchmark

El benchmark mide coste operativo observado en laboratorio.

No mide calidad de detección, cobertura, precisión, evidencia ni falsos
positivos.

Los resultados no se extrapolan a producción.

## D-FINAL-007 - Falsos positivos

Los falsos positivos definitivos se calculan solo con fuentes reconciliadas.

Filas fuera de ventana, contaminadas o no reconciliadas se documentan como
discrepancia, pero no se agregan como FP definitivo.

## D-FINAL-008 - Sysmon ID 26

Sysmon ID 26 es evidencia forense de borrado.

No se usa como alerta individual ni como criterio único de detección.

## D-FINAL-009 - Artifacts validados

No se sobrescriben artifacts en `01_ARTIFACTS/validated`.

Las versiones nuevas, si se requieren, se crean únicamente en
`01_ARTIFACTS/candidate` y no se leen por defecto.

## D-FINAL-010 - Candidatos y debug

Las carpetas `candidate` y `debug` son material no final.

Solo se consultan por petición expresa o si la tarea trata exactamente de
desarrollo/depuración de artifacts.

## D-FINAL-011 - Límites de alcance

El TFM no afirma cobertura de toda MITRE ATT&CK Enterprise.

El TFM no afirma validez operacional en producción.

Las conclusiones se limitan al laboratorio, técnicas y fuentes documentadas.

## D-FINAL-012 - Rol de Codex

Codex puede preparar scripts, revisar coherencia, generar documentación y
validar estructura local.

Codex no valida experimentalmente resultados de Velociraptor, Wazuh ni la VM.
La validación experimental corresponde al autor.

## D-FINAL-013 - Artifacts públicos y capas

Los artifacts públicos se interpretan por capacidad demostrada, no por nombre ni por volumen de filas.

- ProcessCreation, ServiceCreation, SysmonLogForward y TrackNetwork aportan visibilidad o evidencia forense según el caso.
- ServiceCreation 7045 acredita creación observada; no atribuye actividad maliciosa por sí solo.
- Hayabusa Monitoring CH puede acreditar detección por coincidencia Sigma; la cobertura específica final es 3/9.
- No existe campaña Hayabusa CHM separada y no se incorporan resultados Medium como si la hubiera.

## D-FINAL-014 - ETW

La campaña `Windows.ETW.Monitoring` queda `NO CONCLUYENTE`.

El resultado cero no se clasifica como negativo válido porque no se preservaron runner, log/summary TEC ni manifiesto de parámetros. Tampoco se concluye que ETW no funcione.

## D-FINAL-015 - TrackNetwork y salida HTTP

`Generic.Events.TrackNetworkConnections` acredita conexión visible a `192.168.1.129:8088`, pero las filas objetivo presentan `Timestamp=1601`, `PID=0` y `ProcInfo` vacío; no se atribuyen a PowerShell.

El `HTTPStatus=200`, `UploadSucceeded=True`, ZIP de 6.245 B y hash coincidente de TEC-009 se acreditan mediante el runner y la respuesta del receptor, de forma independiente al artifact TrackNetwork.

## D-FINAL-016 - Separación Wazuh base/custom

La comparación Wazuh usa como unidad el número de técnicas TEC únicas detectadas sobre las nueve ejecutadas, no el volumen de archives, alertas o rule_id.

- Wazuh base = ruleset nativo: 3/9, TEC-002, TEC-004 y TEC-006; reglas principales 92032, 92302, 91835 y 92077.
- Wazuh custom = reglas específicas TFM 110xxx: 4/9, TEC-001, TEC-005, TEC-007 y TEC-008; reglas principales 110301, 110402, 110203 y 110202.
- La unión base+custom alcanza 7/9 únicamente como dato complementario y no constituye una quinta solución.
- La baseline separada de 58 alertas no contiene marcadores TEC y no acredita cobertura técnica.
- La regla nativa 92307/T1543.003 es una alerta relacionada con creación de servicio, pero no se contabiliza como TEC-005/T1569.002.
- Alertas genéricas, coincidencias fuera de ventana, reglas con cero alertas y telemetría sin alerta específica no incrementan el numerador.

## D-FINAL-017 - Cierre del Excel centrado en artifacts custom

El maestro válido es la copia actual de `08_MEMORIA/ENTREGABLE`, respaldada antes de editar con SHA-256 `311DDEBEA38F523FFB63738017371121838EC4B108F549960E8B52A2484C4761`.

- `09_Benchmark_Plan` se sustituye por `09_Resultados_Custom`; el libro de visibilidad no contiene análisis de CPU, RAM, I/O, procesos ni escenarios B0-B8.
- Los resultados custom se derivan de `ALERTAS_VR`: P1=19, P2=118, P3=109, P4=133; total=379.
- Artifact principal y source interna se documentan por separado. Los cuatro P1-P4 son detectores CLIENT_EVENT; el Router es SERVER_EVENT de transporte/salida y no incrementa cobertura.
- Se conservan 30 hojas porque ese era el estado del maestro actual. No se recupera la hoja antigua `99_Listas`, que estaba vacía, únicamente para forzar un recuento de 31.
- Los Excel antiguos solo se usan como contraste. No se incorpora contenido obsoleto, duplicado o de rendimiento.
- El guardado definitivo y la reapertura se realizan con Microsoft Excel COM. El benchmark se verifica por tamaño/hash sin abrirlo ni guardarlo.

## D-FINAL-018 - Semántica final de filas CLIENT_EVENT y falsos positivos

- El total 379 se expresa como filas CLIENT_EVENT emitidas/exportadas,
  incluidas emisiones repetidas.
- No se presenta como 379 alertas únicas.
- Las 82 filas no reconciliadas se excluyen y no se computan como falsos
  positivos definitivos.
- `ALERTAS_VR!I383` se conserva como posible telemetría de validación y no se
  usa como soporte único de cobertura.
- La separación entre `TEC_escenario` y `TEC_detector/source` se documenta
  sin reconstruir valores no acreditados.

## D-FINAL-019 - Neutralización de información sensible en entregables

- Las rutas personales se sustituyen por `<RAÍZ_TFM>` o `<HOST_LAB>`.
- Las IP privadas visibles se sustituyen por `<HOST_LAB>`.
- Las URL completas de webhook se sustituyen por
  `<SALIDA_EXTERNA_REDACTADA>`.
- Esta neutralización se aplica también a la telemetría bruta incluida en los
  Excel finales, sin modificar fórmulas, gráficos ni valores métricos.
- Se conserva la trazabilidad funcional y se evita exponer credenciales o
  rutas personales en el paquete de entrega.

## D-FINAL-020 - Estado del paquete definitivo

- El paquete válido es `08_MEMORIA/ENTREGA/FINAL_PARA_SUBIR`.
- Debe contener exclusivamente los cuatro archivos indicados en
  `CURRENT_TASK.md`.
- El informe de validación permanece fuera de la carpeta de subida.
- Estado documental y técnico: `APTO PARA ENTREGA`.
