# LAST_SESSION

## Resumen

Cierre formal de la iteracion de artifacts historicos documentado en:

- `00_CONTEXT/SESSION_LOGS/session_20260612_1418.md`

Trabajo principal:

- Rehacer P1/P2/P3 como artifacts `CLIENT` de backtesting con `parse_evtx()`.
- Priorizar TEC-009 por PowerShell 4104 fuerte.
- Cubrir 9/9 tecnicas TEC-001 a TEC-009.
- Crear debug RAW para confirmar lectura EVTX.
- Crear paquete `ultima_iteracion_artifacts`.

## Cambios realizados

- Creados:
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P1.Low.Basic_v3.yaml`
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P2.High.Forensic_v2.yaml`
  - `01_ARTIFACTS/candidate/Custom.TFM.HIDS.P3.Medium.Behavioral_v12.yaml`
  - `01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.PowerShell4104.LastHours.yaml`
  - `01_ARTIFACTS/debug/Custom.TFM.Debug.Raw.Sysmon.ID1_3_11.LastHours.yaml`

- Creados dentro de `ultima_iteracion_artifacts`:
  - `artifacts_finales`
  - `artifacts_debug`
  - `backups_artifacts_previos`
  - `analisis`
  - `hashes`
  - `README_ARTIFACTS_FINALES.txt`

- Creado ZIP final:
  - `ultima_iteracion_artifacts.zip`
  - SHA256: `8247413CA455681AF4007EF33FFA89040A4955885028276B50E7F2DCE773365A`

## Bugs/problemas encontrados

- `00_CONTEXT/TEST_MATRIX.md` no existe.
- P2/P3 previos eran live (`CLIENT_EVENT` + `watch_evtx()`), no aptos para backtesting de campana ya ejecutada.
- P3_v11 no cubria TEC-007/008/009 como deteccion fuerte.
- No se pudo compilar VQL localmente por ausencia de binario Velociraptor.
- No se pudo parsear YAML con parser local por ausencia de dependencias.

## Validacion realizada

- Validacion estructural local OK:
  - `type: CLIENT`;
  - `parse_evtx()`;
  - sin consultas `watch_evtx()`;
  - sin `Generic.Events.TrackNetworkConnections`;
  - sin tabuladores;
  - claves `name`, `parameters`, `sources`, `query`.

## Validacion pendiente

- Importar artifacts en Velociraptor.
- Ejecutar debug RAW 4104 y Sysmon.
- Ejecutar P1/P2/P3 con `LookbackHours=2`, `ReceiverIP=192.168.1.129`, `ReceiverPort=8088`.
- Confirmar TEC-009 por P3 4104 aunque no aparezca red.
- Confirmar TEC-007/008 por 4104 fuerte aunque falten Sysmon ID 11 o 23/26.

## Estado final

- Artifacts candidate creados.
- Artifacts previos respaldados.
- Runner, scripts de ataque, receiver, router y Discord no modificados.
- Validacion experimental pendiente por el autor.
