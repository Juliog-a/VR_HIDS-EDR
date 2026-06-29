# Repetición controlada TFM Velociraptor

## Estado y alcance

Paquete para generar fuentes limpias antes de reconstruir CSV y Excel. No
modifica `01_ARTIFACTS/validated`, no activa Discord y no regenera Excel.

Los wrappers conservan cualquier JSONL previo, crean una salida nueva por
campaña/repetición y abortan ante rutas, hashes o precondiciones ambiguas.

### Parche TEC-009 4104 para FINAL02

`TEC_20260619_FINAL01` queda invalidada porque Windows PowerShell 5.1 lanzó una
excepción de conversión `Generic.List` y no se escribió el summary global.

La repetición `TEC_20260619_FINAL02` debe usar exclusivamente:

```text
03_RUNNERS\TFM_Run_All_TEC_Tests_v6_PATCHED_4104.ps1
```

Para desplegar el parche sin reutilizar el bundle anterior, ejecutar en host:

```powershell
cd C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
.\05_Stage_For_VM.ps1 `
  -DestinationRoot "C:\Users\julio\Desktop\TFM\Carpeta_Compartida_TFM\controlled_rerun_bundle_patch_4104"
```

Y en la VM, como administrador:

```powershell
cd \\VBOXSVR\Carpeta_Compartida_TFM\controlled_rerun_bundle_patch_4104\06_CONTROLLED_RERUN
.\06_Install_On_VM.ps1 -AllowPackageUpdate

cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\00_Preflight_Controlled_Rerun.ps1
.\09_Test_Patched_4104.ps1
```

No repetir FINAL02 si el test no termina con `PATCHED_4104 SELFTEST PASS`.

## Componentes

| Fichero | Ejecución | Función |
|---|---|---|
| `00_Preflight_Controlled_Rerun.ps1` | VM | Comprueba rutas, hashes, servicio y carpeta compartida |
| `00_Start_TEC_Receiver.ps1` | Host | Receiver TEC-009 aislado, un ZIP máximo |
| `01_Run_TEC_Canonical.ps1` | VM | TEC-001..TEC-009, JSONL exclusivo, snapshot y salida por técnica |
| `02_Run_FP_Controlled.ps1` | VM | Tres repeticiones FP-001..FP-010 con listener benigno FP-009 |
| `03_Run_Benchmark_3x.ps1` | VM | Tres repeticiones de los tres escenarios |
| `04_Validate_Controlled_Rerun.ps1` | Host | Valida TEC, FP, receiver y nueve runs benchmark |
| `05_Stage_For_VM.ps1` | Host | Prepara bundle inmutable en carpeta compartida |
| `06_Install_On_VM.ps1` | VM | Instala ficheros; con `-AllowPackageUpdate` respalda y actualiza solo el package, nunca `validated` |
| `07_Export_VM_Results.ps1` | VM | Copia resultados a la carpeta compartida con hashes |
| `08_Import_VM_Results.ps1` | Host | Verifica e importa resultados sin sobrescribir |
| `10_Build_ClientEvent_Exports_v2.ps1` | VM/Host | Builder corregido: ventana individual, parser CSV robusto, sanitización y reportes |
| `11_Verify_ClientEvent_Exports.ps1` | VM/Host | Verifica 12/12 CSV con Import-Csv, ventana y ausencia de multilineado físico |
| `CLIENT_EVENT_EXPORT_GUIDE.md` | VM/Host | Sources, ventanas UTC y comandos exactos de exportación/consolidación |

## 1. Preparar el bundle en el host

Abrir PowerShell en el host:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
cd C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
.\05_Stage_For_VM.ps1
```

Debe terminar indicando:

```text
Bundle VM preparado sin modificar validated
```

No reutilizar `controlled_rerun_bundle` para otra versión. Si ya existe, moverlo
como evidencia histórica antes de preparar otro bundle.

## 2. Instalar/verificar en la VM

Abrir PowerShell como administrador en la VM:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
cd \\VBOXSVR\Carpeta_Compartida_TFM\controlled_rerun_bundle\06_CONTROLLED_RERUN
.\06_Install_On_VM.ps1
```

El instalador:

- no escribe en `01_ARTIFACTS\validated`;
- sin `-AllowPackageUpdate`, no reemplaza un fichero distinto;
- con `-AllowPackageUpdate`, respalda y actualiza solo ficheros bajo
  `06_CONTROLLED_RERUN`; los artifacts/runners externos distintos siguen
  bloqueados.

## 3. Configurar Velociraptor antes de las pruebas

En la GUI, guardar/exportar primero la configuración actual. Después importar o
confirmar exactamente estos ficheros de la VM:

```text
C:\Users\seguridad\Desktop\TFM\01_ARTIFACTS\candidate\last_version\Custom.TFM.HIDS.P1.Critical.Priority.Event_v1.yaml
C:\Users\seguridad\Desktop\TFM\01_ARTIFACTS\candidate\Custom.TFM.HIDS.P2.High.Forensic.Event_v1.yaml
C:\Users\seguridad\Desktop\TFM\01_ARTIFACTS\candidate\Custom.TFM.HIDS.P3.Medium.Behavioral.Event_v1.yaml
C:\Users\seguridad\Desktop\TFM\01_ARTIFACTS\candidate\Custom.TFM.HIDS.P4.Low.Basic_v2.yaml
C:\Users\seguridad\Desktop\TFM\01_ARTIFACTS\candidate\last_version\Custom.TFM.HIDS.Router.JSONL.Discord.SOC_v3.yaml
```

Configurar Client Event Monitoring:

- P1 Critical: activo.
- P2 Event: activo.
- P3 Event: activo.
- P4 v2: activo.
- P1 legacy: desactivado.

Configurar Server Event Monitoring:

- SOC_v3: activo.
- SOC_v1 y SOC_v2: desactivados.
- `EnableP1Critical=true`.
- `EnableP1=false`.
- `EnableP2Event=true`.
- `EnableP3=true`.
- `EnableP4=true`.
- `EnableJSONL=true`.
- `EnableDiscord=false`.
- `JsonlPath=\\VBOXSVR\Carpeta_Compartida_TFM\inbox\soc_alerts.jsonl`.

Esperar al menos 60 segundos desde la activación del monitoring antes del
preflight. Sincronizar el reloj de host y VM.

## 4. Preflight en la VM

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\00_Preflight_Controlled_Rerun.ps1
```

No continuar si aparece `Preflight NO APTO`. El preflight automático no puede
leer la selección de artifacts de la GUI; esa confirmación sigue siendo manual.

## 5. Campaña TEC canónica

Elegir un identificador único. El mismo texto debe usarse en host y VM. Ejemplo:

```text
TEC_20260619_FINAL02
```

### 5.1. Host: iniciar receiver

En una consola del host, que debe permanecer abierta:

```powershell
cd C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
.\00_Start_TEC_Receiver.ps1 -CampaignId "TEC_20260619_FINAL02"
```

El receiver escucha en `0.0.0.0:8088`, acepta un único upload y guarda todo en:

```text
06_CONTROLLED_RERUN\OUTPUT\RECEIVER\TEC_20260619_FINAL02
```

### 5.2. VM: ejecutar TEC-001..TEC-009

En otra consola administrativa de la VM:

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\01_Run_TEC_Canonical.ps1 `
  -CampaignId "TEC_20260619_FINAL02" `
  -ReceiverUrl "http://192.168.1.129:8088/upload" `
  -IConfirmMonitoringConfigured
```

Sustituir `192.168.1.129` solo si la IP real del host es distinta. No usar una
IP de terceros.

Salida VM:

```text
06_CONTROLLED_RERUN\OUTPUT\TEC\TEC_20260619_FINAL02
```

Debe contener JSONL raw, summary del runner, snapshot de configuración, hashes,
evidencia TEC-009 y subcarpetas `by_technique\TEC-001` a `TEC-009`.

El summary canónico debe existir también en la raíz como:

```text
OUTPUT\TEC\TEC_20260619_FINAL02\runner_summary.json
```

## 6. Exportar detecciones originales desde Velociraptor

Usar `START_UTC.txt` y `END_UTC.txt` de la campaña TEC para limitar el intervalo.
Exportar las filas originales de Client Event Monitoring, no las del router.

Guardar en:

```text
OUTPUT\TEC\TEC_20260619_FINAL02\velociraptor_exports
```

con estos nombres exactos:

```text
P1_CRITICAL_CLIENT_EVENT.csv
P2_EVENT_CLIENT_EVENT.csv
P3_EVENT_CLIENT_EVENT.csv
P4_CLIENT_EVENT.csv
```

Los CSV deben conservar Artifact, Source, timestamp, TEC/CU, evidencia y campos
originales. No editar conteos ni estados manualmente.

## 7. Falsos positivos

Ejecutar en la VM. El wrapper crea y controla el listener benigno de FP-009 en
`127.0.0.1:80`; aborta si el puerto ya estaba ocupado.

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\02_Run_FP_Controlled.ps1 `
  -CampaignId "FP_20260619_FINAL01" `
  -Repetitions 3 `
  -IConfirmMonitoringConfigured
```

Cada `REP_01`, `REP_02` y `REP_03` tiene su JSONL propio. No concatenarlos.

Después, usando `StartUtc`/`EndUtc` de cada `repetition_manifest.json`, exportar
P1/P2/P3/P4 con los mismos cuatro nombres dentro de:

```text
OUTPUT\FP\FP_20260619_FINAL01\REP_01\velociraptor_exports
OUTPUT\FP\FP_20260619_FINAL01\REP_02\velociraptor_exports
OUTPUT\FP\FP_20260619_FINAL01\REP_03\velociraptor_exports
```

## 8. Benchmark: tres escenarios por tres repeticiones

No necesita receiver. El lanzador de carga desactiva expresamente la
exfiltración para que el benchmark no se confunda con la evidencia TEC-009.

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\03_Run_Benchmark_3x.ps1 `
  -CampaignId "BENCH_20260619_FINAL01" `
  -Repetitions 3
```

El proceso tarda aproximadamente 40 minutos. Es fail-fast: si un escenario no
es `VALID`, conserva la evidencia y no continúa con datos incorrectos.

Al terminar debe haber nueve `summary.json`: tres `BASELINE_NO_VR`, tres
`VR_IDLE` y tres `VR_TEC_RUNNER`.

## 9. Transferir resultados VM al host

Después de añadir todos los exports de Velociraptor, ejecutar en la VM una
transferencia selectiva. Así `FINAL01` TEC no entra en el paquete:

```powershell
cd C:\Users\seguridad\Desktop\TFM\06_CONTROLLED_RERUN
.\07_Export_VM_Results.ps1 `
  -TransferId "RERUN_20260619_FINAL02" `
  -TecCampaignId "TEC_20260619_FINAL02" `
  -FpCampaignId "FP_20260619_FINAL01" `
  -BenchmarkCampaignId "BENCH_20260619_FINAL01"
```

En el host:

```powershell
cd C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
.\08_Import_VM_Results.ps1 -TransferId "RERUN_20260619_FINAL02"
```

El receiver ya está en el host y no se importa desde la VM.

## 10. Validación final en el host

```powershell
cd C:\Users\julio\Desktop\TFM\06_CONTROLLED_RERUN
.\04_Validate_Controlled_Rerun.ps1 `
  -TecCampaignId "TEC_20260619_FINAL02" `
  -FpCampaignId "FP_20260619_FINAL01" `
  -BenchmarkCampaignId "BENCH_20260619_FINAL01"
```

Solo se podrán reconstruir CSV y Excel cuando el informe termine con:

```text
VALIDACIÓN PASS
```

Informes generados:

```text
OUTPUT\VALIDATION\VALIDATION_<timestamp>\VALIDATION_REPORT.md
OUTPUT\VALIDATION\VALIDATION_<timestamp>\VALIDATION_REPORT.json
OUTPUT\VALIDATION\VALIDATION_<timestamp>\VALIDATION_RESULTS.csv
```

## Prohibiciones durante la campaña

- No abrir ni editar los JSONL mientras se ejecuta una prueba.
- No unir JSONL de campañas o repeticiones.
- No cambiar `FAIL`, `WARN`, `INVALID` o `SKIPPED` manualmente.
- No utilizar JSONL, router o Discord como sustituto de `CLIENT_EVENT`.
- No reutilizar un `CampaignId`.
- No regenerar Excel antes de obtener `VALIDACIÓN PASS`.
