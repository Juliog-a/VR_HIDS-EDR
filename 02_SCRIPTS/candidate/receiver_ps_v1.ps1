param(
    [string]$BindPrefix = "http://+:8000/",
    [string]$OutDir = ".\received_ps",
    [string]$LogFile = ".\received_ps\receiver_log.jsonl",
    [int]$MaxSeconds = 600,
    [int]$MaxUploads = 5
)

$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LogFile) | Out-Null

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($BindPrefix)
$listener.Start()

Write-Host "TFM PowerShell receiver listening on $BindPrefix"
Write-Host "Output dir: $OutDir"
Write-Host "Log file: $LogFile"
Write-Host "Limits: max_seconds=$MaxSeconds max_uploads=$MaxUploads"

$uploads = 0
$deadline = (Get-Date).AddSeconds($MaxSeconds)

:main while ((Get-Date) -lt $deadline -and $uploads -lt $MaxUploads) {
    $task = $listener.GetContextAsync()

    while (-not $task.IsCompleted) {
        if ((Get-Date) -ge $deadline) { break main }
        Start-Sleep -Milliseconds 200
    }

    $ctx = $task.Result
    $req = $ctx.Request
    $res = $ctx.Response
    $res.KeepAlive = $false

    try {
        if ($req.Url.AbsolutePath -eq "/health") {
            $body = [Text.Encoding]::UTF8.GetBytes("OK`n")
            $res.StatusCode = 200
            $res.ContentType = "text/plain; charset=utf-8"
            $res.ContentLength64 = $body.Length
            $res.OutputStream.Write($body, 0, $body.Length)
            $res.OutputStream.Close()
            Write-Host "$(Get-Date -Format o) HEALTH from $($req.RemoteEndPoint)"
            continue
        }

        if ($req.Url.AbsolutePath -ne "/upload") {
            $body = [Text.Encoding]::UTF8.GetBytes("Not Found`n")
            $res.StatusCode = 404
            $res.ContentType = "text/plain; charset=utf-8"
            $res.ContentLength64 = $body.Length
            $res.OutputStream.Write($body, 0, $body.Length)
            $res.OutputStream.Close()
            continue
        }

        $ts = Get-Date -Format "yyyyMMdd_HHmmss_ffff"
        $ext = ".bin"
        if ($req.ContentType -match "zip") { $ext = ".zip" }

        $outFile = Join-Path $OutDir "upload_$ts$ext"

        $fs = [IO.File]::Create($outFile)
        try {
            $req.InputStream.CopyTo($fs)
        }
        finally {
            $fs.Close()
        }

        $item = Get-Item $outFile
        $sha = (Get-FileHash $outFile -Algorithm SHA256).Hash.ToLowerInvariant()

        $entry = [ordered]@{
            timestamp = (Get-Date).ToString("o")
            client_ip = $req.RemoteEndPoint.Address.ToString()
            method = $req.HttpMethod
            path = $req.Url.AbsolutePath
            bytes_received = $item.Length
            saved_file = $item.FullName
            sha256 = $sha
            user_agent = $req.UserAgent
            content_type = $req.ContentType
        }

        ($entry | ConvertTo-Json -Compress) | Add-Content -LiteralPath $LogFile -Encoding UTF8

        $uploads++

        $text = "OK sha256=$sha bytes=$($item.Length)`n"
        $body = [Text.Encoding]::UTF8.GetBytes($text)
        $res.StatusCode = 200
        $res.ContentType = "text/plain; charset=utf-8"
        $res.ContentLength64 = $body.Length
        $res.OutputStream.Write($body, 0, $body.Length)
        $res.OutputStream.Close()

        Write-Host "$(Get-Date -Format o) UPLOAD from $($req.RemoteEndPoint) bytes=$($item.Length) sha256=$sha"
    }
    catch {
        $err = [ordered]@{
            timestamp = (Get-Date).ToString("o")
            error = $_.Exception.Message
            stack = $_.ScriptStackTrace
        }
        ($err | ConvertTo-Json -Compress) | Add-Content -LiteralPath $LogFile -Encoding UTF8

        $body = [Text.Encoding]::UTF8.GetBytes("Receiver error`n")
        $res.StatusCode = 500
        $res.ContentType = "text/plain; charset=utf-8"
        $res.ContentLength64 = $body.Length
        $res.OutputStream.Write($body, 0, $body.Length)
        $res.OutputStream.Close()
    }
}

$listener.Stop()
$listener.Close()
Write-Host "TFM PowerShell receiver stopped. uploads=$uploads"