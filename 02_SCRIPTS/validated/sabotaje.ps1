# =====================================================
# TEC-008 - Sabotaje Controlado
# MITRE ATT&CK: T1485 - Data Destruction
# ENTORNO CONTROLADO TFM
# =====================================================

$TargetPath = "C:\Users\seguridad\Desktop\TFM\Pruebas\DUMB_LAB"

Write-Host "=== INICIANDO SABOTAJE CONTROLADO ===" -ForegroundColor Cyan

if (-not (Test-Path $TargetPath)) {
    Write-Host "No existe: $TargetPath" -ForegroundColor Red
    exit
}

# Enumerar archivos
$Files = Get-ChildItem -Path $TargetPath -File -Recurse -ErrorAction SilentlyContinue

Write-Host "Archivos detectados: $($Files.Count)" -ForegroundColor Yellow

foreach ($file in $Files) {

    try {

        Write-Host "BORRANDO: $($file.FullName)" -ForegroundColor Red

        Remove-Item -Path $file.FullName -Force

    }
    catch {

        Write-Host "ERROR: $($file.FullName)" -ForegroundColor DarkRed
    }
}

Write-Host ""
Write-Host "TEC-008 COMPLETADA" -ForegroundColor Green