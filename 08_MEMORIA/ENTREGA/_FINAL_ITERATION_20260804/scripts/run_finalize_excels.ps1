$root = 'C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGA\_FINAL_ITERATION_20260804'
$scriptPath = Join-Path $root 'scripts\finalize_excels.ps1'
$code = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$task = [scriptblock]::Create($code)
& $task `
    -DetectionPath (Join-Path $root 'TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712_WORK.xlsx') `
    -BenchmarkPath (Join-Path $root 'TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712_WORK.xlsx') `
    -LogPath (Join-Path $root 'qa\finalize_excels.json') `
    -ProgressPath (Join-Path $root 'qa\finalize_excels.progress.txt')
