$root = 'C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGA\_FINAL_ITERATION_20260804'
$scriptPath = Join-Path $root 'scripts\finalize_word.ps1'
$code = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$task = [scriptblock]::Create($code)
& $task `
    -DocumentPath (Join-Path $root 'TFM_ENTREGA_FINAL_WORK.docx') `
    -LogPath (Join-Path $root 'qa\finalize_word.json') `
    -ProgressPath (Join-Path $root 'qa\finalize_word.progress.txt')
