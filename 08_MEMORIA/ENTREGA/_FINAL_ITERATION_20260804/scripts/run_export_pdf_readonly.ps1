$root = 'C:\Users\julio\Desktop\TFM\08_MEMORIA\ENTREGA\_FINAL_ITERATION_20260804'
$scriptPath = Join-Path $root 'scripts\export_pdf_readonly.ps1'
$code = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$task = [scriptblock]::Create($code)
& $task `
    -DocumentPath (Join-Path $root 'TFM_ENTREGA_FINAL_WORK.docx') `
    -PdfPath (Join-Path $root 'TFM_ENTREGA_FINAL_WORK.pdf') `
    -LogPath (Join-Path $root 'qa\export_pdf_readonly.json') `
    -ProgressPath (Join-Path $root 'qa\export_pdf_readonly.progress.txt')
