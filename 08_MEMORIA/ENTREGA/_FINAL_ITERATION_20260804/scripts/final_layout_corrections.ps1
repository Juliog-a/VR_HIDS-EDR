param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [Parameter(Mandatory = $true)]
    [string]$PdfPath,
    [Parameter(Mandatory = $true)]
    [string]$OutJson
)

$ErrorActionPreference = 'Stop'

$wdAlertsNone = 0
$wdDoNotSaveChanges = 0
$wdCollapseEnd = 0
$wdFindStop = 0
$wdReplaceAll = 2
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleCaption = -35
$wdLineSpaceSingle = 0
$wdLineSpaceExactly = 4
$wdExportFormatPDF = 17
$wdExportOptimizeForPrint = 0
$wdExportAllDocument = 0
$wdExportDocumentContent = 0
$wdExportCreateHeadingBookmarks = 1

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Clean-Text {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace '[\r\a\v\f]', ' ') -replace '\s+', ' ').Trim()
}

function Find-Paragraph {
    param(
        [object]$Document,
        [string]$Needle,
        [switch]$Exact
    )
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = Clean-Text ([string]$paragraph.Range.Text)
        $match = if ($Exact) { $text -eq $Needle } else { $text.Contains($Needle) }
        if ($match) { return $paragraph }
        Release-ComObject $paragraph
    }
    return $null
}

function Find-HeadingParagraph {
    param(
        [object]$Document,
        [string]$Needle,
        [int]$Level
    )
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $text = Clean-Text ([string]$paragraph.Range.Text)
        $outlineLevel = 10
        try { $outlineLevel = [int]$paragraph.OutlineLevel } catch {}
        if ($outlineLevel -eq $Level -and $text.Contains($Needle)) {
            return $paragraph
        }
        Release-ComObject $paragraph
    }
    return $null
}

function Set-ParagraphText {
    param(
        [object]$Paragraph,
        [string]$Text
    )
    $range = $Paragraph.Range.Duplicate
    try {
        if ($range.End -gt $range.Start) { $range.End = $range.End - 1 }
        $range.Text = $Text
    }
    finally { Release-ComObject $range }
}

function Insert-NormalParagraphAfter {
    param(
        [object]$Document,
        [object]$Paragraph,
        [string]$Text
    )
    $position = [int]$Paragraph.Range.End
    $range = $Document.Range($position, $position)
    try {
        $range.InsertAfter($Text + "`r")
    }
    finally { Release-ComObject $range }
    $inserted = Find-Paragraph -Document $Document -Needle $Text -Exact
    if ($null -ne $inserted) {
        try { $inserted.Range.Style = $Document.Styles.Item($wdStyleNormal) }
        finally { Release-ComObject $inserted }
    }
}

function Replace-AllText {
    param(
        [object]$Document,
        [string]$OldText,
        [string]$NewText
    )
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $OldText
        $find.Replacement.Text = $NewText
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        [void]$find.Execute(
            $OldText,
            $false,
            $false,
            $false,
            $false,
            $false,
            $true,
            $wdFindStop,
            $false,
            $NewText,
            $wdReplaceAll
        )
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Replace-CaptionTitlesBatch {
    param(
        [object]$Document,
        [array]$Replacements
    )

    $pending = @{}
    foreach ($replacement in $Replacements) {
        $oldTitle = [string]$replacement[0]
        if (-not $pending.ContainsKey($oldTitle)) {
            $pending[$oldTitle] = [System.Collections.Generic.Queue[string]]::new()
        }
        $pending[$oldTitle].Enqueue([string]$replacement[1])
    }

    $applied = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        try {
            $text = Clean-Text ([string]$paragraph.Range.Text)
            $match = [regex]::Match($text, '^(Tabla|Figura|Código)\s+\d+:\s*(.+)$')
            if (-not $match.Success) { continue }

            $oldTitle = $match.Groups[2].Value.Trim()
            if (-not $pending.ContainsKey($oldTitle) -or $pending[$oldTitle].Count -eq 0) {
                continue
            }

            $newTitle = $pending[$oldTitle].Dequeue()
            $range = $paragraph.Range.Duplicate
            $find = $range.Find
            try {
                $find.ClearFormatting()
                $find.Text = $oldTitle
                $find.Replacement.Text = $newTitle
                $find.Wrap = $wdFindStop
                [void]$find.Execute(
                    $oldTitle,
                    $false,
                    $false,
                    $false,
                    $false,
                    $false,
                    $true,
                    $wdFindStop,
                    $false,
                    $newTitle,
                    1
                )
            }
            finally {
                Release-ComObject $find
                Release-ComObject $range
            }
            $paragraph.Range.Style = $Document.Styles.Item($wdStyleCaption)
            [void]$applied.Add([pscustomobject]@{
                OldTitle = $oldTitle
                NewTitle = $newTitle
            })
        }
        finally {
            Release-ComObject $paragraph
        }
    }
    return @($applied)
}

function Update-AllFields {
    param([object]$Document)
    foreach ($story in @($Document.StoryRanges)) {
        $range = $story
        try {
            while ($null -ne $range) {
                try { [void]$range.Fields.Update() } catch {}
                $next = $range.NextStoryRange
                if ($null -ne $range -and $range -ne $story) { Release-ComObject $range }
                $range = $next
            }
        }
        finally {
            if ($null -ne $range -and $range -ne $story) { Release-ComObject $range }
            Release-ComObject $story
        }
    }
    for ($i = 1; $i -le $Document.TablesOfContents.Count; $i++) {
        $toc = $Document.TablesOfContents.Item($i)
        try { $toc.Update() } finally { Release-ComObject $toc }
    }
    for ($i = 1; $i -le $Document.TablesOfFigures.Count; $i++) {
        $tof = $Document.TablesOfFigures.Item($i)
        try { $tof.Update() } finally { Release-ComObject $tof }
    }
    try { [void]$Document.Fields.Update() } catch {}
}

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$resolvedPdf = [System.IO.Path]::GetFullPath($PdfPath)
$changes = [System.Collections.Generic.List[object]]::new()
$word = $null
$doc = $null
$pagesBefore = 0
$pagesAfter = 0
$reopenedPages = 0

try {
    Write-Output 'ETAPA 1/6: apertura del DOCX'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = $wdAlertsNone
    $word.ScreenUpdating = $false
    $doc = $word.Documents.Open($resolvedDocx, $false, $false)
    $pagesBefore = [int]$doc.ComputeStatistics(2)

    # Elimina el encabezado vacío que introducía un "Capítulo I:" residual y
    # desplazaba en +1 toda la numeración de capítulos.
    $introduction = Find-Paragraph -Document $doc -Needle 'Introducción' -Exact
    if ($null -ne $introduction) {
        $introStart = [int]$introduction.Range.Start
        for ($i = $doc.Paragraphs.Count; $i -ge 1; $i--) {
            $paragraph = $doc.Paragraphs.Item($i)
            $text = Clean-Text ([string]$paragraph.Range.Text)
            $isBeforeIntroduction = [int]$paragraph.Range.Start -lt $introStart
            $isHeadingOne = $false
            try { $isHeadingOne = ([int]$paragraph.OutlineLevel -eq 1) } catch {}
            if ($isBeforeIntroduction -and $isHeadingOne -and [string]::IsNullOrWhiteSpace($text)) {
                $paragraph.Range.Delete()
                [void]$changes.Add([pscustomobject]@{Category='Jerarquía';Description='Eliminado encabezado vacío Capítulo I'})
                Release-ComObject $paragraph
                break
            }
            Release-ComObject $paragraph
        }
        Release-ComObject $introduction
    }

    $futureSentence = 'Las evoluciones futuras de los artifacts deberán conservar un ciclo de vida trazable: creación como versión de trabajo, validación experimental frente al ground truth, revisión de ruido y promoción a versión validada solo cuando el resultado sea reproducible'
    $futureHeading = Find-HeadingParagraph -Document $doc -Needle 'Las evoluciones futuras de los artifacts deberán conservar un ciclo de vida trazable' -Level 2
    if ($null -ne $futureHeading) {
        Set-ParagraphText -Paragraph $futureHeading -Text 'Ciclo de vida y trazabilidad de los artifacts'
        $futureHeading.Range.Style = $doc.Styles.Item($wdStyleHeading2)
        Insert-NormalParagraphAfter -Document $doc -Paragraph $futureHeading -Text $futureSentence
        [void]$changes.Add([pscustomobject]@{Category='Jerarquía';Description='Convertido el texto narrativo 10.4 en título descriptivo y párrafo normal'})
        Release-ComObject $futureHeading
    }

    $referenceNarrative = 'Las referencias se han normalizado en formato APA y se separan entre fuentes externas oficiales y fuentes internas del TFM. Las fuentes externas sustentan conceptos, taxonomías y documentación de herramientas. Las fuentes internas sustentan resultados experimentales, decisiones metodológicas, datasets, artifacts y evidencias generadas durante el laboratorio.'
    $referenceTitleStandalone = Find-Paragraph -Document $doc -Needle 'Referencias bibliográficas' -Exact
    $referenceHeading = Find-HeadingParagraph -Document $doc -Needle 'Las referencias se han normalizado en formato APA' -Level 1
    if ($null -ne $referenceTitleStandalone -and $null -ne $referenceHeading) {
        if ([int]$referenceTitleStandalone.Range.Start -lt [int]$referenceHeading.Range.Start) {
            $referenceTitleStandalone.Range.Delete()
            [void]$changes.Add([pscustomobject]@{Category='Jerarquía';Description='Eliminado título duplicado previo a referencias'})
        }
    }
    Release-ComObject $referenceTitleStandalone
    if ($null -ne $referenceHeading) {
        Set-ParagraphText -Paragraph $referenceHeading -Text 'Referencias bibliográficas y anexos'
        $referenceHeading.Range.Style = $doc.Styles.Item($wdStyleHeading1)
        Insert-NormalParagraphAfter -Document $doc -Paragraph $referenceHeading -Text $referenceNarrative
        [void]$changes.Add([pscustomobject]@{Category='Jerarquía';Description='Corregido título del capítulo de referencias y anexos'})
        Release-ComObject $referenceHeading
    }

    $headingBenchmark = Find-HeadingParagraph -Document $doc -Needle 'Evaluación de Velociraptor y Benchmark de Rendimiento' -Level 1
    if ($null -ne $headingBenchmark) {
        Set-ParagraphText -Paragraph $headingBenchmark -Text 'Evaluación de Velociraptor y benchmark de rendimiento'
        $headingBenchmark.Range.Style = $doc.Styles.Item($wdStyleHeading1)
        [void]$changes.Add([pscustomobject]@{Category='Jerarquía';Description='Normalizada capitalización del capítulo de benchmark'})
        Release-ComObject $headingBenchmark
    }

    # Neutraliza referencias internas sin exponer árboles locales antiguos.
    $textReplacements = @(
        @('01_ARTIFACTS\validated', '<RUTA_ARTIFACTS>'),
        @('00_CONTEXT/DECISIONS.md; 00_CONTEXT/TEST_MATRIX.md; 00_CONTEXT/EVIDENCE_INDEX.md', '<RAÍZ_TFM>\00_CONTEXT\DECISIONS.md; <RAÍZ_TFM>\00_CONTEXT\TEST_MATRIX.md; <RAÍZ_TFM>\00_CONTEXT\EVIDENCE_INDEX.md'),
        @('10_WAZUH\tfm_wazuh_custom_rules_v2.xml', '<RUTA_ARTIFACTS>\tfm_wazuh_custom_rules_v2.xml')
    )
    foreach ($replacement in $textReplacements) {
        Replace-AllText -Document $doc -OldText $replacement[0] -NewText $replacement[1]
        [void]$changes.Add([pscustomobject]@{Category='Rutas';Description=("Neutralizada referencia interna: " + $replacement[0])})
    }

    $captionReplacements = @(
        @('Tipos de artifacts', 'Tipos de artifacts según su modo de ejecución'),
        @('Planificación Cap.1', 'Planificación del capítulo I'),
        @('Planificación Cap.2', 'Planificación del capítulo II'),
        @('Planificación Cap.3', 'Planificación del capítulo III'),
        @('Planificación Cap.4', 'Planificación del capítulo IV'),
        @('Planificación Cap.5', 'Planificación del capítulo V'),
        @('Planificación Cap.6', 'Planificación del capítulo VI'),
        @('Planificación Cap.7', 'Planificación del capítulo VII'),
        @('Planificación Cap.8', 'Planificación del capítulo VIII'),
        @('Planificación Cap.9', 'Planificación del capítulo IX'),
        @('Planificación Cap.10', 'Planificación del capítulo X'),
        @('Diagrama Gantt General', 'Diagrama de Gantt general'),
        @('Diagrama Gantt Cap.1 y 2', 'Planificación Gantt de los capítulos I y II'),
        @('Diagrama Gantt Cap.3', 'Planificación Gantt del capítulo III'),
        @('Diagrama Gantt Cap. 4', 'Planificación Gantt del capítulo IV'),
        @('Diagrama Gantt Cap.5', 'Planificación Gantt del capítulo V'),
        @('Diagrama Gantt Cap.6', 'Planificación Gantt del capítulo VI'),
        @('Diagrama Gantt Cap.7', 'Planificación Gantt del capítulo VII'),
        @('Diagrama Gantt Cap.8 y 9', 'Planificación Gantt de los capítulos VIII y IX'),
        @('FUNCIONES, LIMITACIONES Y RIESGOS DEL ENFOQUE HIDS', 'Funciones, limitaciones y riesgos del enfoque HIDS'),
        @('FUNCIONES, LIMITACIONES Y RIESGOS DEL ENFOQUE DFIR', 'Funciones, limitaciones y riesgos del enfoque DFIR'),
        @('FUNCIONES, LIMITACIONES Y RIESGOS DEL ENFOQUE EDR', 'Funciones, limitaciones y riesgos del enfoque EDR'),
        @('Capa de componentes VR', 'Capas funcionales de la arquitectura de Velociraptor'),
        @('Envío de datos VR', 'Flujo de datos de Velociraptor'),
        @('Campos Artifacts', 'Campos principales de definición de artifacts'),
        @('Tipos de artifacts', 'Tipos de artifacts según el contexto de ejecución'),
        @('Ejemplo consulta VQL', 'Ejemplo de consulta VQL'),
        @('Ejemplo obtener info VQL', 'Ejemplo de obtención de información mediante VQL'),
        @('Ejemplo consulta de claves VQL', 'Ejemplo de consulta de claves del registro mediante VQL'),
        @('Ejemplo funcionamiento lazy evaluation', 'Ejemplo de evaluación diferida en VQL'),
        @('Parámetros iniciales VM', 'Parámetros iniciales de la máquina virtual'),
        @('Instalador de agente', 'Instalación del agente Velociraptor'),
        @('Datos provenientes de los agentes', 'Datos recibidos desde los agentes Velociraptor'),
        @('Directorio máquina física', 'Estructura de directorios de la máquina física'),
        @('arquitectura de alertas', 'Arquitectura del flujo de alertas'),
        @('Escenarios de ataque', 'Escenarios de ataque evaluados'),
        @('Ejecución TEC-001', 'Procedimiento de ejecución de TEC-001'),
        @('TEC-001 ejecución', 'Evidencia de ejecución de TEC-001'),
        @('Ejecución TEC-002', 'Procedimiento de ejecución de TEC-002'),
        @('TEC002_OK', 'Evidencia de ejecución de TEC-002'),
        @('Ejecución TEC-003', 'Procedimiento de ejecución de TEC-003'),
        @('Scheduled Task runner', 'Evidencia de la tarea programada TEC-003'),
        @('Ejecución TEC-004', 'Procedimiento de ejecución de TEC-004'),
        @('Registry Run Key ejecución', 'Evidencia de persistencia mediante clave Run en TEC-004'),
        @('Ejecución TEC-005', 'Procedimiento de ejecución de TEC-005'),
        @('Servicio ejecución', 'Evidencia de ejecución del servicio TEC-005'),
        @('Ejecución TEC-006-1', 'Procedimiento de ejecución de TEC-006 mediante PowerShell'),
        @('Security Software Discovery ejecución', 'Evidencia de descubrimiento de software de seguridad en TEC-006'),
        @('Ejecución TEC-006-2', 'Procedimiento de ejecución de TEC-006 mediante WMIC'),
        @('Ejecución TEC-007', 'Procedimiento de ejecución de TEC-007'),
        @('Script Ransom ejecutado', 'Evidencia de cifrado controlado en TEC-007'),
        @('Ejecución TEC-008', 'Procedimiento de ejecución de TEC-008'),
        @('Receiver', 'Ejecución del receptor HTTP controlado'),
        @('Script receiver', 'Evidencia del receptor HTTP controlado'),
        @('Ejecución TEC-009', 'Procedimiento de ejecución de TEC-009'),
        @('Script TEC-009 ejecutada', 'Evidencia de ejecución de TEC-009'),
        @('Evidencias VR', 'Evidencias obtenidas en Velociraptor'),
        @('Wazuh dashboard', 'Panel de acceso al dashboard de Wazuh')
    )
    Write-Output 'ETAPA 2/6: corrección de captions en un único recorrido'
    $captionChanges = Replace-CaptionTitlesBatch -Document $doc -Replacements $captionReplacements
    foreach ($captionChange in $captionChanges) {
        [void]$changes.Add([pscustomobject]@{
            Category='Captions'
            Description=($captionChange.OldTitle + ' -> ' + $captionChange.NewTitle)
        })
    }

    # Compacta el resumen sin reducirlo a un tamaño ilegible, para evitar que
    # las palabras clave queden aisladas en una página.
    $summaryHeading = Find-Paragraph -Document $doc -Needle 'Resumen' -Exact
    $keywords = Find-Paragraph -Document $doc -Needle 'Palabras clave:'
    if ($null -ne $summaryHeading -and $null -ne $keywords) {
        $start = [int]$summaryHeading.Range.End
        $end = [int]$keywords.Range.End
        $summaryRange = $doc.Range($start, $end)
        try {
            $summaryRange.Font.Size = 9.5
            $summaryRange.ParagraphFormat.SpaceBefore = 0
            $summaryRange.ParagraphFormat.SpaceAfter = 0
            $summaryRange.ParagraphFormat.LineSpacingRule = $wdLineSpaceExactly
            $summaryRange.ParagraphFormat.LineSpacing = 11
            $summaryRange.ParagraphFormat.WidowControl = $true
        }
        finally { Release-ComObject $summaryRange }
        [void]$changes.Add([pscustomobject]@{Category='Maquetación';Description='Compactado el resumen para mantener sus palabras clave en la misma página'})
    }
    Release-ComObject $summaryHeading
    Release-ComObject $keywords

    $endpointParagraph = Find-Paragraph -Document $doc -Needle 'Para evaluar Velociraptor como HIDS'
    if ($null -ne $endpointParagraph) {
        $endpointParagraph.Range.Font.Size = 9.5
        $endpointParagraph.Format.SpaceBefore = 0
        $endpointParagraph.Format.SpaceAfter = 0
        $endpointParagraph.Format.LineSpacingRule = $wdLineSpaceExactly
        $endpointParagraph.Format.LineSpacing = 10.5
        [void]$changes.Add([pscustomobject]@{Category='Maquetación';Description='Corregida continuación aislada en la sección de artifacts'})
        Release-ComObject $endpointParagraph
    }

    $annexM = Find-HeadingParagraph -Document $doc -Needle 'Anexo M.' -Level 2
    if ($null -ne $annexM) {
        Set-ParagraphText -Paragraph $annexM -Text 'Anexo M. Declaración responsable sobre el uso de inteligencia artificial'
        $annexM.Range.Style = $doc.Styles.Item($wdStyleHeading2)
        $annexRange = $doc.Range([int]$annexM.Range.End, [int]$doc.Content.End)
        try {
            $annexRange.Font.Size = 9.5
            $annexRange.ParagraphFormat.SpaceBefore = 0
            $annexRange.ParagraphFormat.SpaceAfter = 0
            $annexRange.ParagraphFormat.LineSpacingRule = $wdLineSpaceExactly
            $annexRange.ParagraphFormat.LineSpacing = 10.8
            $annexRange.ParagraphFormat.WidowControl = $true
        }
        finally { Release-ComObject $annexRange }
        [void]$changes.Add([pscustomobject]@{Category='Maquetación';Description='Normalizado y compactado el Anexo M para evitar una última página residual'})
        Release-ComObject $annexM
    }

    Write-Output 'ETAPA 3/6: actualización de campos e índices'
    Update-AllFields -Document $doc
    $doc.Repaginate()
    Update-AllFields -Document $doc
    $doc.Repaginate()
    $pagesAfter = [int]$doc.ComputeStatistics(2)
    Write-Output 'ETAPA 4/6: guardado del DOCX'
    $doc.Save()
    $doc.Close($wdDoNotSaveChanges)
    Release-ComObject $doc
    $doc = $null

    # Reapertura de estabilidad y exportación exacta desde el DOCX guardado.
    Write-Output 'ETAPA 5/6: reapertura de estabilidad'
    $doc = $word.Documents.Open($resolvedDocx, $false, $true)
    $doc.Repaginate()
    $reopenedPages = [int]$doc.ComputeStatistics(2)
    $doc.ExportAsFixedFormat(
        $resolvedPdf,
        $wdExportFormatPDF,
        $false,
        $wdExportOptimizeForPrint,
        $wdExportAllDocument,
        1,
        1,
        $wdExportDocumentContent,
        $true,
        $true,
        $wdExportCreateHeadingBookmarks,
        $true,
        $true,
        $false
    )
    Write-Output 'ETAPA 6/6: PDF exportado'
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $doc
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$report = [pscustomobject]@{
    DocxPath = $resolvedDocx
    PdfPath = $resolvedPdf
    PagesBefore = $pagesBefore
    PagesAfter = $pagesAfter
    ReopenedPages = $reopenedPages
    Corrections = $changes.Count
    Changes = @($changes)
    DocxSHA256 = (Get-FileHash -LiteralPath $resolvedDocx -Algorithm SHA256).Hash
    PdfSHA256 = (Get-FileHash -LiteralPath $resolvedPdf -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 7
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutJson),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
