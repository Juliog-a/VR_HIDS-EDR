param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-EntryText {
    param([System.IO.Compression.ZipArchiveEntry]$Entry)
    $reader = [System.IO.StreamReader]::new($Entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-XmlEntry {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Name,
        [xml]$Xml
    )
    $old = $Zip.GetEntry($Name)
    if ($null -ne $old) { $old.Delete() }
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try { $Xml.Save($writer) } finally { $writer.Dispose(); $stream.Dispose() }
}

function Get-ParagraphText {
    param([System.Xml.XmlNode]$Paragraph)
    $builder = [System.Text.StringBuilder]::new()
    foreach ($node in @($Paragraph.SelectNodes('.//w:t', $script:ns))) {
        [void]$builder.Append([string]$node.InnerText)
    }
    return (($builder.ToString() -replace '\s+', ' ').Trim())
}

function Get-ParagraphStyle {
    param([System.Xml.XmlNode]$Paragraph)
    $style = $Paragraph.SelectSingleNode('./w:pPr/w:pStyle', $script:ns)
    if ($null -eq $style) { return '' }
    return [string]$style.GetAttribute('val', $script:w)
}

function Set-Attribute {
    param(
        [System.Xml.XmlElement]$Element,
        [string]$Name,
        [string]$Value
    )
    [void]$Element.SetAttribute($Name, $script:w, $Value)
}

function Ensure-Child {
    param(
        [xml]$Xml,
        [System.Xml.XmlNode]$Parent,
        [string]$LocalName,
        [string]$BeforeLocalName = ''
    )
    $child = $Parent.SelectSingleNode('./w:' + $LocalName, $script:ns)
    if ($null -ne $child) { return $child }
    $child = $Xml.CreateElement('w', $LocalName, $script:w)
    if ($BeforeLocalName) {
        $before = $Parent.SelectSingleNode('./w:' + $BeforeLocalName, $script:ns)
        if ($null -ne $before) {
            [void]$Parent.InsertBefore($child, $before)
            return $child
        }
    }
    [void]$Parent.AppendChild($child)
    return $child
}

function Set-ParagraphText {
    param(
        [xml]$Xml,
        [System.Xml.XmlNode]$Paragraph,
        [string]$Text
    )
    foreach ($child in @($Paragraph.ChildNodes)) {
        if ($child.LocalName -ne 'pPr') {
            [void]$Paragraph.RemoveChild($child)
        }
    }
    $run = $Xml.CreateElement('w', 'r', $script:w)
    $textNode = $Xml.CreateElement('w', 't', $script:w)
    $textNode.InnerText = $Text
    [void]$run.AppendChild($textNode)
    [void]$Paragraph.AppendChild($run)
}

function New-NormalParagraph {
    param(
        [xml]$Xml,
        [string]$Text
    )
    $paragraph = $Xml.CreateElement('w', 'p', $script:w)
    $pPr = $Xml.CreateElement('w', 'pPr', $script:w)
    $style = $Xml.CreateElement('w', 'pStyle', $script:w)
    Set-Attribute -Element $style -Name 'val' -Value 'Normal'
    [void]$pPr.AppendChild($style)
    [void]$paragraph.AppendChild($pPr)
    $run = $Xml.CreateElement('w', 'r', $script:w)
    $textNode = $Xml.CreateElement('w', 't', $script:w)
    $textNode.InnerText = $Text
    [void]$run.AppendChild($textNode)
    [void]$paragraph.AppendChild($run)
    return $paragraph
}

function Set-CompactFormatting {
    param(
        [xml]$Xml,
        [System.Xml.XmlNode]$Paragraph,
        [int]$FontHalfPoints,
        [int]$LineTwips
    )
    $pPr = $Paragraph.SelectSingleNode('./w:pPr', $script:ns)
    if ($null -eq $pPr) {
        $pPr = $Xml.CreateElement('w', 'pPr', $script:w)
        if ($Paragraph.HasChildNodes) {
            [void]$Paragraph.InsertBefore($pPr, $Paragraph.FirstChild)
        }
        else {
            [void]$Paragraph.AppendChild($pPr)
        }
    }
    $spacing = Ensure-Child -Xml $Xml -Parent $pPr -LocalName 'spacing'
    Set-Attribute -Element $spacing -Name 'before' -Value '0'
    Set-Attribute -Element $spacing -Name 'after' -Value '0'
    Set-Attribute -Element $spacing -Name 'line' -Value ([string]$LineTwips)
    Set-Attribute -Element $spacing -Name 'lineRule' -Value 'exact'
    $widow = Ensure-Child -Xml $Xml -Parent $pPr -LocalName 'widowControl'
    Set-Attribute -Element $widow -Name 'val' -Value '1'

    foreach ($run in @($Paragraph.SelectNodes('.//w:r', $script:ns))) {
        $rPr = $run.SelectSingleNode('./w:rPr', $script:ns)
        if ($null -eq $rPr) {
            $rPr = $Xml.CreateElement('w', 'rPr', $script:w)
            if ($run.HasChildNodes) {
                [void]$run.InsertBefore($rPr, $run.FirstChild)
            }
            else {
                [void]$run.AppendChild($rPr)
            }
        }
        $size = Ensure-Child -Xml $Xml -Parent $rPr -LocalName 'sz'
        Set-Attribute -Element $size -Name 'val' -Value ([string]$FontHalfPoints)
        $sizeCs = Ensure-Child -Xml $Xml -Parent $rPr -LocalName 'szCs'
        Set-Attribute -Element $sizeCs -Name 'val' -Value ([string]$FontHalfPoints)
    }
}

$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path
$changes = [System.Collections.Generic.List[object]]::new()
$captionMissing = [System.Collections.Generic.List[string]]::new()
$script:w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$zip = [System.IO.Compression.ZipFile]::Open(
    $resolvedDocument,
    [System.IO.Compression.ZipArchiveMode]::Update
)
try {
    $documentEntry = $zip.GetEntry('word/document.xml')
    $settingsEntry = $zip.GetEntry('word/settings.xml')
    if ($null -eq $documentEntry -or $null -eq $settingsEntry) {
        throw 'El DOCX no contiene document.xml o settings.xml.'
    }
    [xml]$documentXml = Read-EntryText $documentEntry
    [xml]$settingsXml = Read-EntryText $settingsEntry
    $script:ns = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
    $script:ns.AddNamespace('w', $script:w)

    $paragraphs = @($documentXml.SelectNodes('//w:body//w:p', $script:ns))
    $introductionIndex = -1
    for ($i = 0; $i -lt $paragraphs.Count; $i++) {
        if ((Get-ParagraphText $paragraphs[$i]) -eq 'Introducción' -and
            (Get-ParagraphStyle $paragraphs[$i]) -eq 'Ttulo1') {
            $introductionIndex = $i
            break
        }
    }
    if ($introductionIndex -ge 0) {
        for ($i = $introductionIndex - 1; $i -ge 0; $i--) {
            if ((Get-ParagraphStyle $paragraphs[$i]) -eq 'Ttulo1' -and
                [string]::IsNullOrWhiteSpace((Get-ParagraphText $paragraphs[$i]))) {
                [void]$paragraphs[$i].ParentNode.RemoveChild($paragraphs[$i])
                [void]$changes.Add([pscustomobject]@{
                    Category='Jerarquía'
                    Description='Eliminado encabezado vacío que generaba Capítulo I residual'
                })
                break
            }
        }
    }

    $futureSentence = 'Las evoluciones futuras de los artifacts deberán conservar un ciclo de vida trazable: creación como versión de trabajo, validación experimental frente al ground truth, revisión de ruido y promoción a versión validada solo cuando el resultado sea reproducible'
    $referenceNarrative = 'Las referencias se han normalizado en formato APA y se separan entre fuentes externas oficiales y fuentes internas del TFM. Las fuentes externas sustentan conceptos, taxonomías y documentación de herramientas. Las fuentes internas sustentan resultados experimentales, decisiones metodológicas, datasets, artifacts y evidencias generadas durante el laboratorio.'

    $paragraphs = @($documentXml.SelectNodes('//w:body//w:p', $script:ns))
    $futureHeading = $null
    $referenceHeading = $null
    $referenceStandalone = $null
    foreach ($paragraph in $paragraphs) {
        $text = Get-ParagraphText $paragraph
        $style = Get-ParagraphStyle $paragraph
        if ($null -eq $futureHeading -and $style -eq 'Ttulo2' -and
            $text.Contains('Las evoluciones futuras de los artifacts deberán conservar un ciclo de vida trazable')) {
            $futureHeading = $paragraph
        }
        if ($null -eq $referenceHeading -and $style -eq 'Ttulo1' -and
            $text.Contains('Las referencias se han normalizado en formato APA')) {
            $referenceHeading = $paragraph
        }
        if ($null -eq $referenceStandalone -and $text -eq 'Referencias bibliográficas') {
            $referenceStandalone = $paragraph
        }
        if ($style -eq 'Ttulo1' -and
            $text -eq 'Evaluación de Velociraptor y Benchmark de Rendimiento') {
            Set-ParagraphText -Xml $documentXml -Paragraph $paragraph -Text 'Evaluación de Velociraptor y benchmark de rendimiento'
            [void]$changes.Add([pscustomobject]@{
                Category='Jerarquía'
                Description='Normalizada la capitalización del capítulo de benchmark'
            })
        }
    }

    if ($null -ne $futureHeading) {
        Set-ParagraphText -Xml $documentXml -Paragraph $futureHeading -Text 'Ciclo de vida y trazabilidad de los artifacts'
        $normal = New-NormalParagraph -Xml $documentXml -Text $futureSentence
        [void]$futureHeading.ParentNode.InsertAfter($normal, $futureHeading)
        [void]$changes.Add([pscustomobject]@{
            Category='Jerarquía'
            Description='Separado el título 10.4 de su contenido narrativo'
        })
    }

    if ($null -ne $referenceHeading) {
        if ($null -ne $referenceStandalone) {
            [void]$referenceStandalone.ParentNode.RemoveChild($referenceStandalone)
            [void]$changes.Add([pscustomobject]@{
                Category='Jerarquía'
                Description='Eliminado el título duplicado previo a las referencias'
            })
        }
        Set-ParagraphText -Xml $documentXml -Paragraph $referenceHeading -Text 'Referencias bibliográficas y anexos'
        $normal = New-NormalParagraph -Xml $documentXml -Text $referenceNarrative
        [void]$referenceHeading.ParentNode.InsertAfter($normal, $referenceHeading)
        [void]$changes.Add([pscustomobject]@{
            Category='Jerarquía'
            Description='Corregido el título del capítulo de referencias y anexos'
        })
    }

    $textReplacements = @(
        @('01_ARTIFACTS\validated', '<RUTA_ARTIFACTS>'),
        @('00_CONTEXT/DECISIONS.md; 00_CONTEXT/TEST_MATRIX.md; 00_CONTEXT/EVIDENCE_INDEX.md', '<RAÍZ_TFM>\00_CONTEXT\DECISIONS.md; <RAÍZ_TFM>\00_CONTEXT\TEST_MATRIX.md; <RAÍZ_TFM>\00_CONTEXT\EVIDENCE_INDEX.md'),
        @('10_WAZUH\tfm_wazuh_custom_rules_v2.xml', '<RUTA_ARTIFACTS>\tfm_wazuh_custom_rules_v2.xml')
    )
    foreach ($replacement in $textReplacements) {
        $replacementCount = 0
        foreach ($textNode in @($documentXml.SelectNodes('//w:t', $script:ns))) {
            if ([string]$textNode.InnerText -like ('*' + $replacement[0] + '*')) {
                $textNode.InnerText = ([string]$textNode.InnerText).Replace($replacement[0], $replacement[1])
                $replacementCount++
            }
        }
        if ($replacementCount -gt 0) {
            [void]$changes.Add([pscustomobject]@{
                Category='Rutas'
                Description=('Neutralizada referencia interna: ' + $replacement[0])
                Count=$replacementCount
            })
        }
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

    $pending = @{}
    foreach ($replacement in $captionReplacements) {
        if (-not $pending.ContainsKey($replacement[0])) {
            $pending[$replacement[0]] = [System.Collections.Generic.Queue[string]]::new()
        }
        $pending[$replacement[0]].Enqueue([string]$replacement[1])
    }
    foreach ($paragraph in @($documentXml.SelectNodes('//w:p[w:pPr/w:pStyle[@w:val="Descripcin"]]', $script:ns))) {
        $paragraphText = Get-ParagraphText $paragraph
        $match = [regex]::Match($paragraphText, '^(Tabla|Figura|Código)\s+\d+:\s*(.+)$')
        if (-not $match.Success) { continue }
        $oldTitle = $match.Groups[2].Value.Trim()
        if (-not $pending.ContainsKey($oldTitle) -or $pending[$oldTitle].Count -eq 0) { continue }
        $newTitle = $pending[$oldTitle].Dequeue()
        $replaced = $false
        $field = $paragraph.SelectSingleNode('./w:fldSimple[contains(@w:instr, "SEQ ")]', $script:ns)
        if ($null -ne $field) {
            foreach ($child in @($paragraph.ChildNodes)) {
                if ($child.LocalName -notin @('pPr', 'bookmarkStart', 'bookmarkEnd', 'fldSimple')) {
                    [void]$paragraph.RemoveChild($child)
                }
            }
            $prefixRun = $documentXml.CreateElement('w', 'r', $script:w)
            $prefixText = $documentXml.CreateElement('w', 't', $script:w)
            [void]$prefixText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
            $prefixText.InnerText = $match.Groups[1].Value + ' '
            [void]$prefixRun.AppendChild($prefixText)
            [void]$paragraph.InsertBefore($prefixRun, $field)

            $suffixRun = $documentXml.CreateElement('w', 'r', $script:w)
            $suffixText = $documentXml.CreateElement('w', 't', $script:w)
            [void]$suffixText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
            $suffixText.InnerText = ': ' + $newTitle
            [void]$suffixRun.AppendChild($suffixText)
            [void]$paragraph.InsertAfter($suffixRun, $field)
            $replaced = $true
        }
        if ($replaced) {
            [void]$changes.Add([pscustomobject]@{
                Category='Captions'
                Description=($oldTitle + ' -> ' + $newTitle)
            })
        }
        else {
            [void]$captionMissing.Add($oldTitle)
        }
    }
    foreach ($key in $pending.Keys) {
        while ($pending[$key].Count -gt 0) {
            [void]$captionMissing.Add($key)
            [void]$pending[$key].Dequeue()
        }
    }

    $paragraphs = @($documentXml.SelectNodes('//w:body//w:p', $script:ns))
    $summaryIndex = -1
    $keywordsIndex = -1
    $annexMIndex = -1
    for ($i = 0; $i -lt $paragraphs.Count; $i++) {
        $text = Get-ParagraphText $paragraphs[$i]
        $style = Get-ParagraphStyle $paragraphs[$i]
        if ($summaryIndex -lt 0 -and $text -eq 'Resumen') { $summaryIndex = $i }
        if ($summaryIndex -ge 0 -and $keywordsIndex -lt 0 -and $text.Contains('Palabras clave:')) {
            $keywordsIndex = $i
        }
        if ($text.Contains('Para evaluar Velociraptor como HIDS')) {
            Set-CompactFormatting -Xml $documentXml -Paragraph $paragraphs[$i] -FontHalfPoints 19 -LineTwips 210
            [void]$changes.Add([pscustomobject]@{
                Category='Maquetación'
                Description='Compactada la continuación aislada de la sección de artifacts'
            })
        }
        if ($annexMIndex -lt 0 -and $style -eq 'Ttulo2' -and $text.StartsWith('Anexo M.')) {
            $annexMIndex = $i
            Set-ParagraphText -Xml $documentXml -Paragraph $paragraphs[$i] -Text 'Anexo M. Declaración responsable sobre el uso de inteligencia artificial'
        }
    }
    if ($summaryIndex -ge 0 -and $keywordsIndex -ge $summaryIndex) {
        for ($i = $summaryIndex + 1; $i -le $keywordsIndex; $i++) {
            Set-CompactFormatting -Xml $documentXml -Paragraph $paragraphs[$i] -FontHalfPoints 19 -LineTwips 220
        }
        [void]$changes.Add([pscustomobject]@{
            Category='Maquetación'
            Description='Compactado el resumen para mantener sus palabras clave en la misma página'
        })
    }
    if ($annexMIndex -ge 0) {
        for ($i = $annexMIndex + 1; $i -lt $paragraphs.Count; $i++) {
            Set-CompactFormatting -Xml $documentXml -Paragraph $paragraphs[$i] -FontHalfPoints 19 -LineTwips 216
        }
        [void]$changes.Add([pscustomobject]@{
            Category='Maquetación'
            Description='Normalizado y compactado el Anexo M para evitar una última página residual'
        })
    }

    $settingsNs = [System.Xml.XmlNamespaceManager]::new($settingsXml.NameTable)
    $settingsNs.AddNamespace('w', $script:w)
    $updateFields = $settingsXml.SelectSingleNode('/w:settings/w:updateFields', $settingsNs)
    if ($null -eq $updateFields) {
        $updateFields = $settingsXml.CreateElement('w', 'updateFields', $script:w)
        [void]$settingsXml.DocumentElement.AppendChild($updateFields)
    }
    [void]$updateFields.SetAttribute('val', $script:w, 'true')

    Write-XmlEntry -Zip $zip -Name 'word/document.xml' -Xml $documentXml
    Write-XmlEntry -Zip $zip -Name 'word/settings.xml' -Xml $settingsXml
}
finally {
    $zip.Dispose()
}

$report = [pscustomobject]@{
    DocumentPath=$resolvedDocument
    Corrections=$changes.Count
    CaptionReplacements=@($changes | Where-Object Category -eq 'Captions').Count
    CaptionReplacementsMissing=@($captionMissing)
    Changes=@($changes)
    UpdateFieldsOnOpen=$true
    SHA256=(Get-FileHash -LiteralPath $resolvedDocument -Algorithm SHA256).Hash
}
$json = $report | ConvertTo-Json -Depth 7
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($LogPath),
    $json,
    [System.Text.UTF8Encoding]::new($false)
)
$json
