param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$ProgressPath
)

$ErrorActionPreference = 'Stop'
$wdFindStop = 0
$wdFindContinue = 1
$wdReplaceAll = 2
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdStyleCaption = -35
$wdFieldEmpty = -1
$wdAutoFitWindow = 2
$wdPreferredWidthPercent = 2
$wdRowHeightAuto = 0
$wdFormatDocumentDefault = 16

$script:corrections = 0
$script:changes = [System.Collections.Generic.List[object]]::new()

function Set-Progress {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$stamp`t$Message" | Set-Content -LiteralPath $ProgressPath -Encoding UTF8
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Add-Change {
    param(
        [string]$Category,
        [string]$Description,
        [int]$Count = 1
    )
    $script:corrections += $Count
    [void]$script:changes.Add([pscustomobject]@{
        Category = $Category
        Description = $Description
        Count = $Count
    })
}

function Get-OccurrenceCount {
    param(
        [object]$Document,
        [string]$Text
    )
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    return [regex]::Matches([string]$Document.Content.Text, [regex]::Escape($Text)).Count
}

function Replace-AllLiteral {
    param(
        [object]$Document,
        [string]$OldText,
        [string]$NewText,
        [string]$Category,
        [string]$Description
    )
    $count = Get-OccurrenceCount -Document $Document -Text $OldText
    if ($count -eq 0) { return 0 }
    $range = $null
    $find = $null
    try {
        $range = $Document.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $OldText
        $find.Replacement.Text = $NewText
        $find.Forward = $true
        $find.Wrap = $wdFindContinue
        $find.Format = $false
        $find.MatchCase = $false
        $find.MatchWholeWord = $false
        $find.MatchWildcards = $false
        [void]$find.Execute($OldText, $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, $NewText, $wdReplaceAll)
        Add-Change -Category $Category -Description $Description -Count $count
        return $count
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Find-Paragraph {
    param(
        [object]$Document,
        [string]$Needle,
        [Nullable[int]]$StyleConstant = $null
    )
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        $find.ClearFormatting()
        $find.Text = $Needle
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        $find.Format = $false
        if ($null -ne $StyleConstant) {
            $find.Format = $true
            $find.Style = $Document.Styles.Item([int]$StyleConstant)
        }
        if ($find.Execute()) {
            return $range.Paragraphs.Item(1)
        }
        return $null
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Set-ParagraphByNeedle {
    param(
        [object]$Document,
        [string]$Needle,
        [string]$NewText,
        [string]$Category,
        [string]$Description,
        [Nullable[int]]$StyleConstant = $null,
        [Nullable[int]]$NewStyleConstant = $null
    )
    $paragraph = Find-Paragraph -Document $Document -Needle $Needle -StyleConstant $StyleConstant
    if ($null -eq $paragraph) { return $false }
    try {
        $paragraph.Range.Text = $NewText + "`r"
        if ($null -ne $NewStyleConstant) {
            $paragraph.Range.Style = $Document.Styles.Item([int]$NewStyleConstant)
        }
        Add-Change -Category $Category -Description $Description
        return $true
    }
    finally {
        Release-ComObject $paragraph
    }
}

function Delete-ParagraphByNeedle {
    param(
        [object]$Document,
        [string]$Needle,
        [string]$Category,
        [string]$Description
    )
    $paragraph = Find-Paragraph -Document $Document -Needle $Needle
    if ($null -eq $paragraph) { return $false }
    try {
        [void]$paragraph.Range.Delete()
        Add-Change -Category $Category -Description $Description
        return $true
    }
    finally {
        Release-ComObject $paragraph
    }
}

function Delete-ExactParagraph {
    param(
        [object]$Document,
        [string]$ExactText,
        [string]$Category,
        [string]$Description
    )
    $range = $Document.Content.Duplicate
    $find = $range.Find
    try {
        if ($ExactText -eq '/') {
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Text = '^p/^p'
            $find.Replacement.Text = '^p'
            $find.Wrap = $wdFindStop
            $find.MatchWildcards = $false
            if ($find.Execute('^p/^p', $false, $false, $false, $false, $false, $true, $wdFindStop, $false, '^p', 1)) {
                Add-Change -Category $Category -Description $Description
                return $true
            }
            return $false
        }

        $find.ClearFormatting()
        $find.Text = $ExactText
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        $find.MatchCase = $true
        $find.MatchWholeWord = $false
        while ($find.Execute()) {
            $paragraph = $range.Paragraphs.Item(1)
            try {
                $text = ([string]$paragraph.Range.Text -replace "[`r`a]", '').Trim()
                if ($text -eq $ExactText) {
                    [void]$paragraph.Range.Delete()
                    Add-Change -Category $Category -Description $Description
                    return $true
                }
            }
            finally {
                Release-ComObject $paragraph
            }
            $range.Start = $range.End
            $range.End = $Document.Content.End
            $find = $range.Find
            $find.ClearFormatting()
            $find.Text = $ExactText
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            $find.MatchCase = $true
            $find.MatchWholeWord = $false
        }
        return $false
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
    }
}

function Replace-InStyledParagraph {
    param(
        [object]$Document,
        [string]$Needle,
        [int]$StyleConstant,
        [string]$OldText,
        [string]$NewText,
        [string]$Description
    )
    $paragraph = Find-Paragraph -Document $Document -Needle $Needle -StyleConstant $StyleConstant
    if ($null -eq $paragraph) { return $false }
    $range = $null
    $find = $null
    try {
        $range = $paragraph.Range.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $OldText
        $find.Replacement.Text = $NewText
        $find.Wrap = $wdFindStop
        if ($find.Execute($OldText, $false, $false, $false, $false, $false, $true, $wdFindStop, $false, $NewText, 1)) {
            Add-Change -Category 'Captions' -Description $Description
            return $true
        }
        return $false
    }
    finally {
        Release-ComObject $find
        Release-ComObject $range
        Release-ComObject $paragraph
    }
}

function Rebuild-Caption {
    param(
        [object]$Document,
        [string]$Needle,
        [int]$CurrentStyleConstant,
        [string]$Label,
        [string]$Title,
        [string]$Description
    )
    $paragraph = Find-Paragraph -Document $Document -Needle $Needle -StyleConstant $CurrentStyleConstant
    if ($null -eq $paragraph) { return $false }
    $range = $null
    $fieldRange = $null
    $field = $null
    try {
        $start = [int]$paragraph.Range.Start
        $range = $paragraph.Range.Duplicate
        $range.End = $range.End - 1
        $range.Text = "$Label : $Title"
        $fieldPosition = $start + $Label.Length + 1
        $fieldRange = $Document.Range($fieldPosition, $fieldPosition)
        $field = $Document.Fields.Add($fieldRange, $wdFieldEmpty, "SEQ $Label \* ARABIC", $true)
        $paragraph.Range.Style = $Document.Styles.Item($wdStyleCaption)
        Add-Change -Category 'Captions' -Description $Description
        return $true
    }
    finally {
        Release-ComObject $field
        Release-ComObject $fieldRange
        Release-ComObject $range
        Release-ComObject $paragraph
    }
}

function Set-CellText {
    param(
        [object]$Table,
        [int]$Row,
        [int]$Column,
        [string]$Text
    )
    $cell = $Table.Cell($Row, $Column)
    try { $cell.Range.Text = $Text } finally { Release-ComObject $cell }
}

function Clean-CellText {
    param([string]$Text)
    return ($Text -replace "`r", '' -replace "`a", '').Trim()
}

function Normalize-AcronymTable {
    param(
        [object]$Document,
        [int]$TableIndex
    )
    $table = $Document.Tables.Item($TableIndex)
    try {
        $entries = [System.Collections.Generic.List[object]]::new()
        for ($r = 1; $r -le $table.Rows.Count; $r++) {
            $a = Clean-CellText ([string]$table.Cell($r, 1).Range.Text)
            $d = Clean-CellText ([string]$table.Cell($r, 2).Range.Text)
            if (-not [string]::IsNullOrWhiteSpace($a)) {
                [void]$entries.Add([pscustomobject]@{Acronym=$a;Definition=$d})
            }
        }
        $unique = @($entries | Group-Object Acronym | ForEach-Object { $_.Group | Select-Object -First 1 })
        foreach ($entry in $unique) {
            switch ($entry.Acronym) {
                'MSE' { $entry.Definition = 'Mean Squared Error. Error cuadrático medio.' }
                'HIDS' { $entry.Definition = 'Host-based Intrusion Detection System. Sistema de detección de intrusiones basado en host.' }
                'YAML' { $entry.Definition = "YAML Ain't Markup Language. Formato de serialización legible empleado para definir artifacts y configuraciones." }
            }
            if ($entry.Definition -and $entry.Definition -notmatch '[.!?]$') {
                $entry.Definition += '.'
            }
        }
        $sorted = @($unique | Sort-Object Acronym)
        while ($table.Rows.Count -gt $sorted.Count) {
            $table.Rows.Item($table.Rows.Count).Delete()
        }
        while ($table.Rows.Count -lt $sorted.Count) {
            [void]$table.Rows.Add()
        }
        for ($r = 1; $r -le $sorted.Count; $r++) {
            Set-CellText -Table $table -Row $r -Column 1 -Text $sorted[$r-1].Acronym
            Set-CellText -Table $table -Row $r -Column 2 -Text $sorted[$r-1].Definition
        }
        $table.Range.Font.Size = 9.5
        $table.TopPadding = 2.5
        $table.BottomPadding = 2.5
        $table.LeftPadding = 3
        $table.RightPadding = 3
        $table.AllowAutoFit = $true
        $table.AutoFitBehavior($wdAutoFitWindow)
        Add-Change -Category 'Acrónimos' -Description 'Eliminación del duplicado YAML, orden alfabético y normalización de definiciones' -Count ($entries.Count + 3)
    }
    finally {
        Release-ComObject $table
    }
}

function Format-TableForDelivery {
    param(
        [object]$Table,
        [double]$FontSize
    )
    try {
        $Table.AllowAutoFit = $true
        $Table.AutoFitBehavior($wdAutoFitWindow)
        $Table.PreferredWidthType = $wdPreferredWidthPercent
        $Table.PreferredWidth = 100
        $Table.Range.Font.Size = $FontSize
        $Table.TopPadding = 2.5
        $Table.BottomPadding = 2.5
        $Table.LeftPadding = 3
        $Table.RightPadding = 3
        try { $Table.Rows.AllowBreakAcrossPages = -1 } catch {}
        try {
            $Table.Rows.Item(1).HeadingFormat = -1
            $Table.Rows.Item(1).Range.Font.Bold = -1
        } catch {}
        try { $Table.Rows.HeightRule = $wdRowHeightAuto } catch {}
    }
    catch {}
}

$word = $null
$doc = $null
Set-Progress 'Inicio'
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.Options.UpdateLinksAtOpen = $false
    $word.Options.ConfirmConversions = $false
    $word.Options.SaveInterval = 0

    $doc = $word.Documents.Open($DocumentPath, $false, $false, $false)
    $doc.TrackRevisions = $false

    Set-Progress 'Eliminación de líneas Fuente'
    $sourceLines = @(
        'Fuente: elaboración propia a partir de las evidencias y referencias citadas en el texto.',
        'Fuente: elaboración propia a partir de TFM_VISIBILIDAD_DETECCION_FP_WAZUH_DEFINITIVO_20260712.xlsx.',
        'Fuente: elaboración o captura propia del laboratorio.',
        'Fuente: elaboración propia a partir de TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx.'
    )
    $sourceRemoved = 0
    foreach ($source in $sourceLines) {
        $count = Get-OccurrenceCount -Document $doc -Text $source
        if ($count -gt 0) {
            $range = $doc.Content.Duplicate
            $find = $range.Find
            try {
                $find.ClearFormatting()
                $find.Replacement.ClearFormatting()
                $find.Text = $source + '^p'
                $find.Replacement.Text = ''
                $find.Forward = $true
                $find.Wrap = $wdFindContinue
                $find.Format = $false
                $find.MatchCase = $false
                $find.MatchWildcards = $false
                [void]$find.Execute($source + '^p', $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, '', $wdReplaceAll)
                $sourceRemoved += $count
                Add-Change -Category 'Fuentes' -Description "Eliminación de párrafos genéricos: $source" -Count $count
            }
            finally {
                Release-ComObject $find
                Release-ComObject $range
            }
        }
    }

    Set-Progress 'Redacción técnica y académica'
    [void](Replace-AllLiteral $doc 'y emitieron 379 alertas:' 'y emitieron/exportaron 379 filas CLIENT_EVENT, incluidas emisiones repetidas:' 'Terminología' 'Reformulación del resumen: 379 filas emitidas/exportadas')
    [void](Replace-AllLiteral $doc '379 alertas CLIENT_EVENT' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas' 'Terminología' 'Reformulación de 379 como filas CLIENT_EVENT, no alertas únicas')
    [void](Replace-AllLiteral $doc 'and emitted 379 alerts:' 'and emitted/exported 379 CLIENT_EVENT rows, including repeated emissions:' 'Terminología' 'Reformulación inglesa de 379 como filas emitidas/exportadas')
    [void](Replace-AllLiteral $doc '379 filas reales incluidas en la ventana de ejecución' '379 filas CLIENT_EVENT emitidas/exportadas, incluidas emisiones repetidas, dentro de la ventana de ejecución' 'Terminología' 'Unidad correcta en el anexo G')
    [void](Replace-AllLiteral $doc '10/10 pruebas benignas con 0 hits.' '10/10 pruebas benignas con 0 hits; 82 filas no reconciliadas fueron excluidas y no se computaron como falsos positivos definitivos.' 'Falsos positivos' 'Advertencia de 82 filas junto al KPI en el resumen')
    [void](Replace-AllLiteral $doc '10/10 benign tests with 0 hits.' '10/10 benign tests with 0 definitive hits; 82 unreconciled rows were excluded from the definitive false-positive count.' 'Falsos positivos' 'Advertencia de 82 filas en el abstract')
    [void](Replace-AllLiteral $doc '10/10 acciones benignas y 0 hits definitivos,' '10/10 acciones benignas y 0 hits definitivos; 82 filas no reconciliadas fueron excluidas y no se computaron como falsos positivos definitivos,' 'Falsos positivos' 'Advertencia de 82 filas en la metodología de FP')
    [void](Replace-AllLiteral $doc '10/10 acciones benignas ejecutadas y 0 hits definitivos.' '10/10 acciones benignas ejecutadas y 0 hits definitivos; 82 filas no reconciliadas fueron excluidas y no se computaron como falsos positivos definitivos.' 'Falsos positivos' 'Advertencia de 82 filas en anexos')
    [void](Replace-AllLiteral $doc '10/10 pruebas OK; 0 hits custom' '10/10 pruebas OK; 0 hits custom; 82 filas no reconciliadas excluidas del cómputo definitivo' 'Falsos positivos' 'Advertencia en tabla de resultados FP')
    [void](Replace-AllLiteral $doc 'tfm_wazuh_custom_rules_v1.xml' 'tfm_wazuh_custom_rules_v2.xml' 'Wazuh' 'Ruleset efectivo Wazuh v2')
    [void](Delete-ExactParagraph $doc '$ /home/admin/tfm_wazuh_custom_rules_v2.DISABLED.xml' 'Wazuh' 'Eliminación de comando histórico duplicado tras normalizar el ruleset v2')

    [void](Replace-AllLiteral $doc 'Cobertura táctica MITRE ATT&CK Enterprise' 'Tácticas ATT&CK representadas por el subconjunto evaluado' 'Alcance' 'Denominación acotada de cobertura táctica')
    [void](Replace-AllLiteral $doc 'Cobertura táctica MITRE ATT&CK' 'Tácticas ATT&CK representadas por el subconjunto evaluado' 'Alcance' 'Denominación acotada de cobertura táctica')
    [void](Replace-AllLiteral $doc 'Cobertura táctica' 'Tácticas ATT&CK representadas por el subconjunto evaluado' 'Alcance' 'Denominación acotada de cobertura táctica')

    [void](Replace-AllLiteral $doc 'visibilidad en herramienta webhook (Discord)' 'notificación/salida externa opcional mediante Discord' 'Capas' 'Discord reclasificado como salida externa')
    [void](Replace-AllLiteral $doc 'ósea' 'o sea' 'Ortografía' 'Corrección de «ósea»')
    [void](Replace-AllLiteral $doc 'almacenas evidencias' 'almacena evidencias' 'Ortografía' 'Concordancia verbal')
    [void](Replace-AllLiteral $doc 'script block login' 'Script Block Logging' 'Terminología' 'Normalización de Script Block Logging')
    [void](Replace-AllLiteral $doc 'evaluación HID,' 'evaluación HIDS,' 'Terminología' 'Corrección de HIDS')
    [void](Replace-AllLiteral $doc 'a mi como autor' 'a mí, como autor' 'Ortografía' 'Corrección de la declaración de IA')
    [void](Replace-AllLiteral $doc 'a mi, como autor' 'a mí, como autor' 'Ortografía' 'Corrección de la declaración de IA')

    Set-Progress 'Neutralización de rutas y enlaces privados'
    [void](Replace-AllLiteral $doc 'C:\Users\julio\Desktop\TFM' '<RAÍZ_TFM>' 'Privacidad' 'Neutralización de ruta personal del autor')
    [void](Replace-AllLiteral $doc 'C:\Users\seguridad\' '<HOST_LAB>\' 'Privacidad' 'Neutralización de rutas personales del host de laboratorio')
    [void](Replace-AllLiteral $doc '/home/admin/' '<HOST_LAB>/' 'Privacidad' 'Neutralización de ruta Linux personal')
    [void](Replace-AllLiteral $doc '192.168.1.145' '<HOST_LAB>' 'Privacidad' 'Neutralización de IP privada del servidor')
    [void](Replace-AllLiteral $doc '192.168.1.143' '<HOST_LAB>' 'Privacidad' 'Neutralización de IP privada del endpoint')
    [void](Replace-AllLiteral $doc '192.168.1.129:8088/upload' '<HOST_LAB>:8088/upload' 'Privacidad' 'Neutralización de IP privada del receptor')
    [void](Replace-AllLiteral $doc 'DUMB_LAB' '<RUTA_EVIDENCIAS>' 'Privacidad' 'Neutralización de etiqueta interna no válida como IOC')
    [void](Replace-AllLiteral $doc '08_MEMORIA/ENTREGABLE/ENTREGA_FINAL_EXCELES_TFM' '<RUTA_EVIDENCIAS>' 'Trazabilidad' 'Sustitución de ruta interna obsoleta')
    [void](Replace-AllLiteral $doc '<RAÍZ_TFM>\02_SCRIPTS\candidate' '<RUTA_ARTIFACTS>' 'Trazabilidad' 'Neutralización de ruta editorial candidate')
    [void](Replace-AllLiteral $doc 'https://documentation.wazuh.com/current/user-manual/capabilities/log-data-collection/how-to-collect-wlogs.html' 'https://documentation.wazuh.com/current/user-manual/capabilities/log-data-collection/configuration.html' 'Referencias' 'Actualización de URL oficial de recolección de eventos de Wazuh')

    for ($i = $doc.Hyperlinks.Count; $i -ge 1; $i--) {
        $link = $doc.Hyperlinks.Item($i)
        try {
            if ([string]$link.Address -match '^(?i)https?://(?:10\.|172\.(?:1[6-9]|2\d|3[01])\.|192\.168\.)') {
                $link.Delete()
                Add-Change -Category 'Privacidad' -Description 'Conversión de hipervínculo privado activo a texto neutralizado'
            }
        }
        finally {
            Release-ComObject $link
        }
    }

    Set-Progress 'Preguntas de investigación'
    [void](Set-ParagraphByNeedle $doc 'RQ1:' 'RQ1: ¿Qué cobertura de detección ofrece Velociraptor, mediante artifacts CLIENT_EVENT personalizados, sobre el subconjunto TEC-001 a TEC-009 evaluado en Windows?' 'RQ' 'Reescritura académica de RQ1')
    [void](Set-ParagraphByNeedle $doc 'RQ2:' 'RQ2: ¿Qué capacidades de visibilidad y detección específica aportan los artifacts públicos evaluados, y qué limitaciones presentan en el subconjunto experimental?' 'RQ' 'Reescritura académica de RQ2')
    [void](Set-ParagraphByNeedle $doc 'RQ3:' 'RQ3: ¿Qué impacto en CPU, memoria y duración introduce el cliente Velociraptor en los escenarios BASELINE_NO_VR, VR_IDLE y VR_TEC_RUNNER?' 'RQ' 'Reescritura académica de RQ3')
    [void](Set-ParagraphByNeedle $doc 'RQ4:' 'RQ4: ¿Cómo responde la arquitectura personalizada P1–P4 a las brechas de detección observadas en los artifacts públicos sin confundir detección, contexto y evidencia forense?' 'RQ' 'Reescritura académica de RQ4')
    [void](Set-ParagraphByNeedle $doc 'RQ5:' 'RQ5: ¿Qué comportamiento presentan los artifacts personalizados ante las pruebas benignas FP-001 a FP-010 y qué limitaciones afectan al cómputo definitivo de falsos positivos?' 'RQ' 'Reescritura académica de RQ5')
    [void](Set-ParagraphByNeedle $doc 'RQ6 ' 'RQ6: ¿Cómo deben separarse metodológicamente visibilidad, detección CLIENT_EVENT, alerta, evidencia forense y salida externa SERVER_EVENT/JSONL/Discord?' 'RQ' 'Reescritura académica de RQ6')
    $rq7Text = 'RQ7: ¿Cómo se compara el modelo HIDS/DFIR de Velociraptor con Wazuh en cobertura específica, alertabilidad y trazabilidad dentro del mismo subconjunto evaluado?' +
        "`r" + 'La comparación prevista inicialmente con FortiEDR se sustituyó por Wazuh como adaptación metodológica motivada por la disponibilidad de una plataforma HIDS/SIEM reproducible en el laboratorio. Este cambio no altera el objetivo comparativo y no implica equivalencia directa entre las unidades de ambos sistemas.' +
        "`r" + 'La metodología y los resultados mantienen una correspondencia explícita: RQ1, RQ2, RQ4 y RQ5 se responden en el capítulo VI; RQ3, en el capítulo VII; RQ6, mediante la separación de capas descrita en los capítulos V y VI; y RQ7, en el capítulo VIII y en la tabla final de correspondencia.'
    [void](Set-ParagraphByNeedle $doc 'RQ7 ' $rq7Text 'RQ' 'Reescritura académica de RQ7 y explicación de la adaptación FortiEDR–Wazuh')

    Set-Progress 'Control de versiones y metodología'
    [void](Set-ParagraphByNeedle $doc 'Durante el desarrollo se aplicó un versionado progresivo' 'El desarrollo aplicó control de versiones a scripts, artifacts, runners, configuraciones y documentos de apoyo. Cada elemento se vinculó a su versión, estado de validación, hash SHA-256 y evidencia asociada para conservar la trazabilidad entre código, ejecución y resultado.' 'Redacción' 'Síntesis académica del control de versiones')
    [void](Set-ParagraphByNeedle $doc 'Las iteraciones se identificaron mediante números y estados' 'La promoción de un componente a estado validado dependió de la revisión del autor y, cuando correspondía, de su ejecución en el laboratorio. La denominación del archivo se utilizó como identificador auxiliar y nunca como prueba suficiente de validez experimental.' 'Redacción' 'Eliminación de proceso editorial interno')
    [void](Delete-ParagraphByNeedle $doc 'Las versiones anteriores no se sobrescribieron de forma deliberada' 'Redacción' 'Eliminación de texto editorial sobre candidatos y backups')
    [void](Delete-ParagraphByNeedle $doc 'Además del versionado nominal, se utilizó Git' 'Redacción' 'Eliminación de texto editorial sobre commit y push')
    [void](Delete-ParagraphByNeedle $doc 'Antes de registrar cambios se revisó el estado del árbol' 'Redacción' 'Eliminación de texto editorial sobre git add')
    [void](Set-ParagraphByNeedle $doc 'Además, cualquier nuevo artifact debería mantenerse bajo el mismo ciclo de vida' 'Las evoluciones futuras de los artifacts deberán conservar un ciclo de vida trazable: creación como versión de trabajo, validación experimental frente al ground truth, revisión de ruido y promoción a versión validada solo cuando el resultado sea reproducible.' 'Redacción' 'Reescritura académica del ciclo de vida de artifacts')
    [void](Set-ParagraphByNeedle $doc 'D.1. Carpeta válida.' 'D.1. Control de versiones. Los artifacts validados se conservan en una ubicación separada y no se sobrescriben. Cualquier evolución futura se realiza sobre una nueva versión de trabajo, que debe validarse experimentalmente antes de su promoción.' 'Redacción' 'Neutralización de estados editoriales internos en el anexo D')

    $benchmarkParagraph = 'La fuente canónica del benchmark es TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_DEFINITIVO_20260712.xlsx. Sus controles registran 9/9 runs válidos, tres escenarios y tres repeticiones por escenario. El benchmark principal mide exclusivamente el impacto del cliente Velociraptor sobre el endpoint. SERVER_GUI se registra separadamente y se excluye del cálculo principal porque pertenece al servidor y no al cliente evaluado. Los flags reproducibles son IncludeServerGuiInTotal=False y ServerGuiExcludedFromClientMetrics=True. En VR_TEC_RUNNER se observaron 0,95 % de CPU media, 54 % de pico, 57,42 MB de RAM media y 57,97 MB de pico. RunnerStillRunningAtEnd=True se conserva como WARN no bloqueante en las tres repeticiones VR_TEC_RUNNER. Estas métricas no atribuyen CPU o RAM a P1, P2, P3 o P4 individualmente y no miden calidad de detección (García-Amorena Reina, 2026d).'
    [void](Set-ParagraphByNeedle $doc 'La fuente canónica del benchmark es' $benchmarkParagraph 'Benchmark' 'Metodología correcta de exclusión de SERVER_GUI y flags literales')
    [void](Set-ParagraphByNeedle $doc 'H.3. Exclusión SERVER_GUI.' 'H.3. Exclusión de SERVER_GUI. El benchmark principal mide exclusivamente el impacto del cliente Velociraptor sobre el endpoint. SERVER_GUI se registra separadamente y se excluye del cálculo principal porque pertenece al servidor y no al cliente evaluado. Los flags empleados son IncludeServerGuiInTotal=False y ServerGuiExcludedFromClientMetrics=True.' 'Benchmark' 'Metodología y flags en el anexo H')

    $rulesParagraph = 'Para comparar Wazuh frente a Velociraptor en condiciones metodológicamente comparables, se creó una capa de reglas personalizadas TFM. El ruleset efectivo de la campaña fue tfm_wazuh_custom_rules_v2.xml, compatible con los identificadores 110301, 110402, 110203 y 110202 y con el gap 110201=0 para TEC-009. Las reglas custom se interpretan como aproximación funcional y no como equivalentes directos de los artifacts de Velociraptor (Wazuh, s. f.-a; García-Amorena Reina, 2026c; García-Amorena Reina, 2026e).'
    [void](Set-ParagraphByNeedle $doc 'Para comparar Wazuh frente a Velociraptor en condiciones metodológicamente comparables' $rulesParagraph 'Wazuh' 'Identificación del ruleset efectivo v2')
    [void](Set-ParagraphByNeedle $doc 'I.5. Reglas custom.' 'I.5. Reglas custom. El ruleset efectivo de la campaña fue tfm_wazuh_custom_rules_v2.xml. El Excel consolidado conserva los resultados finales, los identificadores 110301, 110402, 110203 y 110202, y el gap 110201=0 para TEC-009.' 'Wazuh' 'Ruleset v2 en el anexo I')

    $discordParagraph = 'Discord no se considera una fuente primaria de telemetría ni un mecanismo de detección del host, sino un canal externo de notificación. Su función dentro del TFM es comprobar que una detección generada por un artifact CLIENT_EVENT puede procesarse mediante SERVER_EVENT y llegar a una salida externa.' +
        "`r" + 'En las capturas, el campo denominado «IOC» se interpreta como contexto o IOA capturado por el laboratorio cuando contiene rutas, nombres de scripts o etiquetas TEC. Esos valores no son IOC reales ni condiciones de detección y no se utilizan para acreditar cobertura.'
    [void](Set-ParagraphByNeedle $doc 'Discord no se considera una fuente primaria de telemetría' $discordParagraph 'Capas' 'Separación de Discord y explicación del campo IOC')

    [void](Delete-ExactParagraph $doc '/' 'Limpieza' 'Eliminación de carácter residual aislado')

    Set-Progress 'Captions y títulos'
    [void](Rebuild-Caption $doc 'Tabla 32' $wdStyleHeading3 'Tabla' 'Criterios de control de versiones y trazabilidad del proyecto' 'Tabla 32 con título y estilo de leyenda')
    [void](Rebuild-Caption $doc 'Código 5Control del laboratorio' $wdStyleHeading3 'Código' 'Control del laboratorio' 'Código 5 con dos puntos, espacio y estilo de leyenda')
    [void](Rebuild-Caption $doc 'Tabla 63: Comparación final entre Wazuh base' $wdStyleCaption 'Tabla' 'Comparación final entre Wazuh base, Wazuh custom y Velociraptor custom' 'Tabla comparativa reconstruida con un único campo SEQ')
    [void](Replace-InStyledParagraph $doc 'Figura 15: Directorio máquina física' $wdStyleCaption 'Directorio máquina física' 'Directorio de la máquina virtual' 'Figura 15: título correcto')
    [void](Replace-InStyledParagraph $doc 'Tabla 45: Artifacts públicos considerados' $wdStyleCaption 'Artifacts públicos considerados en la evaluación metodológica' 'Brechas y limitaciones de los artifacts públicos en la evaluación' 'Tabla 45: título sobre brechas y limitaciones')
    [void](Replace-InStyledParagraph $doc 'Tabla 46: Escenarios Benchmark' $wdStyleCaption 'Escenarios Benchmark' 'Escenarios ejecutados en el benchmark de rendimiento' 'Tabla 46: escenarios finales ejecutados')
    [void](Replace-InStyledParagraph $doc 'Tabla 54: Distribución de alertas CLIENT_EVENT' $wdStyleCaption 'Distribución de alertas CLIENT_EVENT por perfil custom' 'Distribución de filas CLIENT_EVENT emitidas/exportadas por perfil custom' 'Tabla 54: unidad de filas CLIENT_EVENT')
    [void](Replace-InStyledParagraph $doc 'Tabla 68: Respuesta verificable a los objetivos del TFM' $wdStyleCaption 'Respuesta verificable a los objetivos del TFM' 'Correspondencia entre preguntas de investigación, resultados y limitaciones' 'Tabla 68: correspondencia RQ–resultados')

    $captionRenames = @(
        @('Tabla 13: FLR - HIDS','FLR - HIDS','Funciones, limitaciones y riesgos del enfoque HIDS'),
        @('Tabla 14: FLR - DFIR','FLR - DFIR','Funciones, limitaciones y riesgos del enfoque DFIR'),
        @('Tabla 15: FLR - EDR','FLR - EDR','Funciones, limitaciones y riesgos del enfoque EDR'),
        @('Tabla 16: Resumen posición VR','Resumen posición VR','Posición de Velociraptor frente a HIDS, DFIR y EDR'),
        @('Tabla 24: Nota Cap. IV, metodología','Nota Cap. IV, metodología','Criterios metodológicos del capítulo IV'),
        @('Tabla 30: Nota Cap. IV, Sysmon ID 26','Nota Cap. IV, Sysmon ID 26','Tratamiento metodológico de Sysmon Event ID 26'),
        @('Tabla 39: Nota Cap. IV, 4.3.8. Scripts de laboratorio validados','Nota Cap. IV, 4.3.8. Scripts de laboratorio validados','Criterios de validación de los scripts de laboratorio'),
        @('Tabla 67: MITRE ampliación','MITRE ampliación','Técnicas MITRE ATT&CK propuestas para trabajo futuro')
    )
    foreach ($item in $captionRenames) {
        [void](Replace-InStyledParagraph $doc $item[0] $wdStyleCaption $item[1] $item[2] "Título descriptivo: $($item[2])")
    }

    foreach ($styleConstant in @($wdStyleCaption, $wdStyleHeading1, $wdStyleHeading2, $wdStyleHeading3)) {
        $range = $doc.Content.Duplicate
        $find = $range.Find
        try {
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Style = $doc.Styles.Item($styleConstant)
            $find.Text = '.^p'
            $find.Replacement.Text = '^p'
            $find.Format = $true
            $find.MatchWildcards = $false
            if ($find.Execute('.^p', $false, $false, $false, $false, $false, $true, $wdFindContinue, $true, '^p', $wdReplaceAll)) {
                Add-Change -Category 'Títulos' -Description "Eliminación de puntos finales en estilo $styleConstant"
            }
        }
        finally {
            Release-ComObject $find
            Release-ComObject $range
        }
    }

    Set-Progress 'Tablas técnicas'
    $benchmarkTable = $doc.Tables.Item(51)
    try {
        while ($benchmarkTable.Rows.Count -gt 4) {
            $benchmarkTable.Rows.Item($benchmarkTable.Rows.Count).Delete()
        }
        Set-CellText $benchmarkTable 1 1 'Escenario'
        Set-CellText $benchmarkTable 1 2 'Objetivo'
        Set-CellText $benchmarkTable 2 1 'BASELINE_NO_VR'
        Set-CellText $benchmarkTable 2 2 'Medir el rendimiento basal de la máquina virtual sin el cliente Velociraptor activo.'
        Set-CellText $benchmarkTable 3 1 'VR_IDLE'
        Set-CellText $benchmarkTable 3 2 'Medir el impacto del cliente Velociraptor activo sin ejecutar la carga de técnicas.'
        Set-CellText $benchmarkTable 4 1 'VR_TEC_RUNNER'
        Set-CellText $benchmarkTable 4 2 'Medir el impacto del cliente Velociraptor durante la ejecución controlada del runner TEC.'
        Format-TableForDelivery $benchmarkTable 9.5
        Add-Change -Category 'Benchmark' -Description 'Tabla 46 corregida a tres escenarios reales y tres repeticiones por escenario' -Count 8
    }
    finally { Release-ComObject $benchmarkTable }

    $profileTable = $doc.Tables.Item(70)
    try {
        Set-CellText $profileTable 1 4 'Filas CLIENT_EVENT'
        Set-CellText $profileTable 2 5 'Señales críticas selectivas'
        Set-CellText $profileTable 3 5 'Detección, contexto y evidencia forense; no constituye una unidad homogénea'
        Set-CellText $profileTable 6 4 'No aplica'
        Format-TableForDelivery $profileTable 8.5
        Add-Change -Category 'Tablas' -Description 'Tabla 52: unidad de filas y rol heterogéneo de P2' -Count 4
    }
    finally { Release-ComObject $profileTable }

    $distributionTable = $doc.Tables.Item(72)
    try {
        Set-CellText $distributionTable 1 4 'Filas emitidas/exportadas'
        Set-CellText $distributionTable 3 5 'Detección, contexto y evidencia forense; interpretar según la source'
        Set-CellText $distributionTable 6 5 '379 filas emitidas/exportadas, incluidas emisiones repetidas; no equivalen a alertas únicas ni a cobertura'
        Format-TableForDelivery $distributionTable 8.5
        Add-Change -Category 'Tablas' -Description 'Tabla 54: formulación exacta de 379 filas y rol P2' -Count 3
    }
    finally { Release-ComObject $distributionTable }

    foreach ($idx in @(29, 50, 52, 71, 75, 100)) {
        $table = $doc.Tables.Item($idx)
        try {
            Format-TableForDelivery $table 8.5
            try { $table.Rows.AllowBreakAcrossPages = 0 } catch {}
        }
        finally { Release-ComObject $table }
    }
    Add-Change -Category 'Maquetación' -Description 'Ajuste a ventana, márgenes internos y encabezados de tablas 23, 45, 47, 53, 56 y 65' -Count 6

    $rqTable = $doc.Tables.Item(104)
    try {
        while ($rqTable.Rows.Count -lt 8) { [void]$rqTable.Rows.Add() }
        while ($rqTable.Rows.Count -gt 8) { $rqTable.Rows.Item($rqTable.Rows.Count).Delete() }
        $rqRows = @(
            @('Pregunta de investigación','Resultado','Evidencia principal','Limitación'),
            @('RQ1 · Cobertura custom','9/9 en TEC-001 a TEC-009','Filas CLIENT_EVENT de P1–P4 y matriz por técnica','Solo aplica al subconjunto y laboratorio evaluados'),
            @('RQ2 · Artifacts públicos','Cinco campañas positivas de visibilidad y una no concluyente; Hayabusa CH 3/9','Campañas públicas y matriz de resultados','ETW no concluyente; TrackNetwork sin atribución fiable'),
            @('RQ3 · Rendimiento','9/9 runs válidos en tres escenarios','BASELINE_NO_VR, VR_IDLE y VR_TEC_RUNNER; CPU, RAM y duración','Coste del cliente en laboratorio; no extrapolable a producción'),
            @('RQ4 · Arquitectura personalizada','P1–P4 cubren las brechas evaluadas y Router no detecta','Cuatro detectores CLIENT_EVENT y un Router SERVER_EVENT','P2 mezcla detección, contexto y evidencia; requiere interpretación por source'),
            @('RQ5 · Falsos positivos','10/10 pruebas benignas; 0 hits definitivos','summary.json y vr_hits.csv reconciliados','82 filas no reconciliadas excluidas del cómputo definitivo'),
            @('RQ6 · Separación de capas','Visibilidad, detección, alerta, evidencia y salida externa se tratan por separado','CLIENT_EVENT detecta; SERVER_EVENT/JSONL/Discord enrutan, persisten o notifican','La salida externa no acredita detección del host'),
            @('RQ7 · Comparación con Wazuh','Wazuh base 3/9; custom 4/9; unión complementaria 7/9','Reglas nativas y tfm_wazuh_custom_rules_v2.xml','Unidades y modelos operativos no equivalentes')
        )
        for ($r = 1; $r -le 8; $r++) {
            for ($c = 1; $c -le 4; $c++) {
                Set-CellText $rqTable $r $c $rqRows[$r-1][$c-1]
            }
        }
        Format-TableForDelivery $rqTable 8
        Add-Change -Category 'RQ' -Description 'Tabla 68 reescrita con correspondencia uno a uno RQ–resultado–evidencia–limitación' -Count 32
    }
    finally { Release-ComObject $rqTable }

    for ($i = 1; $i -le $doc.Tables.Count; $i++) {
        $table = $doc.Tables.Item($i)
        try {
            if ($table.Rows.Count -gt 1) {
                if ($i -ne 2) {
                    try { $table.Rows.Item(1).HeadingFormat = -1 } catch {}
                }
                if ($i -in @(29, 50, 51, 52, 70, 71, 72, 75, 100, 104)) {
                    try { $table.Rows.AllowBreakAcrossPages = 0 } catch {}
                }
                else {
                    try { $table.Rows.AllowBreakAcrossPages = -1 } catch {}
                }
            }
        }
        finally { Release-ComObject $table }
    }
    Add-Change -Category 'Maquetación' -Description 'Encabezados repetibles y filas multipágina habilitados en tablas aplicables' -Count $doc.Tables.Count

    Set-Progress 'Acrónimos'
    Normalize-AcronymTable -Document $doc -TableIndex 2

    Set-Progress 'Jerarquía e índices'
    [void](Set-ParagraphByNeedle $doc 'Referencias bibliográficas' 'Referencias bibliográficas' 'Jerarquía' 'Referencias incluidas en el índice' -NewStyleConstant $wdStyleHeading1)
    [void](Set-ParagraphByNeedle $doc 'Anexos' 'Anexos' 'Jerarquía' 'Anexos incluidos en el índice' -NewStyleConstant $wdStyleHeading1)
    foreach ($letter in [char[]]'ABCDEFGHIJKLM') {
        $needle = "Anexo $letter."
        $paragraph = Find-Paragraph -Document $doc -Needle $needle
        if ($null -ne $paragraph) {
            try {
                $paragraph.Range.Style = $doc.Styles.Item($wdStyleHeading2)
                Add-Change -Category 'Jerarquía' -Description "Anexo $letter aplicado como Título 2"
            }
            finally { Release-ComObject $paragraph }
        }
    }

    try {
        $normalStyle = $doc.Styles.Item(-1)
        $normalStyle.ParagraphFormat.WidowControl = -1
        Release-ComObject $normalStyle
    } catch {}
    foreach ($styleConstant in @($wdStyleHeading1, $wdStyleHeading2, $wdStyleHeading3)) {
        $style = $doc.Styles.Item($styleConstant)
        try {
            $style.ParagraphFormat.KeepWithNext = -1
            $style.ParagraphFormat.KeepTogether = -1
            $style.ParagraphFormat.WidowControl = -1
        }
        finally { Release-ComObject $style }
    }
    Add-Change -Category 'Maquetación' -Description 'Control de viudas, huérfanas y títulos con contenido siguiente' -Count 4

    Set-Progress 'Actualización de campos'
    foreach ($field in @($doc.Fields)) {
        try { [void]$field.Update() } catch {}
        Release-ComObject $field
    }
    foreach ($storyType in 1..17) {
        try {
            $story = $doc.StoryRanges.Item($storyType)
            while ($null -ne $story) {
                foreach ($field in @($story.Fields)) {
                    try { [void]$field.Update() } catch {}
                    Release-ComObject $field
                }
                $next = $story.NextStoryRange
                Release-ComObject $story
                $story = $next
            }
        } catch {}
    }
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $toc = $doc.TablesOfContents.Item($i)
        try { $toc.Update() } finally { Release-ComObject $toc }
    }
    for ($i = 1; $i -le $doc.TablesOfFigures.Count; $i++) {
        $tof = $doc.TablesOfFigures.Item($i)
        try { $tof.Update() } finally { Release-ComObject $tof }
    }
    foreach ($field in @($doc.Fields)) {
        try { [void]$field.Update() } catch {}
        Release-ComObject $field
    }
    Add-Change -Category 'Campos' -Description 'Actualización completa de campos, índice general y listas de tablas, figuras y códigos' -Count 4

    Set-Progress 'Guardado'
    $doc.Save()
    $doc.Repaginate()
    $doc.Save()
    $pages = [int]$doc.ComputeStatistics(2)
    $words = [int]$doc.ComputeStatistics(0)

    $remainingText = [string]$doc.Content.Text
    $remainingSources = [regex]::Matches($remainingText, '(?i)Fuente:\s*(elaboración|elaboración o captura)').Count
    $remainingPrivateLinks = 0
    foreach ($link in @($doc.Hyperlinks)) {
        try {
            if ([string]$link.Address -match '^(?i)https?://(?:10\.|172\.(?:1[6-9]|2\d|3[01])\.|192\.168\.)') {
                $remainingPrivateLinks++
            }
        }
        finally { Release-ComObject $link }
    }

    $result = [pscustomobject]@{
        DocumentPath = $DocumentPath
        SourceParagraphsRemoved = $sourceRemoved
        Corrections = $script:corrections
        Pages = $pages
        Words = $words
        RemainingGenericSources = $remainingSources
        RemainingPrivateHyperlinks = $remainingPrivateLinks
        Changes = @($script:changes)
        SHA256 = (Get-FileHash -LiteralPath $DocumentPath -Algorithm SHA256).Hash
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $LogPath -Encoding UTF8

    $doc.Close($false)
    Release-ComObject $doc
    $doc = $null
    $word.Quit()
    Release-ComObject $word
    $word = $null
    Set-Progress 'COMPLETADO'
    $result | ConvertTo-Json -Depth 5
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
        Release-ComObject $doc
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
