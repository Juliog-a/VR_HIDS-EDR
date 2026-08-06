[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$WorkingDoc,
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [Parameter(Mandatory = $true)][string]$FiguresDir,
    [Parameter(Mandatory = $true)][string]$OutputDoc,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [ValidateSet('all', 'phase1', 'phase2', 'phase3')][string]$Phase = 'all'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$wdFindStop = 0
$wdReplaceAll = 2
$wdCollapseEnd = 0
$wdFieldEmpty = -1
$wdStyleNormal = -1
$wdStyleHeading2 = -3
$wdStyleCaption = -35
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdCellAlignVerticalCenter = 1
$wdAutoFitFixed = 0
$wdRowHeightAuto = 0
$wdPreferredWidthPoints = 2
$wdFormatDocumentDefault = 16
$wdDoNotSaveChanges = 0
$wdStatisticPages = 2
$wdStatisticWords = 0
$wdStatisticCharactersWithSpaces = 5
$script:ProgressPath = $null

function Write-BuildProgress {
    param([string]$Message)
    if ([string]::IsNullOrWhiteSpace($script:ProgressPath)) { return }
    try { [IO.File]::AppendAllText($script:ProgressPath, "$(Get-Date -Format o) | $Message`r`n", $utf8NoBom) } catch {}
}

function Release-ComObject {
    param([AllowNull()][object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Clean-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace "`r", ' ' -replace "`a", ' ' -replace "`v", ' ' -replace "`t", ' ') -replace '\s+', ' ').Trim()
}

function Get-StyleName {
    param([object]$Range)
    $style = $null
    try {
        $style = $Range.Style
        if ($style -is [string]) { return [string]$style }
        try { return [string]$style.NameLocal } catch { return [string]$style }
    }
    catch { return '' }
    finally { if ($null -ne $style -and $style -isnot [string]) { Release-ComObject $style } }
}

function Find-ParagraphByPrefix {
    param([object]$Doc, [string]$Prefix, [switch]$RequireHeading, [switch]$RequireSeq)
    $search = $null
    $find = $null
    try {
        $search = $Doc.Content.Duplicate
        $find = $search.Find
        $find.ClearFormatting()
        $find.Text = $Prefix
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        $find.Format = $false
        $find.MatchCase = $false
        $find.MatchWholeWord = $false
        while ($find.Execute()) {
            $para = $null
            $pr = $null
            try {
                $para = $search.Paragraphs.Item(1)
                $pr = $para.Range.Duplicate
                $ok = $true
                if ($RequireHeading) {
                    $outline = 10
                    try { $outline = [int]$para.OutlineLevel } catch {}
                    $styleName = Get-StyleName $pr
                    $ok = (($outline -ge 1 -and $outline -le 9) -and ($styleName -match '^(Heading|T[ií]tulo)\s*[1-9]'))
                }
                if ($ok -and $RequireSeq) {
                    $ok = $false
                    $fields = $null
                    try {
                        $fields = $pr.Fields
                        for ($i = 1; $i -le $fields.Count; $i++) {
                            $field = $null
                            $code = $null
                            try {
                                $field = $fields.Item($i)
                                $code = $field.Code
                                if ((Clean-Text ([string]$code.Text)) -match '^SEQ\s+(Tabla|Figura)\b') { $ok = $true; break }
                            }
                            finally { Release-ComObject $code; Release-ComObject $field }
                        }
                    }
                    finally { Release-ComObject $fields }
                }
                if ($ok) { return $pr.Duplicate }
            }
            finally {
                Release-ComObject $pr
                Release-ComObject $para
            }
            $next = [math]::Min($Doc.Content.End, $search.End + 1)
            $search.SetRange($next, $Doc.Content.End)
            $find = $search.Find
            $find.ClearFormatting()
            $find.Text = $Prefix
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            $find.Format = $false
            $find.MatchCase = $false
            $find.MatchWholeWord = $false
        }
        return $null
    }
    finally { Release-ComObject $find; Release-ComObject $search }
}

function Find-SeqCaptionByPrefix {
    param([object]$Doc, [string]$Prefix)
    $fields = $null
    try {
        $fields = $Doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            $para = $null
            $pr = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                if ((Clean-Text ([string]$code.Text)) -notmatch '^SEQ\s+(Tabla|Figura)\b') { continue }
                $para = $field.Result.Paragraphs.Item(1)
                $pr = $para.Range
                $captionText = Clean-Text ([string]$pr.Text)
                if ($captionText.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) {
                    return $pr.Duplicate
                }
            }
            finally { Release-ComObject $pr; Release-ComObject $para; Release-ComObject $code; Release-ComObject $field }
        }
        return $null
    }
    finally { Release-ComObject $fields }
}

function Find-SeqCaptionByTitle {
    param([object]$Doc, [string]$Title)
    $fields = $null
    try {
        $fields = $Doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            $para = $null
            $pr = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                if ((Clean-Text ([string]$code.Text)) -notmatch '^SEQ\s+(Tabla|Figura)\b') { continue }
                $para = $field.Result.Paragraphs.Item(1)
                $pr = $para.Range
                $captionText = Clean-Text ([string]$pr.Text)
                if ($captionText.IndexOf($Title, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    return $pr.Duplicate
                }
            }
            finally { Release-ComObject $pr; Release-ComObject $para; Release-ComObject $code; Release-ComObject $field }
        }
        return $null
    }
    finally { Release-ComObject $fields }
}

function Find-HeadingByText {
    param([object]$Doc, [string]$Text, [int]$StyleId)
    $range = $null
    $find = $null
    $para = $null
    try {
        $range = $Doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $Text
        $find.Style = $StyleId
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        $find.MatchCase = $false
        if (-not $find.Execute()) { return $null }
        $para = $range.Paragraphs.Item(1)
        return $para.Range.Duplicate
    }
    finally { Release-ComObject $para; Release-ComObject $find; Release-ComObject $range }
}

function Replace-ParagraphByPrefix {
    param([object]$Doc, [string]$Prefix, [string]$Text)
    $pr = Find-ParagraphByPrefix -Doc $Doc -Prefix $Prefix
    if ($null -eq $pr) { return 0 }
    $body = $null
    try {
        $end = [int]$pr.End
        if ($end -gt [int]$pr.Start) { $end-- }
        $body = $Doc.Range([int]$pr.Start, $end)
        $body.Text = $Text
        return 1
    }
    finally { Release-ComObject $body; Release-ComObject $pr }
}

function Replace-LiteralAll {
    param([object]$Doc, [string]$Old, [string]$New)
    $range = $null
    $find = $null
    try {
        $range = $Doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $Old
        $find.Replacement.Text = $New
        $find.Forward = $true
        $find.Wrap = 1
        $find.Format = $false
        $find.MatchCase = $true
        [void]$find.Execute($Old, $true, $false, $false, $false, $false, $true, 1, $false, $New, $wdReplaceAll)
        return 1
    }
    finally { Release-ComObject $find; Release-ComObject $range }
}

function Get-ClosestTableBefore {
    param([object]$Doc, [int]$Position)
    $tables = $null
    $best = $null
    $bestEnd = -1
    try {
        $tables = $Doc.Tables
        for ($i = 1; $i -le $tables.Count; $i++) {
            $table = $null
            $range = $null
            try {
                $table = $tables.Item($i)
                $range = $table.Range
                if ([int]$range.End -le $Position -and [int]$range.End -gt $bestEnd) {
                    Release-ComObject $best
                    $best = $table
                    $table = $null
                    $bestEnd = [int]$range.End
                }
            }
            finally { Release-ComObject $range; Release-ComObject $table }
        }
        return $best
    }
    finally { Release-ComObject $tables }
}

function Format-Table {
    param([object]$Table, [double[]]$Widths)
    try { $Table.Range.Style = $wdStyleNormal } catch {}
    try { $Table.Range.ListFormat.RemoveNumbers() } catch {}
    try { $Table.AllowAutoFit = $false } catch {}
    try { $Table.AutoFitBehavior($wdAutoFitFixed) } catch {}
    try { $Table.PreferredWidthType = $wdPreferredWidthPoints; $Table.PreferredWidth = 450 } catch {}
    try { $Table.Borders.Enable = 1 } catch {}
    try { $Table.Rows.AllowBreakAcrossPages = 0 } catch {}
    try { $Table.Rows.HeightRule = $wdRowHeightAuto } catch {}
    try { $Table.Rows.Item(1).HeadingFormat = -1 } catch {}
    try { $Table.Range.Font.Name = 'Arial'; $Table.Range.Font.Size = 8.5 } catch {}
    try {
        $Table.Range.ParagraphFormat.SpaceBefore = 0
        $Table.Range.ParagraphFormat.SpaceAfter = 2
        $Table.Range.ParagraphFormat.LineSpacingRule = 0
    } catch {}
    for ($c = 1; $c -le $Table.Columns.Count; $c++) {
        if ($c -le $Widths.Count) {
            try { $Table.Columns.Item($c).PreferredWidthType = $wdPreferredWidthPoints; $Table.Columns.Item($c).PreferredWidth = 450 * $Widths[$c - 1] } catch {}
        }
    }
    for ($c = 1; $c -le $Table.Columns.Count; $c++) {
        $cell = $null
        try {
            $cell = $Table.Cell(1, $c)
            $cell.Shading.BackgroundPatternColor = 7949855
            $cell.Range.Font.Color = 16777215
            $cell.Range.Font.Bold = -1
            $cell.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
            $cell.VerticalAlignment = $wdCellAlignVerticalCenter
        }
        finally { Release-ComObject $cell }
    }
    for ($r = 2; $r -le $Table.Rows.Count; $r++) {
        for ($c = 1; $c -le $Table.Columns.Count; $c++) {
            $cell = $null
            try {
                $cell = $Table.Cell($r, $c)
                $cell.VerticalAlignment = $wdCellAlignVerticalCenter
                $cell.Range.ParagraphFormat.Alignment = $wdAlignParagraphLeft
                if (($r % 2) -eq 0) { $cell.Shading.BackgroundPatternColor = 15921906 }
            }
            finally { Release-ComObject $cell }
        }
    }
}

function Create-TableAt {
    param([object]$Doc, [int]$Position, [object[]]$Headers, [object[]]$Rows, [double[]]$Widths)
    $insert = $null
    $table = $null
    try {
        $insert = $Doc.Range($Position, $Position)
        $table = $Doc.Tables.Add($insert, $Rows.Count + 1, $Headers.Count)
        for ($c = 1; $c -le $Headers.Count; $c++) { $table.Cell(1, $c).Range.Text = [string]$Headers[$c - 1] }
        for ($r = 1; $r -le $Rows.Count; $r++) {
            $row = @($Rows[$r - 1])
            for ($c = 1; $c -le $Headers.Count; $c++) { $table.Cell($r + 1, $c).Range.Text = [string]$row[$c - 1] }
        }
        Format-Table -Table $table -Widths $Widths
        return $table
    }
    finally { Release-ComObject $insert }
}

function Set-CaptionTitle {
    param([object]$Doc, [object]$CaptionRange, [string]$Title)
    $fields = $null
    $field = $null
    $code = $null
    $search = $null
    $find = $null
    $insert = $null
    try {
        $fields = $CaptionRange.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $candidate = $null
            $candidateCode = $null
            try {
                $candidate = $fields.Item($i)
                $candidateCode = $candidate.Code
                if ((Clean-Text ([string]$candidateCode.Text)) -match '^SEQ\s+(Tabla|Figura)\b') {
                    $field = $candidate
                    $candidate = $null
                    break
                }
            }
            finally { Release-ComObject $candidateCode; Release-ComObject $candidate }
        }
        if ($null -eq $field) { throw 'Caption without SEQ field.' }
        $captionText = Clean-Text ([string]$CaptionRange.Text)
        $colon = $captionText.IndexOf(':')
        if ($colon -ge 0 -and $colon + 1 -lt $captionText.Length) {
            $oldTitle = $captionText.Substring($colon + 1).Trim()
            $search = $CaptionRange.Duplicate
            $find = $search.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Text = $oldTitle
            $find.Replacement.Text = $Title
            $find.Forward = $true
            $find.Wrap = $wdFindStop
            [void]$find.Execute($oldTitle, $true, $false, $false, $false, $false, $true, $wdFindStop, $false, $Title, $wdReplaceAll)
        }
        else {
            $insert = $Doc.Range([int]$field.Result.End, [int]$field.Result.End)
            $insert.InsertAfter(': ' + $Title)
        }
        try { $CaptionRange.Style = $wdStyleCaption } catch {}
    }
    finally { Release-ComObject $insert; Release-ComObject $find; Release-ComObject $search; Release-ComObject $code; Release-ComObject $field; Release-ComObject $fields }
}

function Replace-TableByCaption {
    param([object]$Doc, [object]$Spec)
    Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + ' | localizando caption')
    $caption = Find-SeqCaptionByPrefix -Doc $Doc -Prefix ([string]$Spec.caption_prefix)
    if ($null -eq $caption) { throw "Caption not found: $($Spec.caption_prefix)" }
    $oldTable = $null
    $oldRange = $null
    $separator = $null
    $newTable = $null
    try {
        Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + ' | localizando tabla precedente')
        $oldTable = Get-ClosestTableBefore -Doc $Doc -Position ([int]$caption.Start)
        if ($null -eq $oldTable) { throw "Table not found before $($Spec.caption_prefix)" }
        $oldRange = $oldTable.Range
        $pos = [int]$oldRange.Start
        Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + " | eliminando tabla en posicion $pos")
        $oldTable.Delete()
        $separator = $Doc.Range($pos, $pos)
        $separator.InsertBefore("`r")
        Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + ' | creando tabla canonica')
        $newTable = Create-TableAt -Doc $Doc -Position $pos -Headers @($Spec.headers) -Rows @($Spec.rows) -Widths ([double[]]@($Spec.widths))
        Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + ' | actualizando caption')
        Release-ComObject $caption
        $caption = Find-SeqCaptionByPrefix -Doc $Doc -Prefix ([string]$Spec.caption_prefix)
        Set-CaptionTitle -Doc $Doc -CaptionRange $caption -Title ([string]$Spec.caption_title)
        Write-BuildProgress ("Tabla " + [string]$Spec.caption_prefix + ' | completada')
    }
    finally { Release-ComObject $separator; Release-ComObject $newTable; Release-ComObject $oldRange; Release-ComObject $oldTable; Release-ComObject $caption }
}

function Find-ExactTokenRange {
    param([object]$Doc, [string]$Token)
    $range = $null
    $find = $null
    try {
        $range = $Doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Text = $Token
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        $find.MatchCase = $true
        if ($find.Execute()) { return $range.Duplicate }
        return $null
    }
    finally { Release-ComObject $find; Release-ComObject $range }
}

function Replace-TokenWithSeqField {
    param([object]$Doc, [string]$Token, [string]$Category, [string]$Bookmark)
    $range = Find-ExactTokenRange -Doc $Doc -Token $Token
    $field = $null
    $result = $null
    try {
        if ($null -eq $range) { throw "SEQ token not found: $Token" }
        $range.Text = ''
        $field = $Doc.Fields.Add($range, $wdFieldEmpty, "SEQ $Category \* ARABIC", $true)
        [void]$field.Update()
        $result = $field.Result.Duplicate
        if ($Doc.Bookmarks.Exists($Bookmark)) { $Doc.Bookmarks.Item($Bookmark).Delete() }
        [void]$Doc.Bookmarks.Add($Bookmark, $result)
    }
    finally { Release-ComObject $result; Release-ComObject $field; Release-ComObject $range }
}

function Replace-RefTokens {
    param([object]$Doc)
    $count = 0
    while ($true) {
        $content = [string]$Doc.Content.Text
        $match = [regex]::Match($content, '\[\[REF:([A-Za-z0-9_]+)\]\]')
        if (-not $match.Success) { break }
        $token = $match.Value
        $bookmark = $match.Groups[1].Value
        if (-not $Doc.Bookmarks.Exists($bookmark)) { throw "Bookmark missing for REF: $bookmark" }
        $range = Find-ExactTokenRange -Doc $Doc -Token $token
        $field = $null
        try {
            if ($null -eq $range) { throw "REF token not found: $token" }
            $range.Text = ''
            $field = $Doc.Fields.Add($range, $wdFieldEmpty, "REF $bookmark \h", $true)
            [void]$field.Update()
            $count++
        }
        finally { Release-ComObject $field; Release-ComObject $range }
    }
    return $count
}

function Style-ParagraphByPrefix {
    param([object]$Doc, [string]$Prefix, [int]$Style)
    $pr = Find-ParagraphByPrefix -Doc $Doc -Prefix $Prefix
    try { if ($null -ne $pr) { $pr.Style = $Style; return 1 }; return 0 }
    finally { Release-ComObject $pr }
}

function Replace-PlaceholderWithTable {
    param([object]$Doc, [string]$Placeholder, [object[]]$Headers, [object[]]$Rows, [double[]]$Widths)
    $token = Find-ExactTokenRange -Doc $Doc -Token $Placeholder
    $para = $null
    $body = $null
    $table = $null
    try {
        if ($null -eq $token) { throw "Table placeholder not found: $Placeholder" }
        $para = $token.Paragraphs.Item(1)
        $start = [int]$para.Range.Start
        $body = $Doc.Range($start, [int]$para.Range.End - 1)
        $body.Text = ''
        $table = Create-TableAt -Doc $Doc -Position $start -Headers $Headers -Rows $Rows -Widths $Widths
        return $table
    }
    finally { Release-ComObject $body; Release-ComObject $para; Release-ComObject $token }
}

function Replace-PlaceholderWithImage {
    param([object]$Doc, [string]$Placeholder, [string]$ImagePath, [string]$AltText)
    $token = Find-ExactTokenRange -Doc $Doc -Token $Placeholder
    $para = $null
    $body = $null
    $insert = $null
    $shape = $null
    try {
        if ($null -eq $token) { throw "Image placeholder not found: $Placeholder" }
        if (-not (Test-Path -LiteralPath $ImagePath)) { throw "Image missing: $ImagePath" }
        $para = $token.Paragraphs.Item(1)
        $start = [int]$para.Range.Start
        $body = $Doc.Range($start, [int]$para.Range.End - 1)
        $body.Text = ''
        $insert = $Doc.Range($start, $start)
        $shape = $Doc.InlineShapes.AddPicture($ImagePath, $false, $true, $insert)
        $shape.LockAspectRatio = -1
        if ([double]$shape.Width -gt 440) { $shape.Width = 440 }
        try { $shape.AlternativeText = $AltText; $shape.Title = $AltText } catch {}
        $shape.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
        return [ordered]@{ start = [int]$shape.Range.Start; end = [int]$shape.Range.End; width = [double]$shape.Width; height = [double]$shape.Height }
    }
    finally { Release-ComObject $shape; Release-ComObject $insert; Release-ComObject $body; Release-ComObject $para; Release-ComObject $token }
}

function Format-SourceParagraphByPrefix {
    param([object]$Doc, [string]$Prefix)
    $pr = Find-ParagraphByPrefix -Doc $Doc -Prefix $Prefix
    try {
        if ($null -eq $pr) { return 0 }
        $pr.Style = $wdStyleNormal
        $pr.Font.Name = 'Arial'
        $pr.Font.Size = 8
        $pr.Font.Italic = -1
        $pr.ParagraphFormat.SpaceBefore = 0
        $pr.ParagraphFormat.SpaceAfter = 6
        $pr.ParagraphFormat.KeepTogether = -1
        return 1
    }
    finally { Release-ComObject $pr }
}

function Format-AllSourceParagraphs {
    param([object]$Doc)
    $search = $null
    $find = $null
    $count = 0
    try {
        $search = $Doc.Content.Duplicate
        $find = $search.Find
        $find.ClearFormatting()
        $find.Text = 'Fuente:'
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        while ($find.Execute()) {
            $para = $null
            try {
                $para = $search.Paragraphs.Item(1)
                $para.Range.Style = $wdStyleNormal
                $para.Range.Font.Name = 'Arial'
                $para.Range.Font.Size = 8
                $para.Range.Font.Italic = -1
                $para.Range.ParagraphFormat.SpaceBefore = 0
                $para.Range.ParagraphFormat.SpaceAfter = 6
                $para.Range.ParagraphFormat.KeepTogether = -1
                $count++
            }
            finally { Release-ComObject $para }
            if ($count -gt 500) { throw 'Source paragraph search guard exceeded.' }
            $next = [math]::Min([int]$Doc.Content.End, [int]$search.End + 1)
            if ($next -ge [int]$Doc.Content.End) { break }
            $search.SetRange($next, [int]$Doc.Content.End)
            $find = $search.Find
            $find.ClearFormatting()
            $find.Text = 'Fuente:'
            $find.Forward = $true
            $find.Wrap = $wdFindStop
        }
        return $count
    }
    finally { Release-ComObject $find; Release-ComObject $search }
}

function Insert-PublicSection {
    param([object]$Doc, [object]$Spec, [string]$FigureRoot)
    Write-BuildProgress 'Publicos | localizando Heading 1 de benchmark'
    $target = Find-HeadingByText -Doc $Doc -Text ([string]$Spec.insert_before_heading) -StyleId -2
    $insert = $null
    $inserted = $null
    try {
        if ($null -eq $target) { throw "Public section target heading not found: $($Spec.insert_before_heading)" }
        $block = @(
            [string]$Spec.heading,
            [string]$Spec.intro,
            '[[TABLE_PUBLIC]]',
            "Tabla [[SEQ_PUBLIC_TABLE]]: $($Spec.table_caption)"
        ) + @($Spec.after_table) + @(
            '[[IMAGE_PUBLIC]]',
            "Figura [[SEQ_PUBLIC_FIGURE]]: $($Spec.figure_caption)"
        )
        Write-BuildProgress 'Publicos | insertando bloque textual'
        $blockText = ($block -join "`r") + "`r"
        $start = [int]$target.Start
        $insert = $Doc.Range($start, $start)
        $insert.InsertBefore($blockText)
        $inserted = $Doc.Range($start, $start + $blockText.Length)
        $inserted.Style = $wdStyleNormal
        try { $inserted.ListFormat.RemoveNumbers() } catch {}
        Write-BuildProgress 'Publicos | aplicando estilo de Heading 2'
        [void](Style-ParagraphByPrefix -Doc $Doc -Prefix ([string]$Spec.heading) -Style $wdStyleHeading2)
        Write-BuildProgress 'Publicos | creando tabla'
        $table = Replace-PlaceholderWithTable -Doc $Doc -Placeholder '[[TABLE_PUBLIC]]' -Headers @($Spec.table_headers) -Rows @($Spec.table_rows) -Widths ([double[]]@($Spec.table_widths))
        Release-ComObject $table
        Write-BuildProgress 'Publicos | insertando figura'
        [void](Replace-PlaceholderWithImage -Doc $Doc -Placeholder '[[IMAGE_PUBLIC]]' -ImagePath (Join-Path $FigureRoot ([string]$Spec.figure_file)) -AltText ([string]$Spec.figure_caption))
        Write-BuildProgress 'Publicos | creando campos SEQ y bookmarks'
        Replace-TokenWithSeqField -Doc $Doc -Token '[[SEQ_PUBLIC_TABLE]]' -Category 'Tabla' -Bookmark ([string]$Spec.table_bookmark)
        Replace-TokenWithSeqField -Doc $Doc -Token '[[SEQ_PUBLIC_FIGURE]]' -Category 'Figura' -Bookmark ([string]$Spec.figure_bookmark)
        Write-BuildProgress 'Publicos | aplicando estilos de caption y fuente'
        $capTable = Find-ParagraphByPrefix -Doc $Doc -Prefix ([string]$Spec.table_caption)
        Release-ComObject $capTable
        $cap = Find-ParagraphByPrefix -Doc $Doc -Prefix ([string]$Spec.table_caption)
        if ($null -ne $cap) { $cap.Style = $wdStyleCaption }
        Release-ComObject $cap
        $figCap = Find-ParagraphByPrefix -Doc $Doc -Prefix ([string]$Spec.figure_caption)
        if ($null -ne $figCap) { $figCap.Style = $wdStyleCaption }
        Release-ComObject $figCap
        Write-BuildProgress 'Publicos | completado'
    }
    finally { Release-ComObject $inserted; Release-ComObject $insert; Release-ComObject $target }
}

function Style-CaptionContainingTitle {
    param([object]$Doc, [string]$Title)
    $range = $null
    $find = $null
    $para = $null
    try {
        $range = $Doc.Content.Duplicate
        $find = $range.Find
        $find.Text = $Title
        $find.Forward = $true
        $find.Wrap = $wdFindStop
        if ($find.Execute()) {
            $para = $range.Paragraphs.Item(1)
            $para.Range.Style = $wdStyleCaption
            $para.Range.ParagraphFormat.KeepWithNext = -1
            return 1
        }
        return 0
    }
    finally { Release-ComObject $para; Release-ComObject $find; Release-ComObject $range }
}

function Insert-FigureAfterCaption {
    param([object]$Doc, [object]$Spec, [string]$FigureRoot, [string]$TokenStem)
    $caption = Find-SeqCaptionByPrefix -Doc $Doc -Prefix ([string]$Spec.after_caption_prefix)
    $insert = $null
    try {
        if ($null -eq $caption) { throw "Anchor caption not found: $($Spec.after_caption_prefix)" }
        $imageToken = "[[IMAGE_$TokenStem]]"
        $seqToken = "[[SEQ_$TokenStem]]"
        $block = @(
            [string]$Spec.intro,
            $imageToken,
            "Figura $seqToken`: $($Spec.caption)"
        )
        $pos = [int]$caption.End - 1
        $insert = $Doc.Range($pos, $pos)
        $insert.InsertBefore("`r" + ($block -join "`r"))
        [void](Replace-PlaceholderWithImage -Doc $Doc -Placeholder $imageToken -ImagePath (Join-Path $FigureRoot ([string]$Spec.file)) -AltText ([string]$Spec.caption))
        Replace-TokenWithSeqField -Doc $Doc -Token $seqToken -Category 'Figura' -Bookmark ([string]$Spec.bookmark)
        [void](Style-CaptionContainingTitle -Doc $Doc -Title ([string]$Spec.caption))
    }
    finally { Release-ComObject $insert; Release-ComObject $caption }
}

function Replace-BenchmarkFigure {
    param([object]$Doc, [object]$Spec, [string]$FigureRoot)
    $caption = Find-SeqCaptionByTitle -Doc $Doc -Title ([string]$Spec.caption_old_title)
    $table = $null
    $range = $null
    $separator = $null
    $insert = $null
    $shape = $null
    try {
        if ($null -eq $caption) { throw 'Benchmark figure caption not found.' }
        $table = Get-ClosestTableBefore -Doc $Doc -Position ([int]$caption.Start)
        if ($null -eq $table) { throw 'Benchmark data block not found.' }
        $range = $table.Range
        $sample = Clean-Text ([string]$range.Text)
        if ($sample -notmatch 'BASELINE_NO_VR' -or $sample -notmatch 'VR_TEC_RUNNER') { throw 'Closest table is not the benchmark figure data block.' }
        $pos = [int]$range.Start
        $table.Delete()
        $separator = $Doc.Range($pos, $pos)
        $separator.InsertBefore("`r")
        $insert = $Doc.Range($pos, $pos)
        $shape = $Doc.InlineShapes.AddPicture((Join-Path $FigureRoot ([string]$Spec.file)), $false, $true, $insert)
        $shape.LockAspectRatio = -1
        if ([double]$shape.Width -gt 440) { $shape.Width = 440 }
        $shape.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
        try { $shape.AlternativeText = [string]$Spec.caption_title; $shape.Title = [string]$Spec.caption_title } catch {}
        Release-ComObject $caption
        $caption = Find-SeqCaptionByTitle -Doc $Doc -Title ([string]$Spec.caption_old_title)
        Set-CaptionTitle -Doc $Doc -CaptionRange $caption -Title ([string]$Spec.caption_title)
    }
    finally { Release-ComObject $shape; Release-ComObject $insert; Release-ComObject $separator; Release-ComObject $range; Release-ComObject $table; Release-ComObject $caption }
}

function Replace-ObjectiveTable {
    param([object]$Doc, [object]$Spec)
    $tables = $null
    $old = $null
    $oldRange = $null
    $new = $null
    $introRange = $null
    $after = $null
    $caption = $null
    $captionFields = $null
    $seqField = $null
    $seqCode = $null
    $seqResult = $null
    try {
        $tables = $Doc.Tables
        for ($i = 1; $i -le $tables.Count; $i++) {
            $candidate = $null
            $candidateRange = $null
            try {
                $candidate = $tables.Item($i)
                $candidateRange = $candidate.Range
                $sample = Clean-Text ([string]$candidateRange.Text)
                if ($sample -match 'Evaluar Velociraptor como HIDS/DFIR' -and $sample -match '379 filas CLIENT_EVENT' -and $sample -match 'Medir falsos positivos') {
                    $old = $candidate
                    $candidate = $null
                    break
                }
            }
            finally { Release-ComObject $candidateRange; Release-ComObject $candidate }
        }
        if ($null -eq $old) { throw 'Objective table not found.' }
        $oldRange = $old.Range
        $pos = [int]$oldRange.Start
        $old.Delete()
        $introRange = $Doc.Range($pos, $pos)
        $introRange.InsertBefore(([string]$Spec.intro) + "`r")
        $newPos = $pos + ([string]$Spec.intro).Length + 1
        $new = Create-TableAt -Doc $Doc -Position $newPos -Headers @($Spec.headers) -Rows @($Spec.rows) -Widths ([double[]]@($Spec.widths))
        $new.Range.Select()
        $Doc.Application.Selection.InsertCaption('Tabla', (': ' + [string]$Spec.caption_title), [Type]::Missing, 1, $false)
        $caption = Find-SeqCaptionByTitle -Doc $Doc -Title ([string]$Spec.caption_title)
        if ($null -eq $caption) { throw 'Objective caption was not created.' }
        $caption.Style = $wdStyleCaption
        $captionFields = $caption.Fields
        for ($i = 1; $i -le $captionFields.Count; $i++) {
            $candidate = $null
            $candidateCode = $null
            try {
                $candidate = $captionFields.Item($i)
                $candidateCode = $candidate.Code
                if ((Clean-Text ([string]$candidateCode.Text)) -match '^SEQ\s+Tabla\b') { $seqField = $candidate; $candidate = $null; break }
            }
            finally { Release-ComObject $candidateCode; Release-ComObject $candidate }
        }
        if ($null -eq $seqField) { throw 'Objective caption SEQ field was not found.' }
        [void]$seqField.Update()
        $seqResult = $seqField.Result.Duplicate
        if ($Doc.Bookmarks.Exists([string]$Spec.bookmark)) { $Doc.Bookmarks.Item([string]$Spec.bookmark).Delete() }
        [void]$Doc.Bookmarks.Add([string]$Spec.bookmark, $seqResult)
        $source = 'Fuente: elaboración propia a partir de los resultados experimentales consolidados.'
        $sourcePos = [int]$caption.End - 1
        $after = $Doc.Range($sourcePos, $sourcePos)
        $after.InsertBefore("`r" + $source)
        [void](Format-SourceParagraphByPrefix -Doc $Doc -Prefix $source)
    }
    finally { Release-ComObject $seqResult; Release-ComObject $seqCode; Release-ComObject $seqField; Release-ComObject $captionFields; Release-ComObject $caption; Release-ComObject $after; Release-ComObject $introRange; Release-ComObject $new; Release-ComObject $oldRange; Release-ComObject $old; Release-ComObject $tables }
}

function Insert-Discussion {
    param([object]$Doc, [object]$Spec)
    $chapter = Find-HeadingByText -Doc $Doc -Text ([string]$Spec.chapter_heading_old) -StyleId -2
    $body = $null
    $target = $null
    $insert = $null
    $inserted = $null
    try {
        if ($null -eq $chapter) { throw 'Future work chapter heading not found.' }
        $end = [int]$chapter.End - 1
        $body = $Doc.Range([int]$chapter.Start, $end)
        $body.Text = [string]$Spec.chapter_heading_new
        $target = Find-HeadingByText -Doc $Doc -Text ([string]$Spec.insert_before_heading) -StyleId -3
        if ($null -eq $target) { throw 'Discussion insertion heading not found.' }
        $block = @([string]$Spec.heading) + @($Spec.paragraphs)
        $blockText = ($block -join "`r") + "`r"
        $start = [int]$target.Start
        $insert = $Doc.Range($start, $start)
        $insert.InsertBefore($blockText)
        $inserted = $Doc.Range($start, $start + $blockText.Length)
        $inserted.Style = $wdStyleNormal
        try { $inserted.ListFormat.RemoveNumbers() } catch {}
        [void](Style-ParagraphByPrefix -Doc $Doc -Prefix ([string]$Spec.heading) -Style $wdStyleHeading2)
    }
    finally { Release-ComObject $inserted; Release-ComObject $insert; Release-ComObject $target; Release-ComObject $body; Release-ComObject $chapter }
}

function Normalize-CaptionColons {
    param([object]$Doc)
    $records = [System.Collections.Generic.List[object]]::new()
    $fields = $null
    try {
        $fields = $Doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                $codeText = Clean-Text ([string]$code.Text)
                if ($codeText -match '^SEQ\s+(Tabla|Figura)\b') {
                    $records.Add([ordered]@{ pos = [int]$field.Result.End })
                }
            }
            finally { Release-ComObject $code; Release-ComObject $field }
        }
    }
    finally { Release-ComObject $fields }
    $count = 0
    foreach ($rec in @($records | Sort-Object pos -Descending)) {
        $check = $null
        $insert = $null
        try {
            $check = $Doc.Range([int]$rec.pos, [math]::Min([int]$rec.pos + 1, [int]$Doc.Content.End))
            $ch = [string]$check.Text
            if ($ch -ne ':' -and $ch -ne "`r" -and $ch -ne "`a") {
                $insert = $Doc.Range([int]$rec.pos, [int]$rec.pos)
                try { $insert.InsertBefore(':'); $count++ } catch {}
            }
        }
        finally { Release-ComObject $insert; Release-ComObject $check }
    }
    return $count
}

function Get-SourceForCaption {
    param([string]$Category, [string]$CaptionText, [object]$Rules)
    if ($CaptionText -match '(?i)benchmark|rendimiento|CPU|RAM|escenario') { return [string]$Rules.benchmark }
    if ($CaptionText -match '(?i)Wazuh|CLIENT_EVENT|detecci[oó]n|falsos positivos|artifacts p[uú]blicos|runner|TEC-00|Sysmon|PowerShell|alerta|cobertura') { return [string]$Rules.visibility }
    if ($Category -eq 'Tabla') { return [string]$Rules.default_table }
    return [string]$Rules.default_figure
}

function Add-MissingCaptionSources {
    param([object]$Doc, [object]$Rules)
    $records = [System.Collections.Generic.List[object]]::new()
    $fields = $null
    try {
        $fields = $Doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            $para = $null
            $pr = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                $codeText = Clean-Text ([string]$code.Text)
                if ($codeText -match '^SEQ\s+(Tabla|Figura)\b') {
                    $category = $matches[1]
                    $para = $field.Result.Paragraphs.Item(1)
                    $pr = $para.Range
                    $records.Add([ordered]@{ start = [int]$pr.Start; category = $category; text = Clean-Text ([string]$pr.Text) })
                }
            }
            finally { Release-ComObject $pr; Release-ComObject $para; Release-ComObject $code; Release-ComObject $field }
        }
    }
    finally { Release-ComObject $fields }
    $added = 0
    foreach ($rec in @($records | Sort-Object start -Descending)) {
        $anchor = $null
        $captionPara = $null
        $captionRange = $null
        $probe = $null
        $nextPara = $null
        $insert = $null
        $sourceRange = $null
        try {
            $anchor = $Doc.Range([int]$rec.start, [math]::Min([int]$rec.start + 1, [int]$Doc.Content.End))
            $captionPara = $anchor.Paragraphs.Item(1)
            $captionRange = $captionPara.Range
            $captionRaw = [string]$captionRange.Text
            if ((Clean-Text $captionRaw) -match 'Fuente:') { continue }
            $probeEnd = [math]::Min([int]$Doc.Content.End, [int]$captionRange.End + 500)
            $probe = $Doc.Range([int]$captionRange.End, $probeEnd)
            if ($probe.Paragraphs.Count -gt 0) {
                $nextPara = $probe.Paragraphs.Item(1)
                $nextText = Clean-Text ([string]$nextPara.Range.Text)
                if ($nextText -match '^Fuente:') {
                    $nextPara.Range.Style = $wdStyleNormal
                    $nextPara.Range.Font.Name = 'Arial'
                    $nextPara.Range.Font.Size = 8
                    $nextPara.Range.Font.Italic = -1
                    $nextPara.Range.ParagraphFormat.SpaceBefore = 0
                    $nextPara.Range.ParagraphFormat.SpaceAfter = 6
                    $nextPara.Range.ParagraphFormat.KeepTogether = -1
                    continue
                }
            }
            $source = Get-SourceForCaption -Category ([string]$rec.category) -CaptionText ([string]$rec.text) -Rules $Rules
            $splitPos = [int]$captionRange.End - 1
            if ($captionRaw.EndsWith("`r`a")) { $splitPos = [int]$captionRange.End - 2 }
            $insert = $Doc.Range($splitPos, $splitPos)
            $insert.InsertBefore("`r" + $source)
            $sourceStart = $splitPos + 1
            $sourceRange = $Doc.Range($sourceStart, $sourceStart + $source.Length)
            $sourceRange.Style = $wdStyleNormal
            $sourceRange.Font.Name = 'Arial'
            $sourceRange.Font.Size = 8
            $sourceRange.Font.Italic = -1
            $sourceRange.ParagraphFormat.SpaceBefore = 0
            $sourceRange.ParagraphFormat.SpaceAfter = 6
            $sourceRange.ParagraphFormat.KeepTogether = -1
            $added++
        }
        finally { Release-ComObject $sourceRange; Release-ComObject $insert; Release-ComObject $nextPara; Release-ComObject $probe; Release-ComObject $captionRange; Release-ComObject $captionPara; Release-ComObject $anchor }
    }
    return [ordered]@{ total_caption_fields = $records.Count; sources_added = $added }
}

function Normalize-Layout {
    param([object]$Doc)
    $tables = $null
    try {
        $tables = $Doc.Tables
        for ($i = 1; $i -le $tables.Count; $i++) {
            $table = $null
            try {
                $table = $tables.Item($i)
                try { $table.Rows.AllowBreakAcrossPages = 0 } catch {}
                try { $table.Rows.HeightRule = $wdRowHeightAuto } catch {}
                if ($table.Rows.Count -gt 1) { try { $table.Rows.Item(1).HeadingFormat = -1 } catch {} }
                $paragraphs = $null
                try {
                    $paragraphs = $table.Range.Paragraphs
                    for ($pIndex = 1; $pIndex -le $paragraphs.Count; $pIndex++) {
                        $paragraph = $null
                        $paragraphRange = $null
                        try {
                            $paragraph = $paragraphs.Item($pIndex)
                            $paragraphRange = $paragraph.Range
                            $styleName = Get-StyleName $paragraphRange
                            if ($styleName -match '^(Heading|T[ií]tulo)\s*[1-9]') {
                                $paragraphRange.Style = $wdStyleNormal
                                try { $paragraphRange.ListFormat.RemoveNumbers() } catch {}
                            }
                        }
                        finally { Release-ComObject $paragraphRange; Release-ComObject $paragraph }
                    }
                }
                finally { Release-ComObject $paragraphs }
            }
            finally { Release-ComObject $table }
        }
        foreach ($styleId in @(-2, -3, -4, -5, -6, -7, -8, -9, -10, $wdStyleCaption)) {
            $style = $null
            try {
                $style = $Doc.Styles.Item($styleId)
                $style.ParagraphFormat.KeepWithNext = -1
                $style.ParagraphFormat.KeepTogether = -1
                $style.ParagraphFormat.WidowControl = -1
            }
            catch {}
            finally { Release-ComObject $style }
        }
    }
    finally { Release-ComObject $tables }
}

function Normalize-VersioningTableStyles {
    param([object]$Doc)
    $caption = Find-SeqCaptionByPrefix -Doc $Doc -Prefix 'Tabla 32: Ciclo de versionado'
    $table = $null
    $paragraphs = $null
    try {
        if ($null -eq $caption) { return 0 }
        $table = Get-ClosestTableBefore -Doc $Doc -Position ([int]$caption.Start)
        if ($null -eq $table) { return 0 }
        $paragraphs = $table.Range.Paragraphs
        for ($i = 1; $i -le $paragraphs.Count; $i++) {
            $p = $null
            try { $p = $paragraphs.Item($i); $p.Range.Style = $wdStyleNormal }
            finally { Release-ComObject $p }
        }
        return $paragraphs.Count
    }
    finally { Release-ComObject $paragraphs; Release-ComObject $table; Release-ComObject $caption }
}

function Update-AllFields {
    param([object]$Doc)
    $updated = 0
    for ($storyType = 1; $storyType -le 17; $storyType++) {
        $story = $null
        try { $story = $Doc.StoryRanges.Item($storyType) } catch { $story = $null }
        $guard = 0
        while ($null -ne $story) {
            $fields = $null
            $next = $null
            try {
                $fields = $story.Fields
                $fieldCount = [int]$fields.Count
                if ($fieldCount -gt 0) { try { [void]$fields.Update(); $updated += $fieldCount } catch {} }
                try { $next = $story.NextStoryRange } catch { $next = $null }
            }
            finally { Release-ComObject $fields; Release-ComObject $story }
            $story = $next
            $guard++
            if ($guard -gt 50) { throw "Story range guard exceeded for story $storyType" }
        }
    }
    $tocs = $null
    $tofs = $null
    try {
        $tocs = $Doc.TablesOfContents
        for ($i = 1; $i -le $tocs.Count; $i++) {
            $toc = $null
            try { $toc = $tocs.Item($i); $toc.Update(); $toc.UpdatePageNumbers() } finally { Release-ComObject $toc }
        }
        $tofs = $Doc.TablesOfFigures
        for ($i = 1; $i -le $tofs.Count; $i++) {
            $tof = $null
            try { $tof = $tofs.Item($i); $tof.Update() } finally { Release-ComObject $tof }
        }
    }
    finally { Release-ComObject $tofs; Release-ComObject $tocs }
    return $updated
}

function Get-SeqCounts {
    param([object]$Doc)
    $counts = [ordered]@{ Tabla = 0; Figura = 0; Codigo = 0 }
    $fields = $null
    try {
        $fields = $Doc.Fields
        for ($i = 1; $i -le $fields.Count; $i++) {
            $field = $null
            $code = $null
            try {
                $field = $fields.Item($i)
                $code = $field.Code
                $text = Clean-Text ([string]$code.Text)
                if ($text -match '^SEQ\s+Tabla\b') { $counts.Tabla++ }
                elseif ($text -match '^SEQ\s+Figura\b') { $counts.Figura++ }
                elseif ($text -match '^SEQ\s+C[oó]digo\b') { $counts.Codigo++ }
            }
            finally { Release-ComObject $code; Release-ComObject $field }
        }
    }
    finally { Release-ComObject $fields }
    return $counts
}

function Scan-SuspiciousText {
    param([object]$Doc)
    $text = [string]$Doc.Content.Text
    $patterns = @('Error! Reference source not found', '¡Error! No se encuentra el origen de la referencia', 'Error! Bookmark not defined', 'TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026.xlsx', 'TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026.xlsx', '09_Benchmark_Plan', 'v15.21', '[[REF:', '[[SEQ_', '[[TABLE_', '[[IMAGE_')
    $hits = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $patterns) { if ($text.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hits.Add($p) } }
    if ($text -cmatch '\bTODO\b') { $hits.Add('TODO') }
    if ($text -cmatch '\bTBD\b') { $hits.Add('TBD') }
    return @($hits)
}

function Prepare-OutputFile {
    param([string]$Path, [string]$WorkingSource)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $Path) {
        $archive = Join-Path (Split-Path -Parent $WorkingSource) ("previous_output_" + (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '.docx')
        Move-Item -LiteralPath $Path -Destination $archive
    }
}

$workingFull = (Resolve-Path -LiteralPath $WorkingDoc).Path
$manifestFull = (Resolve-Path -LiteralPath $ManifestPath).Path
$figuresFull = (Resolve-Path -LiteralPath $FiguresDir).Path
$outputFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDoc))
$reportFull = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ReportPath))
$script:ProgressPath = Join-Path (Split-Path -Parent $reportFull) ("word_build_" + $Phase + '_progress.log')
[IO.File]::WriteAllText($script:ProgressPath, "INICIO $(Get-Date -Format o)`r`n", $utf8NoBom)
$manifest = Get-Content -LiteralPath $manifestFull -Raw -Encoding UTF8 | ConvertFrom-Json

$report = [ordered]@{
    generated_at = (Get-Date).ToString('o')
    phase = $Phase
    working_doc = $workingFull
    output_doc = $outputFull
    word_version = $null
    word_build = $null
    paragraph_replacements = @()
    literal_replacements = @()
    table_replacements = @()
    added_sections = @()
    references_inserted = 0
    normalized_caption_colons = 0
    normalized_versioning_cells = 0
    caption_sources = $null
    fields_updated = 0
    final_stats = $null
    suspicious_text = @()
    source_sha256_before = (Get-FileHash -Algorithm SHA256 -LiteralPath $workingFull).Hash
    output_sha256 = $null
    forced_process_cleanup = @()
}

$baselinePids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$documents = $null
$doc = $null

try {
    Write-BuildProgress 'Creando Word.Application'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}
    try { $word.Options.Pagination = $false } catch {}
    try { $word.Options.CheckSpellingAsYouType = $false } catch {}
    try { $word.Options.CheckGrammarAsYouType = $false } catch {}
    $report.word_version = [string]$word.Version
    $report.word_build = [string]$word.Build
    $documents = $word.Documents
    Write-BuildProgress 'Abriendo copia temporal sin reparacion'
    $doc = $documents.OpenNoRepairDialog($workingFull, $false, $false, $false)
    if ($null -eq $doc) { throw 'Word returned a null document.' }
    $doc.TrackRevisions = $false
    try { if ($doc.Revisions.Count -gt 0) { $doc.Revisions.AcceptAll() } } catch {}
    try { for ($i = $doc.Comments.Count; $i -ge 1; $i--) { $doc.Comments.Item($i).Delete() } } catch {}

    try { $doc.BuiltInDocumentProperties.Item('Title').Value = [string]$manifest.document_metadata.title } catch {}
    try { $doc.BuiltInDocumentProperties.Item('Subject').Value = [string]$manifest.document_metadata.subject } catch {}
    try { $doc.BuiltInDocumentProperties.Item('Author').Value = [string]$manifest.document_metadata.author } catch {}
    try { $doc.BuiltInDocumentProperties.Item('Keywords').Value = [string]$manifest.document_metadata.keywords } catch {}

    if ($Phase -in @('all', 'phase1')) {
        Write-BuildProgress 'Aplicando reemplazos literales'
        foreach ($item in @($manifest.literal_replacements)) {
            $result = Replace-LiteralAll -Doc $doc -Old ([string]$item.old) -New ([string]$item.new)
            $report.literal_replacements += [ordered]@{ old = [string]$item.old; new = [string]$item.new; attempted = $result }
        }
        Write-BuildProgress 'Aplicando reemplazos de parrafos'
        foreach ($item in @($manifest.paragraph_replacements)) {
            $result = Replace-ParagraphByPrefix -Doc $doc -Prefix ([string]$item.prefix) -Text ([string]$item.text)
            $report.paragraph_replacements += [ordered]@{ prefix = [string]$item.prefix; replaced = $result }
            if ($result -ne 1) { throw "Paragraph replacement failed: $($item.prefix)" }
        }

        Write-BuildProgress 'Normalizando tabla de versionado'
        $report.normalized_versioning_cells = Normalize-VersioningTableStyles -Doc $doc
        Write-BuildProgress 'Sustituyendo tablas canonicas'
        foreach ($spec in @($manifest.table_replacements)) {
            Replace-TableByCaption -Doc $doc -Spec $spec
            $report.table_replacements += [string]$spec.caption_prefix
        }
        Write-BuildProgress 'Sustituyendo tabla de objetivos'
        Replace-ObjectiveTable -Doc $doc -Spec $manifest.objective_table
    }

    if ($Phase -eq 'phase1') {
        Write-BuildProgress 'Guardando checkpoint de fase 1'
        Prepare-OutputFile -Path $outputFull -WorkingSource $workingFull
        $doc.SaveAs2($outputFull, $wdFormatDocumentDefault)
        $doc.Close($wdDoNotSaveChanges)
        Release-ComObject $doc; $doc = $null
        $word.Quit()
        Release-ComObject $documents; $documents = $null
        Release-ComObject $word; $word = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $report.output_sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFull).Hash
        Write-BuildProgress 'FASE 1 COMPLETADA'
        return
    }

    if ($Phase -in @('all', 'phase2')) {
        Write-BuildProgress 'Insertando resultados de artifacts publicos'
        Insert-PublicSection -Doc $doc -Spec $manifest.public_section -FigureRoot $figuresFull
        Write-BuildProgress 'Insertando figura de alertas custom'
        Insert-FigureAfterCaption -Doc $doc -Spec $manifest.custom_figure -FigureRoot $figuresFull -TokenStem 'CUSTOM_ALERTS'
        Write-BuildProgress 'Insertando figura de cobertura'
        Insert-FigureAfterCaption -Doc $doc -Spec $manifest.coverage_figure -FigureRoot $figuresFull -TokenStem 'COVERAGE'
        Write-BuildProgress 'Sustituyendo figura de benchmark'
        Replace-BenchmarkFigure -Doc $doc -Spec $manifest.benchmark_figure -FigureRoot $figuresFull
        Write-BuildProgress 'Insertando discusion integrada'
        Insert-Discussion -Doc $doc -Spec $manifest.discussion
        $report.added_sections = @('public_artifacts_results', 'integrated_discussion', 'custom_alerts_figure', 'coverage_figure', 'benchmark_figure')

        Write-BuildProgress 'Actualizando referencias internas'
        $report.normalized_caption_colons = 0
        $report.references_inserted = Replace-RefTokens -Doc $doc
        Write-BuildProgress 'Fuentes de captions aplazadas a la correccion Open XML controlada'
        $report.caption_sources = [ordered]@{ total_caption_fields = 0; sources_added = 0 }
        Write-BuildProgress 'Normalizando paginacion de titulos y tablas'
        Normalize-Layout -Doc $doc
    }

    if ($Phase -eq 'phase2') {
        Write-BuildProgress 'Guardando checkpoint de fase 2'
        Prepare-OutputFile -Path $outputFull -WorkingSource $workingFull
        $doc.SaveAs2($outputFull, $wdFormatDocumentDefault)
        $doc.Close($wdDoNotSaveChanges)
        Release-ComObject $doc; $doc = $null
        $word.Quit()
        Release-ComObject $documents; $documents = $null
        Release-ComObject $word; $word = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $report.output_sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFull).Hash
        Write-BuildProgress 'FASE 2 COMPLETADA'
        return
    }

    Write-BuildProgress 'Actualizando campos e indices - pase 1'
    try { $word.Options.Pagination = $true } catch {}
    $report.fields_updated = Update-AllFields -Doc $doc
    Write-BuildProgress 'Repaginando - pase 1'
    [void]$doc.Repaginate()
    Write-BuildProgress 'Actualizando campos e indices - pase 2'
    $report.fields_updated += Update-AllFields -Doc $doc
    Write-BuildProgress 'Repaginando - pase 2'
    [void]$doc.Repaginate()

    Write-BuildProgress 'Escaneando errores y recopilando estadisticas'
    $report.suspicious_text = @(Scan-SuspiciousText -Doc $doc)
    if ($report.suspicious_text.Count -gt 0) { throw "Suspicious unresolved text: $($report.suspicious_text -join ', ')" }

    $seq = Get-SeqCounts -Doc $doc
    $report.final_stats = [ordered]@{
        pages = [int]$doc.ComputeStatistics($wdStatisticPages)
        words = [int]$doc.ComputeStatistics($wdStatisticWords)
        characters_with_spaces = [int]$doc.ComputeStatistics($wdStatisticCharactersWithSpaces)
        word_tables = [int]$doc.Tables.Count
        inline_shapes = [int]$doc.InlineShapes.Count
        floating_shapes = [int]$doc.Shapes.Count
        sections = [int]$doc.Sections.Count
        fields = [int]$doc.Fields.Count
        tables_of_contents = [int]$doc.TablesOfContents.Count
        tables_of_figures = [int]$doc.TablesOfFigures.Count
        bookmarks = [int]$doc.Bookmarks.Count
        comments = [int]$doc.Comments.Count
        revisions = [int]$doc.Revisions.Count
        seq_tables = [int]$seq.Tabla
        seq_figures = [int]$seq.Figura
        seq_code = [int]$seq.Codigo
    }

    Prepare-OutputFile -Path $outputFull -WorkingSource $workingFull
    Write-BuildProgress 'Guardando DOCX final'
    $doc.SaveAs2($outputFull, $wdFormatDocumentDefault)
    Write-BuildProgress 'Cerrando documento y Word'
    $doc.Close($wdDoNotSaveChanges)
    Release-ComObject $doc
    $doc = $null
    $word.Quit()
    Release-ComObject $documents
    $documents = $null
    Release-ComObject $word
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (-not (Test-Path -LiteralPath $outputFull)) { throw 'Final DOCX was not created.' }
    $report.output_sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFull).Hash
    Write-BuildProgress 'COMPLETADO'
}
catch {
    $report.error = $_.Exception.ToString()
    Write-BuildProgress ("ERROR: " + $_.Exception.Message)
    throw
}
finally {
    if ($null -ne $doc) { try { $doc.Close($wdDoNotSaveChanges) } catch {}; Release-ComObject $doc }
    if ($null -ne $word) { try { $word.Quit() } catch {}; Release-ComObject $word }
    Release-ComObject $documents
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    $afterPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids = @($afterPids | Where-Object { $_ -notin $baselinePids })
    foreach ($pidValue in $newPids) {
        try { Stop-Process -Id $pidValue -Force -ErrorAction Stop; $report.forced_process_cleanup += $pidValue } catch {}
    }
    [IO.File]::WriteAllText($reportFull, ($report | ConvertTo-Json -Depth 12), $utf8NoBom)
}

$report | ConvertTo-Json -Depth 12
