[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DataPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$outFull = (Resolve-Path -LiteralPath $OutputDir).Path
$data = Get-Content -Encoding UTF8 -Raw -LiteralPath (Resolve-Path -LiteralPath $DataPath).Path | ConvertFrom-Json

function New-BaseChart {
    param(
        [string]$Title,
        [int]$Width = 1800,
        [int]$Height = 1000
    )
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = $Width
    $chart.Height = $Height
    $chart.BackColor = [Drawing.Color]::White
    $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All
    $chart.TextAntiAliasingQuality = [System.Windows.Forms.DataVisualization.Charting.TextAntiAliasingQuality]::High
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 'Principal'
    $area.BackColor = [Drawing.Color]::White
    $area.AxisX.MajorGrid.Enabled = $false
    $area.AxisY.MajorGrid.LineColor = [Drawing.Color]::FromArgb(225, 230, 235)
    $area.AxisY.MajorGrid.LineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Dot
    $area.AxisX.LabelStyle.Font = New-Object Drawing.Font('Arial', 18, [Drawing.FontStyle]::Regular)
    $area.AxisY.LabelStyle.Font = New-Object Drawing.Font('Arial', 18, [Drawing.FontStyle]::Regular)
    $area.AxisX.TitleFont = New-Object Drawing.Font('Arial', 20, [Drawing.FontStyle]::Bold)
    $area.AxisY.TitleFont = New-Object Drawing.Font('Arial', 20, [Drawing.FontStyle]::Bold)
    $area.AxisX.LineColor = [Drawing.Color]::FromArgb(90, 100, 110)
    $area.AxisY.LineColor = [Drawing.Color]::FromArgb(90, 100, 110)
    [void]$chart.ChartAreas.Add($area)
    $titleObj = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $titleObj.Text = $Title
    $titleObj.Font = New-Object Drawing.Font('Arial', 27, [Drawing.FontStyle]::Bold)
    $titleObj.ForeColor = [Drawing.Color]::FromArgb(31, 41, 55)
    [void]$chart.Titles.Add($titleObj)
    return $chart
}

function Add-ColumnSeries {
    param(
        [object]$Chart,
        [string]$Name,
        [string[]]$Categories,
        [double[]]$Values,
        [Drawing.Color]$Color,
        [string]$LabelFormat = '0'
    )
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series $Name
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $series.Color = $Color
    $series.BorderColor = [Drawing.Color]::FromArgb(70, 80, 90)
    $series.BorderWidth = 1
    $series.IsValueShownAsLabel = $true
    $series.LabelFormat = $LabelFormat
    $series.Font = New-Object Drawing.Font('Arial', 17, [Drawing.FontStyle]::Bold)
    $series.LabelForeColor = [Drawing.Color]::FromArgb(31, 41, 55)
    $series['PointWidth'] = '0.62'
    for ($i = 0; $i -lt $Categories.Count; $i++) {
        $pointIndex = $series.Points.AddXY($Categories[$i], $Values[$i])
        $series.Points[$pointIndex].AxisLabel = $Categories[$i]
    }
    [void]$Chart.Series.Add($series)
    return $series
}

function Save-ChartPng {
    param([object]$Chart, [string]$Path)
    $Chart.SaveImage($Path, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
    $Chart.Dispose()
}

$titleDistribution = 'Distribuci' + [char]0x00F3 + 'n de alertas CLIENT_EVENT por perfil'
$titleDetection = 'Detecci' + [char]0x00F3 + 'n espec' + [char]0x00ED + 'fica acreditada por soluci' + [char]0x00F3 + 'n'
$axisTechniques = 'T' + [char]0x00E9 + 'cnicas detectadas (sobre 9)'
$titlePublic = 'Resultado de las seis campa' + [char]0x00F1 + 'as con artifacts p' + [char]0x00FA + 'blicos'
$axisCampaigns = 'Campa' + [char]0x00F1 + 'as'

$customChart = New-BaseChart -Title $titleDistribution
$customChart.ChartAreas['Principal'].AxisY.Title = 'Alertas CLIENT_EVENT'
$customChart.ChartAreas['Principal'].AxisY.Minimum = 0
$customChart.ChartAreas['Principal'].AxisY.Maximum = 150
$customChart.ChartAreas['Principal'].AxisY.Interval = 25
$profileNames = @('P1 Critical', 'P2 High', 'P3 Medium', 'P4 Low')
$profileValues = @(
    [double]$data.velociraptor_custom.profiles.P1,
    [double]$data.velociraptor_custom.profiles.P2,
    [double]$data.velociraptor_custom.profiles.P3,
    [double]$data.velociraptor_custom.profiles.P4
)
$customSeries = Add-ColumnSeries -Chart $customChart -Name 'Alertas' -Categories $profileNames -Values $profileValues -Color ([Drawing.Color]::FromArgb(37, 99, 235))
$colors = @(
    [Drawing.Color]::FromArgb(185, 28, 28),
    [Drawing.Color]::FromArgb(217, 119, 6),
    [Drawing.Color]::FromArgb(37, 99, 235),
    [Drawing.Color]::FromArgb(71, 85, 105)
)
for ($i = 0; $i -lt $customSeries.Points.Count; $i++) { $customSeries.Points[$i].Color = $colors[$i] }
Save-ChartPng -Chart $customChart -Path (Join-Path $outFull 'figure_custom_alerts.png')

$coverageChart = New-BaseChart -Title $titleDetection
$coverageChart.ChartAreas['Principal'].AxisY.Title = $axisTechniques
$coverageChart.ChartAreas['Principal'].AxisY.Minimum = 0
$coverageChart.ChartAreas['Principal'].AxisY.Maximum = 9.8
$coverageChart.ChartAreas['Principal'].AxisY.Interval = 1
$coverageNames = @('Velociraptor custom', 'Hayabusa CH', 'Wazuh base', 'Wazuh custom')
$coverageValues = @(
    [double]$data.velociraptor_custom.techniques_detected,
    [double]$data.public_artifacts.hayabusa_ch_detected,
    [double]$data.wazuh.base_detected,
    [double]$data.wazuh.custom_detected
)
$coverageSeries = Add-ColumnSeries -Chart $coverageChart -Name 'Detección específica' -Categories $coverageNames -Values $coverageValues -Color ([Drawing.Color]::FromArgb(14, 116, 144))
$coverageColors = @(
    [Drawing.Color]::FromArgb(14, 116, 144),
    [Drawing.Color]::FromArgb(124, 58, 237),
    [Drawing.Color]::FromArgb(71, 85, 105),
    [Drawing.Color]::FromArgb(5, 150, 105)
)
for ($i = 0; $i -lt $coverageSeries.Points.Count; $i++) { $coverageSeries.Points[$i].Color = $coverageColors[$i] }
Save-ChartPng -Chart $coverageChart -Path (Join-Path $outFull 'figure_detection_coverage.png')

$publicChart = New-BaseChart -Title $titlePublic
$publicChart.ChartAreas['Principal'].AxisY.Title = $axisCampaigns
$publicChart.ChartAreas['Principal'].AxisY.Minimum = 0
$publicChart.ChartAreas['Principal'].AxisY.Maximum = 6
$publicChart.ChartAreas['Principal'].AxisY.Interval = 1
$publicSeries = Add-ColumnSeries -Chart $publicChart -Name 'Campañas' -Categories @('Positivas', 'No concluyente') -Values @([double]$data.public_artifacts.positive, [double]$data.public_artifacts.inconclusive) -Color ([Drawing.Color]::FromArgb(5, 150, 105))
$publicSeries.Points[0].Color = [Drawing.Color]::FromArgb(5, 150, 105)
$publicSeries.Points[1].Color = [Drawing.Color]::FromArgb(217, 119, 6)
Save-ChartPng -Chart $publicChart -Path (Join-Path $outFull 'figure_public_campaigns.png')

$cpuTemp = Join-Path $outFull '_benchmark_cpu.png'
$ramTemp = Join-Path $outFull '_benchmark_ram.png'
$scenarioNames = @('BASELINE_NO_VR', 'VR_IDLE', 'VR_TEC_RUNNER')

$cpuChart = New-BaseChart -Title 'CPU del cliente Velociraptor' -Width 1800 -Height 850
$cpuChart.ChartAreas['Principal'].AxisY.Title = 'CPU (%)'
$cpuChart.ChartAreas['Principal'].AxisY.Minimum = 0
$cpuChart.ChartAreas['Principal'].AxisY.Maximum = 60
$cpuChart.ChartAreas['Principal'].AxisY.Interval = 10
$cpuAvg = @(); $cpuMax = @()
foreach ($s in $data.benchmark.scenarios) { $cpuAvg += [double]$s.avg_cpu_percent; $cpuMax += [double]$s.max_cpu_percent }
[void](Add-ColumnSeries -Chart $cpuChart -Name 'Media' -Categories $scenarioNames -Values $cpuAvg -Color ([Drawing.Color]::FromArgb(37, 99, 235)) -LabelFormat '0.##')
[void](Add-ColumnSeries -Chart $cpuChart -Name 'Pico' -Categories $scenarioNames -Values $cpuMax -Color ([Drawing.Color]::FromArgb(185, 28, 28)) -LabelFormat '0.##')
$cpuChart.Legends.Add('Leyenda') | Out-Null
$cpuChart.Legends['Leyenda'].Font = New-Object Drawing.Font('Arial', 17)
Save-ChartPng -Chart $cpuChart -Path $cpuTemp

$ramChart = New-BaseChart -Title 'Memoria del cliente Velociraptor' -Width 1800 -Height 850
$ramChart.ChartAreas['Principal'].AxisY.Title = 'RAM (MB)'
$ramChart.ChartAreas['Principal'].AxisY.Minimum = 0
$ramChart.ChartAreas['Principal'].AxisY.Maximum = 65
$ramChart.ChartAreas['Principal'].AxisY.Interval = 10
$ramAvg = @(); $ramMax = @()
foreach ($s in $data.benchmark.scenarios) { $ramAvg += [double]$s.avg_ram_mb; $ramMax += [double]$s.max_ram_mb }
[void](Add-ColumnSeries -Chart $ramChart -Name 'Media' -Categories $scenarioNames -Values $ramAvg -Color ([Drawing.Color]::FromArgb(5, 150, 105)) -LabelFormat '0.##')
[void](Add-ColumnSeries -Chart $ramChart -Name 'Pico' -Categories $scenarioNames -Values $ramMax -Color ([Drawing.Color]::FromArgb(217, 119, 6)) -LabelFormat '0.##')
$ramChart.Legends.Add('Leyenda') | Out-Null
$ramChart.Legends['Leyenda'].Font = New-Object Drawing.Font('Arial', 17)
Save-ChartPng -Chart $ramChart -Path $ramTemp

$cpuBitmap = New-Object Drawing.Bitmap($cpuTemp)
$ramBitmap = New-Object Drawing.Bitmap($ramTemp)
$combined = New-Object Drawing.Bitmap(1800, 1700)
$graphics = [Drawing.Graphics]::FromImage($combined)
try {
    $graphics.Clear([Drawing.Color]::White)
    $graphics.DrawImage($cpuBitmap, 0, 0, 1800, 850)
    $graphics.DrawImage($ramBitmap, 0, 850, 1800, 850)
    $combined.Save((Join-Path $outFull 'figure_benchmark_cost.png'), [Drawing.Imaging.ImageFormat]::Png)
}
finally {
    $graphics.Dispose()
    $combined.Dispose()
    $cpuBitmap.Dispose()
    $ramBitmap.Dispose()
}
Remove-Item -LiteralPath $cpuTemp, $ramTemp -Force

Get-ChildItem -LiteralPath $outFull -Filter 'figure_*.png' | Sort-Object Name | Select-Object Name, Length, FullName
