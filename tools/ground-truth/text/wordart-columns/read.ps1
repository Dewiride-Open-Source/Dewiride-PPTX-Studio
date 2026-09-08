# Experiment T11, step 2 - ask PowerPoint how wide it made each WordArt column.
#
#   powershell -File tools/ground-truth/text/wordart-columns/read.ps1 -Dir <work-dir>
#
# Two instruments. The EMF drawing stream carries the exact pen, the face GDI was
# asked for and the per-character advances, in logical units and with no
# rasteriser in the way. The bitmap carries the one thing a stream cannot say:
# which way the outline actually points. ADR 0040.
#
# Opens with OpenAndRepair:=msoFalse so a package PowerPoint refuses stays
# distinct from one it rewrote. Read-only; never quits a PowerPoint it did not
# start, and writes only inside -Dir.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [int]$RasterWidth = 3840,
    [int]$RasterHeight = 2160
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'column-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no column-inputs.json in $root - run build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$rasterSlides = @{}
foreach ($p in $inputs.probes) { if ($p.raster) { $rasterSlides[[int]$p.slide] = $true } }

$emfDir = Join-Path $root 'emf'
$bmpDir = Join-Path $root 'bmp'
New-Item -ItemType Directory -Force -Path $emfDir | Out-Null
New-Item -ItemType Directory -Force -Path $bmpDir | Out-Null

$created = $false
$app = $null
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
}
catch {
    $app = New-Object -ComObject PowerPoint.Application
    $created = $true
}

$app.DisplayAlerts = $ppAlertsNone
$app.AutomationSecurity = $msoAutomationSecurityForceDisable

$file = Join-Path $root 'wordart-columns.pptx'
$record = [ordered]@{
    file        = 'wordart-columns.pptx'
    opened      = $false
    repaired    = $null
    error       = $null
    rasterWidth = $RasterWidth
    rasterHeight= $RasterHeight
    slides      = @()
    shapes      = @()
}

$pres = $null
try {
    # FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair
    $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoFalse)
    $record.opened = $true
    $record.repaired = $false
}
catch {
    $record.error = $_.Exception.Message
    try {
        $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoTrue)
        $record.opened = $true
        $record.repaired = $true
    }
    catch { $record.opened = $false }
}

if ($record.opened) {
    try {
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $entry = [ordered]@{ slide = $i; emf = $null; bmp = $null; error = $null }

            $emfName = ("s{0}.emf" -f $i)
            $emfPath = Join-Path $emfDir $emfName
            if (Test-Path -LiteralPath $emfPath) { Remove-Item -LiteralPath $emfPath -Force }
            try {
                $slide.Export($emfPath, 'EMF', $RasterWidth, $RasterHeight)
                $entry.emf = "emf/$emfName"
            }
            catch { $entry.error = $_.Exception.Message }

            if ($rasterSlides.ContainsKey($i)) {
                $bmpName = ("s{0}.bmp" -f $i)
                $bmpPath = Join-Path $bmpDir $bmpName
                if (Test-Path -LiteralPath $bmpPath) { Remove-Item -LiteralPath $bmpPath -Force }
                try {
                    $slide.Export($bmpPath, 'BMP', $RasterWidth, $RasterHeight)
                    $entry.bmp = "bmp/$bmpName"
                }
                catch { $entry.error = $_.Exception.Message }
            }

            $record.slides += $entry

            # The object model's own reading of each frame, which is what T6
            # measured; kept so the two experiments can be checked against
            # each other rather than trusted separately.
            foreach ($sh in $slide.Shapes) {
                $row = [ordered]@{
                    id = [string]$sh.Name; slide = $i
                    left = $null; top = $null
                    boundLeft = $null; boundTop = $null; boundWidth = $null; boundHeight = $null
                    orientation = $null
                }
                try { $row.left = [double]$sh.Left } catch {}
                try { $row.top = [double]$sh.Top } catch {}
                try { $row.orientation = [int]$sh.TextFrame2.Orientation } catch {}
                try {
                    $range = $sh.TextFrame2.TextRange
                    $row.boundLeft = [double]$range.BoundLeft
                    $row.boundTop = [double]$range.BoundTop
                    $row.boundWidth = [double]$range.BoundWidth
                    $row.boundHeight = [double]$range.BoundHeight
                }
                catch {}
                $record.shapes += $row
            }
        }
    }
    finally { try { $pres.Close() } catch {} }
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = $record | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText(
    (Join-Path $root 'column-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))

$state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
Write-Host ("wordart-columns {0} {1} slide(s) {2} shape(s)" -f $state, $record.slides.Count, $record.shapes.Count)
