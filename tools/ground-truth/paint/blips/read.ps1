# Experiment C6, step 2 - ask PowerPoint what it painted.
#
#   powershell -File tools/ground-truth/paint/blips/read.ps1 -Dir <work-dir>
#
# Reads `blip-inputs.json`, opens every deck it names, and exports each slide as
# a BMP. The BMP is the oracle: `tools/ground-truth/lib/bmp.ts` reads it with no
# decompressor, so nothing sits between what PowerPoint painted and what
# `analyse.ts` compares against.
#
# ## Two resolutions, for the same reason C3 used three
#
# A tile's natural size is either a fixed number of device pixels or a fixed size
# in shape units, and the two are indistinguishable at one resolution. Exported
# at 1920 and 3840 wide, a tile period that doubles is in shape units and one
# that holds is in pixels. `a:blipFill/@dpi` is the attribute that would decide
# it, and every deck here writes `dpi="0"`.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, because `Open` silently repairs and
# a repaired deck answers for markup we did not write.
#
# Read-only. Nothing is written back to any deck. Attaches to a running
# PowerPoint if there is one and never quits one it did not start.
param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'blip-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no blip-inputs.json in $root - run tools/ground-truth/paint/blips/build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

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

$shots = Join-Path $root 'shots'
if (-not (Test-Path -LiteralPath $shots)) { New-Item -ItemType Directory -Path $shots | Out-Null }

$version = $app.Version
$build = $app.Build
$records = @()

foreach ($deck in $inputs.decks) {
    $path = Join-Path $root $deck.file
    if (-not (Test-Path -LiteralPath $path)) { throw "no such deck: $path" }

    $pres = $null
    try {
        $pres = $app.Presentations.Open2007($path, $msoTrue, $msoFalse, $msoFalse, $msoFalse)
    }
    catch {
        $records += [ordered]@{ deck = $deck.id; opened = $false; error = $_.Exception.Message; slides = @() }
        continue
    }

    $slides = @()
    foreach ($slide in $pres.Slides) {
        $index = $slide.SlideIndex
        $exports = @()
        foreach ($width in $deck.widths) {
            $height = [int][Math]::Round($width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
            $name = "{0}-{1:d2}-{2}.bmp" -f $deck.id, $index, $width
            $out = Join-Path $shots $name
            $slide.Export($out, 'BMP', $width, $height)
            $exports += [ordered]@{ width = $width; height = $height; file = "shots/$name" }
        }

        # The object model corroborates the bitmap for the two properties it can
        # actually report: whether PowerPoint agrees a shape is picture-filled,
        # and what it thinks the texture offsets are.
        $shapes = @()
        foreach ($shape in $slide.Shapes) {
            $entry = [ordered]@{ name = $shape.Name; fillType = $null; textureName = $null }
            try { $entry.fillType = [int]$shape.Fill.Type } catch {}
            try { $entry.textureName = $shape.Fill.TextureName } catch {}
            try { $entry.left = [double]$shape.Left; $entry.top = [double]$shape.Top } catch {}
            try { $entry.width = [double]$shape.Width; $entry.height = [double]$shape.Height } catch {}
            $shapes += $entry
        }

        $slides += [ordered]@{ slide = $index; exports = $exports; shapes = $shapes }
    }

    $records += [ordered]@{
        deck   = $deck.id
        opened = $true
        error  = $null
        slideWidth = [double]$pres.PageSetup.SlideWidth
        slideHeight = [double]$pres.PageSetup.SlideHeight
        slides = $slides
    }
    $pres.Close()
    Write-Host ("{0}: {1} slide(s)" -f $deck.id, $slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{
    powerPoint = @{ version = $version; build = $build }
    decks      = $records
} | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $root 'blip-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "done"
