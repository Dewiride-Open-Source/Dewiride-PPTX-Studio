# Experiment F1, step 1 - ask PowerPoint to draw the corpus.
#
#   powershell -File tools/ground-truth/render/fidelity/read.ps1 -Dir <work-dir>
#
# Reads `oracle-inputs.json`, opens each deck and exports every slide as a PNG
# at the width the harness rasterises at.
#
# ## Every slide is exported twice
#
# The whole design rests on the oracle being a fixed reference, and a reference
# that is not reproducible is not one. So each slide is exported to two paths
# and the bytes are compared before either is kept. Finding out that
# `Slide.Export` is nondeterministic on day one costs one run; finding out after
# 149 grids are committed and trusted costs the sub-phase.
#
# ## Why not a fixed pixel height
#
# The corpus is not all 16:9 - `a41-a4` and `a42-custom-size` exist precisely so
# that something asks this question. The width is fixed and the height follows
# the deck's own aspect, which is the same rule `geometryOf` applies on our side.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason every other read.ps1
# gives: `Open` silently repairs, and a silently repaired deck is not the deck
# whose bytes we committed.
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
$inputsPath = Join-Path $root 'oracle-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no oracle-inputs.json in $root - run tools/ground-truth/render/fidelity/analyse.ts --capture first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$repo = $inputs.repo
$rasterWidth = [int]$inputs.rasterWidth

$started = $false
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
}
catch {
    $app = New-Object -ComObject PowerPoint.Application
    $started = $true
}
$app.DisplayAlerts = $ppAlertsNone
$app.AutomationSecurity = $msoAutomationSecurityForceDisable

# Read before the loop: the `finally` quits an application this script started,
# and every property of a quit application is null.
$version = $app.Version
$build = $app.Build

$records = @()

try {
    foreach ($deck in $inputs.decks) {
        $full = Join-Path $repo $deck.path
        $record = [ordered]@{ id = $deck.id; path = $deck.path; slides = @(); error = $null }
        $pres = $null
        try {
            $pres = $app.Presentations.Open2007($full, $msoTrue, $msoFalse, $msoFalse, $msoFalse)
            $pageWidth = [double]$pres.PageSetup.SlideWidth
            $pageHeight = [double]$pres.PageSetup.SlideHeight
            $height = [int][Math]::Round($rasterWidth * $pageHeight / $pageWidth)
            $record.pointWidth = $pageWidth
            $record.pointHeight = $pageHeight
            $record.pixelWidth = $rasterWidth
            $record.pixelHeight = $height

            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $name = ('{0}-{1:d2}.png' -f $deck.id, $i)
                $keep = Join-Path $root $name
                $again = Join-Path $root ('{0}-{1:d2}.again.png' -f $deck.id, $i)

                $slide.Export($keep, 'PNG', $rasterWidth, $height)
                $slide.Export($again, 'PNG', $rasterWidth, $height)

                # Both are kept. Whether the two agree is decided on their
                # decoded pixels in `analyse.ts`, not here on their bytes:
                # PowerPoint writes a palettised PNG and the palette order is not
                # stable between exports, so byte equality answers a question
                # about the encoder rather than about the picture.
                $a = [IO.File]::ReadAllBytes($keep)
                $b = [IO.File]::ReadAllBytes($again)
                $sameBytes = $a.Length -eq $b.Length
                if ($sameBytes) {
                    for ($k = 0; $k -lt $a.Length; $k++) {
                        if ($a[$k] -ne $b[$k]) { $sameBytes = $false; break }
                    }
                }

                $record.slides += [ordered]@{
                    slide     = $i
                    file      = $name
                    again     = ('{0}-{1:d2}.again.png' -f $deck.id, $i)
                    bytes     = $a.Length
                    sameBytes = $sameBytes
                }
            }
        }
        catch {
            $record.error = $_.Exception.Message
        }
        finally {
            if ($null -ne $pres) { try { $pres.Close() } catch {} }
        }
        $records += $record
    }
}
finally {
    if ($started) { try { $app.Quit() } catch {} }
}

$out = [ordered]@{
    rasterWidth = $rasterWidth
    powerpoint  = ('{0} build {1}' -f $version, $build)
    decks       = $records
}
$json = $out | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText((Join-Path $root 'oracle-export.json'), $json, [Text.UTF8Encoding]::new($false))
Write-Output ('exported {0} deck(s) to {1}' -f $records.Count, $root)
