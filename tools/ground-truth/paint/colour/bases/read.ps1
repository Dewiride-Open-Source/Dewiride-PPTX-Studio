# Experiment C2, step 2 - ask PowerPoint what it painted.
#
#   powershell -File tools/ground-truth/paint/colour/bases/read.ps1 -Dir <work-dir>
#
# Reads `swatch2-inputs.json` from the work directory, opens each probe deck,
# and writes `com-readback2.json` plus one BMP per slide.
#
# Two readbacks, deliberately:
#
#   - `Shape.Fill.ForeColor.RGB` and `Shape.Fill.Transparency` from the object
#     model. Transparency is the only way to see an `a:alpha` at all - a
#     translucent swatch's pixel is a composite with whatever is behind it, and
#     what is behind it is another swatch.
#   - the exported bitmap, which is what a user actually sees. Where the two
#     disagree the pixels win; 0.7-C found they agree on all 214, which is what
#     makes the object model usable as an oracle here.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason spelled out at length
# in packages/cli/scripts/powerpoint-oracle.ps1: `Open` silently repairs, and a
# silently repaired deck reports success on exactly the files we are asking
# about. A deck that fails to open is recorded as a refusal rather than skipped,
# because three of these decks exist to find out whether PowerPoint refuses them.
#
# Read-only. Nothing is written back to any deck. Attaches to a running
# PowerPoint if there is one and never quits it.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'swatch2-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no swatch2-inputs.json in $root - run tools/ground-truth/paint/colour/bases/build-deck.ts first"
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

# The slide is 960 x 540 points. Exporting at 1920 x 1080 gives two pixels per
# point, which is plenty to hit the centre of a 96 x 67 point cell.
$exportW = 1920
$exportH = 1080

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file
    $record = [ordered]@{
        deck      = $deck.deck
        file      = $deck.file
        opened    = $false
        repaired  = $null
        error     = $null
        slides    = 0
        shapes    = @()
        bitmaps   = @()
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
        # Would it open if repair were allowed? That separates "PowerPoint
        # rewrites this" from "PowerPoint refuses this", and the two mean
        # different things about an enumeration value.
        try {
            $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoTrue)
            $record.opened = $true
            $record.repaired = $true
        }
        catch {
            $record.repaired = $null
            $pres = $null
        }
    }

    if ($null -ne $pres) {
        try {
            $record.slides = $pres.Slides.Count
            $shapes = @()
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                foreach ($shape in $slide.Shapes) {
                    $rgb = $null
                    $transparency = $null
                    try { $rgb = [int]$shape.Fill.ForeColor.RGB } catch { $rgb = $null }
                    try { $transparency = [double]$shape.Fill.Transparency } catch { $transparency = $null }
                    $shapes += [ordered]@{
                        id           = [string]$shape.Name
                        slide        = $i
                        comRgb       = $rgb
                        transparency = $transparency
                        left         = [double]$shape.Left
                        top          = [double]$shape.Top
                        width        = [double]$shape.Width
                        height       = [double]$shape.Height
                    }
                }
                $bmp = Join-Path $root ("{0}-slide{1}.bmp" -f $deck.deck, $i)
                $slide.Export($bmp, 'BMP', $exportW, $exportH)
                $record.bitmaps += ("{0}-slide{1}.bmp" -f $deck.deck, $i)
            }
            $record.shapes = $shapes
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if (-not $record.opened) { 'REFUSED' } elseif ($record.repaired) { 'REPAIRED' } else { 'ok' }
    Write-Host ("{0,-10} {1,-9} {2} shape(s)" -f $deck.deck, $state, @($record.shapes).Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$out = [ordered]@{
    slideWidth  = 960
    slideHeight = 540
    exportWidth = $exportW
    exportHeight = $exportH
    decks       = $decks
}
$json = $out | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $root 'com-readback2.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("wrote {0}" -f (Join-Path $root 'com-readback2.json'))
