# Experiment C3, step 2 - ask PowerPoint what it painted.
#
#   powershell -File tools/ground-truth/paint/fills/read.ps1 -Dir <work-dir>
#
# Reads `fill-inputs.json`, opens each probe deck, exports every slide as a BMP
# and records what the object model says about every shape's fill.
#
# ## The bitmap is the oracle here, not the object model
#
# C2 could use `Fill.ForeColor.RGB` because a solid fill is one colour. A
# gradient is a field of colours and the object model has one slot for it, so for
# C3 the pixels are the measurement and COM is corroboration. Two COM properties
# are worth having anyway:
#
#   - `Fill.GradientStops` reports position and colour per stop. If PowerPoint
#     hands them back in a different order from the one in the file, it sorted
#     them, and that is the answer to a question the bitmap can only imply.
#   - `Fill.Pattern` reports which `MsoPatternType` it thinks a pattern is, which
#     round-trips the enumeration the tiles are keyed by.
#
# ## Three export resolutions for the pattern decks
#
# A pattern tile is either a fixed number of device pixels or a fixed size in
# shape units, and those two behave identically until you change the resolution.
# So the pattern decks are exported at 960, 1920 and 3840 wide: if the tile
# period in pixels stays 8 across all three it is device-fixed, and if it doubles
# with the resolution it is in shape units.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason in C2: `Open`
# silently repairs, and a silently repaired deck reports success on exactly the
# files we are asking about. Fifteen of these decks exist to find out whether
# PowerPoint refuses them, so a refusal is recorded, not skipped.
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
$inputsPath = Join-Path $root 'fill-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no fill-inputs.json in $root - run tools/ground-truth/paint/fills/build-deck.ts first"
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

# The slide is 960 x 540 points. 1920 x 1080 is two pixels per point everywhere.
#
# The pattern decks also get 1280 x 720, which is the *only* resolution at which
# a tile can be read exactly: 960 pt is 13.333 in, so 1280 px is 96 DPI on the
# nose, and a pattern tile turns out to be a fixed physical size of one twelfth
# of an inch - eight pixels at 96 DPI. At 1920 the same tile is resampled to
# twelve pixels and every diagonal picks up an antialiased edge. 640 and 2560 are
# exact halves and doubles of 96 DPI, and are there to confirm the tile is
# physical rather than either device-fixed or shape-relative.
#
# A size is an object rather than a pair. A one-element array of arrays unrolls
# to its two scalars on the way out of the `if` below, and the loop then exports
# every deck a second time at width 1080.
function New-Size([int]$w, [int]$h) { [pscustomobject]@{ w = $w; h = $h } }

$defaultSizes = @((New-Size 1920 1080))
$patternSizes = @((New-Size 640 360), (New-Size 1280 720), (New-Size 1920 1080), (New-Size 2560 1440))
# The corner-softening deck: if the rounding is a fixed count of device pixels
# it is the same width at both, and if it is a share of the ramp it doubles.
$softenSizes = @((New-Size 1920 1080), (New-Size 3840 2160))

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file
    $sizes = if ($deck.deck -eq 'pattern' -or $deck.deck -eq 'patsize') { $patternSizes }
    elseif ($deck.deck -eq 'soften') { $softenSizes }
    else { $defaultSizes }

    $record = [ordered]@{
        deck     = $deck.deck
        file     = $deck.file
        opened   = $false
        repaired = $null
        error    = $null
        slides   = 0
        shapes   = @()
        bitmaps  = @()
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
        # rewrites this" from "PowerPoint refuses this", and for an enumeration
        # value the two mean different things.
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
                    $info = [ordered]@{
                        id           = [string]$shape.Name
                        slide        = $i
                        left         = [double]$shape.Left
                        top          = [double]$shape.Top
                        width        = [double]$shape.Width
                        height       = [double]$shape.Height
                        fillType     = $null
                        foreRgb      = $null
                        backRgb      = $null
                        transparency = $null
                        pattern      = $null
                        gradStyle    = $null
                        gradVariant  = $null
                        gradAngle    = $null
                        gradDegree   = $null
                        stops        = @()
                        readError    = $null
                    }
                    try {
                        $fill = $shape.Fill
                        try { $info.fillType = [int]$fill.Type } catch {}
                        try { $info.foreRgb = [int]$fill.ForeColor.RGB } catch {}
                        try { $info.backRgb = [int]$fill.BackColor.RGB } catch {}
                        try { $info.transparency = [double]$fill.Transparency } catch {}
                        try { $info.pattern = [int]$fill.Pattern } catch {}
                        try { $info.gradStyle = [int]$fill.GradientStyle } catch {}
                        try { $info.gradVariant = [int]$fill.GradientVariant } catch {}
                        try { $info.gradAngle = [double]$fill.GradientAngle } catch {}
                        try { $info.gradDegree = [double]$fill.GradientDegree } catch {}
                        try {
                            $gs = $fill.GradientStops
                            $list = @()
                            for ($k = 1; $k -le $gs.Count; $k++) {
                                $stop = $gs.Item($k)
                                $list += [ordered]@{
                                    index        = $k
                                    position     = [double]$stop.Position
                                    rgb          = [int]$stop.Color.RGB
                                    transparency = [double]$stop.Transparency
                                }
                            }
                            $info.stops = $list
                        }
                        catch {}
                    }
                    catch {
                        $info.readError = $_.Exception.Message
                    }
                    $shapes += $info
                }

                foreach ($size in $sizes) {
                    $w = $size.w
                    $h = $size.h
                    $name = "{0}-slide{1}-{2}.bmp" -f $deck.deck, $i, $w
                    $slide.Export((Join-Path $root $name), 'BMP', $w, $h)
                    $record.bitmaps += [ordered]@{ file = $name; slide = $i; width = $w; height = $h }
                }
            }
            $record.shapes = $shapes
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if (-not $record.opened) { 'REFUSED' } elseif ($record.repaired) { 'REPAIRED' } else { 'ok' }
    Write-Host ("{0,-16} {1,-9} {2} shape(s), {3} bitmap(s)" -f $deck.deck, $state, @($record.shapes).Count, @($record.bitmaps).Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$out = [ordered]@{
    slideWidth  = 960
    slideHeight = 540
    decks       = $decks
}
$json = $out | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $root 'com-readback-fills.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("wrote {0}" -f (Join-Path $root 'com-readback-fills.json'))
