# Experiment C4, step 2 - ask PowerPoint what it painted.
#
#   powershell -File tools/ground-truth/read-lines.ps1 -Dir <work-dir>
#
# Reads `line-inputs.json`, opens each probe deck, exports its slide as a BMP and
# records what the object model says about every shape's line and effects.
#
# ## Two resolutions, and why
#
# Everything is exported at 1920 x 1080, which over a 960 x 540 point slide is
# exactly two pixels per point. The dash and arrowhead decks are also exported at
# 3840 x 2160, four pixels per point: a dash run is measured by finding where
# coverage crosses one half, which is good to about a pixel either way, and this
# sub-phase's central number is that run length in multiples of the stroke width.
# Halving the error costs one more export.
#
# ## The object model is corroboration, and for the hostile decks it is evidence
#
# The bitmap is the measurement, as in C3. But COM answers one thing pixels
# cannot: whether PowerPoint *clamped* a value it did not like rather than
# refusing the file. A negative `@w` that opens cleanly and reads back as
# `Weight = 0.75` has been silently corrected, and that is a different fact from
# either "accepted" or "refused" - so Line.Weight, DashStyle, Style, InsetPen and
# the arrowhead triples are all read back for every shape.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason C2 found: `Open`
# repairs silently, and a silently repaired deck reports success on exactly the
# files being asked about.
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
$inputsPath = Join-Path $root 'line-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no line-inputs.json in $root - run build-line-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json
$fine = @($inputs.fineDecks)

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

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file
    # The leading comma is load-bearing. `@(@(1920, 1080))` is a one-element
    # array of arrays, which PowerShell unrolls back to the two-element array
    # `@(1920, 1080)` - and then the loop below iterates over 1920 and 1080 as
    # scalars and exports two wrong bitmaps. `,@(...)` is the only spelling that
    # survives.
    $sizes = , @(1920, 1080)
    if ($fine -contains $deck.deck) { $sizes = @(@(1920, 1080), @(3840, 2160)) }

    $record = [ordered]@{
        deck     = $deck.deck
        file     = $deck.file
        opened   = $false
        repaired = $null
        error    = $null
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
        # Would it open if repair were allowed? "PowerPoint rewrites this" and
        # "PowerPoint refuses this" are different answers to a hostile probe.
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
            $shapes = @()
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                foreach ($shape in $slide.Shapes) {
                    $info = [ordered]@{
                        id            = [string]$shape.Name
                        slide         = $i
                        left          = [double]$shape.Left
                        top           = [double]$shape.Top
                        width         = [double]$shape.Width
                        height        = [double]$shape.Height
                        lineVisible   = $null
                        weight        = $null
                        dashStyle     = $null
                        lineStyle     = $null
                        insetPen      = $null
                        beginType     = $null
                        beginLen      = $null
                        beginWidth    = $null
                        endType       = $null
                        endLen        = $null
                        endWidth      = $null
                        shadowType    = $null
                        shadowVisible = $null
                        shadowBlur    = $null
                        shadowOffX    = $null
                        shadowOffY    = $null
                        glowRadius    = $null
                        softEdgeRad   = $null
                        reflType      = $null
                        readError     = $null
                    }
                    try {
                        $line = $shape.Line
                        try { $info.lineVisible = [int]$line.Visible } catch {}
                        try { $info.weight = [double]$line.Weight } catch {}
                        try { $info.dashStyle = [int]$line.DashStyle } catch {}
                        try { $info.lineStyle = [int]$line.Style } catch {}
                        try { $info.insetPen = [int]$line.InsetPen } catch {}
                        try { $info.beginType = [int]$line.BeginArrowheadStyle } catch {}
                        try { $info.beginLen = [int]$line.BeginArrowheadLength } catch {}
                        try { $info.beginWidth = [int]$line.BeginArrowheadWidth } catch {}
                        try { $info.endType = [int]$line.EndArrowheadStyle } catch {}
                        try { $info.endLen = [int]$line.EndArrowheadLength } catch {}
                        try { $info.endWidth = [int]$line.EndArrowheadWidth } catch {}
                        $shadow = $shape.Shadow
                        try { $info.shadowType = [int]$shadow.Type } catch {}
                        try { $info.shadowVisible = [int]$shadow.Visible } catch {}
                        try { $info.shadowBlur = [double]$shadow.Blur } catch {}
                        try { $info.shadowOffX = [double]$shadow.OffsetX } catch {}
                        try { $info.shadowOffY = [double]$shadow.OffsetY } catch {}
                        try { $info.glowRadius = [double]$shape.Glow.Radius } catch {}
                        try { $info.softEdgeRad = [double]$shape.SoftEdge.Radius } catch {}
                        try { $info.reflType = [int]$shape.Reflection.Type } catch {}
                    }
                    catch {
                        $info.readError = $_.Exception.Message
                    }
                    $shapes += $info
                }

                foreach ($size in $sizes) {
                    $w = $size[0]
                    $h = $size[1]
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
    Write-Host ("{0,-14} {1,-9} {2} shape(s), {3} bitmap(s)" -f $deck.deck, $state, @($record.shapes).Count, @($record.bitmaps).Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$out = [ordered]@{
    slideWidth  = 960
    slideHeight = 540
    decks       = $decks
}
$json = $out | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $root 'com-readback-lines.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("wrote {0}" -f (Join-Path $root 'com-readback-lines.json'))
