# Experiment C5, step 2 - ask PowerPoint what it resolved.
#
#   powershell -File tools/ground-truth/model/read.ps1 -Dir <work-dir>
#
# Reads `sheet-inputs.json`, opens each probe deck, and for every shape on every
# slide records the position, size and resolved colours the object model reports.
# A BMP of each slide is exported alongside.
#
# ## Why the object model is the measurement here, and the bitmap the check
#
# C3 and C4 sampled bitmaps because a fill and a stroke are pictures: no COM
# property says what colour is at a point. An inheritance is not a picture. A
# placeholder with no `a:xfrm` of its own has a position anyway, PowerPoint
# knows it, and `Shape.Left` reports it in points with no antialiasing in the
# way. That is the resolver's own answer, read out of the resolver.
#
# So this script reads:
#
#   - Left/Top/Width/Height, which names the sheet a placeholder matched, because
#     every candidate in the experiment sits in a box no other candidate shares.
#   - Fill.Type and Fill.ForeColor.RGB, which say which entry of which style list
#     a `fillRef` reached and what colour `phClr` was.
#   - Line.Weight, which does the same for `lnRef` against a theme whose three
#     line styles are half a point, two points and four and a half.
#   - CustomLayout.Name and Master.Name, which say which sheets PowerPoint bound
#     the slide to - the only direct evidence for what a broken relationship does.
#   - Background.Fill, for the `p:bgRef` half.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason C2 found: `Open`
# repairs silently, and a silently repaired deck reports success on exactly the
# files being asked about. A second attempt with repair allowed distinguishes
# REFUSED from REPAIRED.
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
$ppShapeFormatBMP = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'sheet-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no sheet-inputs.json in $root - run tools/ground-truth/model/build-deck.ts first"
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

# A COM colour is a BGR long. The conversion is left to the analysis so that
# what is committed is what PowerPoint returned, not an interpretation of it.
function Read-Fill($fill) {
    $out = [ordered]@{ type = $null; visible = $null; rgb = $null; transparency = $null }
    try { $out.type = [int]$fill.Type } catch {}
    try { $out.visible = [int]$fill.Visible } catch {}
    try { $out.rgb = [int]$fill.ForeColor.RGB } catch {}
    try { $out.transparency = [double]$fill.Transparency } catch {}
    return $out
}

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file

    $record = [ordered]@{
        deck     = $deck.deck
        file     = $deck.file
        hostile  = [bool]$deck.hostile
        opened   = $false
        repaired = $null
        error    = $null
        slides   = @()
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
        # "PowerPoint rewrites this" and "PowerPoint refuses this" are different
        # answers, and only the second attempt tells them apart.
        try {
            $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoTrue)
            $record.opened = $true
            $record.repaired = $true
        }
        catch {
            $record.opened = $false
            $record.repaired = $null
        }
    }

    if ($record.opened) {
        try {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $entry = [ordered]@{
                    index               = $i
                    name                = [string]$slide.Name
                    layoutName          = $null
                    masterName          = $null
                    follow              = $null
                    displayMasterShapes = $null
                    background          = $null
                    shapes              = @()
                }
                try { $entry.layoutName = [string]$slide.CustomLayout.Name } catch {}
                try { $entry.masterName = [string]$slide.Master.Name } catch {}
                try { $entry.follow = [int]$slide.FollowMasterBackground } catch {}
                try { $entry.displayMasterShapes = [int]$slide.DisplayMasterShapes } catch {}
                try { $entry.background = Read-Fill $slide.Background.Fill } catch {}

                foreach ($shape in $slide.Shapes) {
                    $s = [ordered]@{
                        name        = [string]$shape.Name
                        left        = $null
                        top         = $null
                        width       = $null
                        height      = $null
                        placeholder = $null
                        contained   = $null
                        fill        = $null
                        lineVisible = $null
                        lineWeight  = $null
                        lineRgb     = $null
                    }
                    try { $s.left = [double]$shape.Left } catch {}
                    try { $s.top = [double]$shape.Top } catch {}
                    try { $s.width = [double]$shape.Width } catch {}
                    try { $s.height = [double]$shape.Height } catch {}
                    try { $s.placeholder = [int]$shape.PlaceholderFormat.Type } catch {}
                    try { $s.contained = [int]$shape.PlaceholderFormat.ContainedType } catch {}
                    try { $s.fill = Read-Fill $shape.Fill } catch {}
                    try { $s.lineVisible = [int]$shape.Line.Visible } catch {}
                    try { $s.lineWeight = [double]$shape.Line.Weight } catch {}
                    try { $s.lineRgb = [int]$shape.Line.ForeColor.RGB } catch {}
                    $entry.shapes += $s
                }

                $record.slides += $entry

                $stem = "{0}-s{1:d2}" -f $deck.deck, $i
                $bmp = Join-Path $root ($stem + '.bmp')
                try {
                    $slide.Export($bmp, 'BMP', 1920, 1080)
                    $record.bitmaps += [ordered]@{ slide = $i; file = ($stem + '.bmp'); width = 1920; height = 1080 }
                }
                catch {
                    $record.bitmaps += [ordered]@{ slide = $i; file = $null; error = $_.Exception.Message }
                }
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-22} {1,-9} {2} slide(s)" -f $deck.deck, $state, $record.slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $root 'sheet-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
