# Experiment C6, step 2 - ask PowerPoint where the shapes ended up.
#
#   powershell -File tools/ground-truth/render/read.ps1 -Dir <work-dir>
#
# Reads `transform-inputs.json`, opens each probe deck, and records every shape
# twice.
#
# ## Two passes, because the object model is inconsistent about one field
#
# Measured on PowerPoint's own output in `tools/ground-truth/render/author.ps1`: for a shape
# inside a rotated, mirrored group, `GroupItems(i).Rotation` reports the
# **composed** rotation while `GroupItems(i).HorizontalFlip` reports the child's
# **own** attribute. Left/Top/Width/Height are composed - a child of a group
# turned 30 degrees and mirrored reported `L=321.2435302734375`, and the same
# number appears in the file as `x="4079793"` once that group is ungrouped.
#
# So the first pass reads the shapes in place, which gives the composed frame,
# and the second ungroups everything and reads again, which gives the composed
# orientation. `Ungroup` is one level deep - measured - so it runs until no
# group is left.
#
# **Nothing is saved.** The presentation is opened read/write because `Ungroup`
# needs it to be, the ungrouping happens in memory, and `Saved` is forced true
# before the close so PowerPoint neither writes nor asks. The bitmap is exported
# before any of that, from the file as built.
#
# Attaches to a running PowerPoint if there is one and never quits one it did
# not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3
$msoGroup = 6

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'transform-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no transform-inputs.json in $root - run tools/ground-truth/render/build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json
$exportW = [int]$inputs.exportPixels.w
$exportH = [int]$inputs.exportPixels.h

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

function Read-Shape($shape) {
    $out = [ordered]@{
        name       = $null
        left       = $null
        top        = $null
        width      = $null
        height     = $null
        rotation   = $null
        flipH      = $null
        flipV      = $null
        type       = $null
        autoShape  = $null
        fillType   = $null
        fillRgb    = $null
        lineVisible = $null
        lineWeight = $null
    }
    try { $out.name = [string]$shape.Name } catch {}
    try { $out.left = [double]$shape.Left } catch {}
    try { $out.top = [double]$shape.Top } catch {}
    try { $out.width = [double]$shape.Width } catch {}
    try { $out.height = [double]$shape.Height } catch {}
    try { $out.rotation = [double]$shape.Rotation } catch {}
    try { $out.flipH = [int]$shape.HorizontalFlip } catch {}
    try { $out.flipV = [int]$shape.VerticalFlip } catch {}
    try { $out.type = [int]$shape.Type } catch {}
    try { $out.autoShape = [int]$shape.AutoShapeType } catch {}
    try { $out.fillType = [int]$shape.Fill.Type } catch {}
    try { $out.fillRgb = [int]$shape.Fill.ForeColor.RGB } catch {}
    try { $out.lineVisible = [int]$shape.Line.Visible } catch {}
    try { $out.lineWeight = [double]$shape.Line.Weight } catch {}
    return $out
}

# Every shape on a slide, groups included, flattened by name. `GroupItems` was
# measured to flatten nested groups already, so this walks anyway and dedupes.
function Read-All($slide) {
    $rows = @()
    $seen = @{}
    $stack = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        [void]$stack.Add($slide.Shapes.Item($i))
    }
    while ($stack.Count -gt 0) {
        $shape = $stack[0]
        $stack.RemoveAt(0)
        $row = Read-Shape $shape
        if (-not $seen.ContainsKey($row.name)) {
            $seen[$row.name] = $true
            $rows += $row
        }
        try {
            if ([int]$shape.Type -eq $msoGroup) {
                for ($j = 1; $j -le $shape.GroupItems.Count; $j++) {
                    [void]$stack.Add($shape.GroupItems.Item($j))
                }
            }
        }
        catch {}
    }
    return $rows
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
    }

    $pres = $null
    try {
        # FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair
        $pres = $app.Presentations.Open2007($file, $msoFalse, $msoFalse, $msoFalse, $msoFalse)
        $record.opened = $true
        $record.repaired = $false
    }
    catch {
        $record.error = $_.Exception.Message
        try {
            $pres = $app.Presentations.Open2007($file, $msoFalse, $msoFalse, $msoFalse, $msoTrue)
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
                    index     = $i
                    inPlace   = (Read-All $slide)
                    ungrouped = @()
                    bitmap    = $null
                    error     = $null
                }

                $stem = "{0}-s{1:d2}" -f $deck.deck, $i
                $bmp = Join-Path $root ($stem + '.bmp')
                try {
                    $slide.Export($bmp, 'BMP', $exportW, $exportH)
                    $entry.bitmap = $stem + '.bmp'
                }
                catch {
                    $entry.error = $_.Exception.Message
                }

                # Ungroup until nothing is a group. One level per call, measured.
                for ($pass = 0; $pass -lt 8; $pass++) {
                    $names = @()
                    for ($k = 1; $k -le $slide.Shapes.Count; $k++) {
                        $sh = $slide.Shapes.Item($k)
                        if ([int]$sh.Type -eq $msoGroup) { $names += [string]$sh.Name }
                    }
                    if ($names.Count -eq 0) { break }
                    foreach ($name in $names) {
                        try { $slide.Shapes.Item($name).Ungroup() | Out-Null } catch {}
                    }
                }
                $entry.ungrouped = (Read-All $slide)

                $record.slides += $entry
            }
        }
        finally {
            # Never write. The ungrouping above is in memory only.
            try { $pres.Saved = $msoTrue } catch {}
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-20} {1,-9} {2} slide(s)" -f $deck.deck, $state, $record.slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText((Join-Path $root 'transform-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
