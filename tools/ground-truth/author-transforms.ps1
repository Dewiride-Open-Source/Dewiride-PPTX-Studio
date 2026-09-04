# Experiment C6, step 0 - ask PowerPoint to author the transforms.
#
#   powershell -File tools/ground-truth/author-transforms.ps1 -Dir <work-dir>
#
# The cheapest half of every ground-truth sub-phase so far, and the highest
# yield per line: before building a single probe, get the format's own author to
# write down the answer.
#
# ## Why a flip can be asked in writing at all
#
# `a:xfrm` carries `@rot`, `@flipH` and `@flipV` and says nothing about the
# order they compose in. The two candidate readings differ - `R(t)F` against
# `F R(t)` - and they are not equivalent, because for any reflection
# `F R(t) F = R(-t)`, so the two orders differ by the sign of the rotation.
#
# That is what makes this measurable without a pixel. Take a shape rotated 30
# degrees and ask PowerPoint to mirror it. The mirrored figure is fixed - it is
# whatever the user sees - and PowerPoint has to write an `a:xfrm` that produces
# it. Under flip-then-rotate it must write `rot=-30 flipH=1`; under
# rotate-then-flip it must write `rot=+30 flipH=1`. Its writer and its renderer
# agree with each other by construction, so what it writes *is* the order it
# renders in.
#
# ## And a group can be asked the same way
#
# `Ungroup` is PowerPoint composing a group's transform into each child and
# writing the result down. Grouping, rotating, flipping and then ungrouping
# gives its own arithmetic for the composition, in numbers, in a file.
#
# Two decks come out of this: `pp-groups.pptx` with the groups intact and
# `pp-ungrouped.pptx` after ungrouping the same shapes. `analyse-transforms.ts`
# reads both.
#
# Writes only into -Dir. Attaches to a running PowerPoint if there is one and
# never quits one it did not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoFalse = 0
$ppAlertsNone = 1
$ppSaveAsOpenXMLPresentation = 24
$ppLayoutBlank = 12
$msoShapeRightTriangle = 8
$msoShapeRectangle = 1
$msoFlipHorizontal = 0
$msoFlipVertical = 1

New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$root = (Resolve-Path -LiteralPath $Dir).Path

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

$log = @()

function Read-Shape($shape, $label) {
    $out = [ordered]@{
        label    = $label
        name     = $null
        left     = $null
        top      = $null
        width    = $null
        height   = $null
        rotation = $null
        flipH    = $null
        flipV    = $null
        type     = $null
        children = @()
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
    return $out
}

$pres = $app.Presentations.Add($msoFalse)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

# ---------------------------------------------------------------------------
# Slide 1 - the flip composition question.
#
# A right triangle, because it is asymmetric on both axes: a rectangle cannot
# tell a mirror from an identity and would report a false agreement.
# ---------------------------------------------------------------------------
$slide1 = $pres.Slides.Add(1, $ppLayoutBlank)
$angles = @(0, 30, 45, 120, 200)
$ops = @('none', 'H', 'V', 'HV')

for ($r = 0; $r -lt $ops.Count; $r++) {
    for ($c = 0; $c -lt $angles.Count; $c++) {
        $x = 30 + $c * 180
        $y = 20 + $r * 130
        $sh = $slide1.Shapes.AddShape($msoShapeRightTriangle, $x, $y, 120, 80)
        $sh.Name = ("flip-{0}-{1}" -f $angles[$c], $ops[$r])
        $sh.Rotation = $angles[$c]

        # Read what a rotation alone did to Width/Height, before any flip. If
        # COM reported a rotated bounding box these would change, and every
        # position in every other C6 reading would need reinterpreting.
        $before = Read-Shape $sh 'before'

        switch ($ops[$r]) {
            'H' { $sh.Flip($msoFlipHorizontal) }
            'V' { $sh.Flip($msoFlipVertical) }
            'HV' { $sh.Flip($msoFlipHorizontal); $sh.Flip($msoFlipVertical) }
        }

        $log += [ordered]@{
            kind     = 'flip'
            angle    = $angles[$c]
            op       = $ops[$r]
            shape    = $sh.Name
            before   = $before
            after    = (Read-Shape $sh 'after')
        }
    }
}

# ---------------------------------------------------------------------------
# Slide 2 - a group, untouched. What does PowerPoint write for chOff/chExt on
# a group it has just made, and what does GroupItems report for the children?
# ---------------------------------------------------------------------------
$slide2 = $pres.Slides.Add(2, $ppLayoutBlank)
$a = $slide2.Shapes.AddShape($msoShapeRectangle, 100, 100, 120, 60)
$a.Name = 'g1-child-a'
$b = $slide2.Shapes.AddShape($msoShapeRightTriangle, 300, 220, 200, 100)
$b.Name = 'g1-child-b'
$g1 = $slide2.Shapes.Range(@('g1-child-a', 'g1-child-b')).Group()
$g1.Name = 'g1'
$entry = [ordered]@{ kind = 'group-fresh'; slide = 2; group = (Read-Shape $g1 'group'); children = @() }
for ($i = 1; $i -le $g1.GroupItems.Count; $i++) {
    $entry.children += (Read-Shape $g1.GroupItems.Item($i) ("child" + $i))
}
$log += $entry

# ---------------------------------------------------------------------------
# Slide 3 - the same group, resized. Does the scale land in the children's own
# `a:off`/`a:ext`, or in the group's `a:chExt`? Phase 5.3 turns on the answer.
# ---------------------------------------------------------------------------
$slide3 = $pres.Slides.Add(3, $ppLayoutBlank)
$a = $slide3.Shapes.AddShape($msoShapeRectangle, 100, 100, 120, 60)
$a.Name = 'g2-child-a'
$b = $slide3.Shapes.AddShape($msoShapeRightTriangle, 300, 220, 200, 100)
$b.Name = 'g2-child-b'
$g2 = $slide3.Shapes.Range(@('g2-child-a', 'g2-child-b')).Group()
$g2.Name = 'g2'
$g2.LockAspectRatio = $msoFalse
$g2.Width = [double]$g2.Width * 2.0
$g2.Height = [double]$g2.Height * 0.5
$entry = [ordered]@{ kind = 'group-resized'; slide = 3; group = (Read-Shape $g2 'group'); children = @() }
for ($i = 1; $i -le $g2.GroupItems.Count; $i++) {
    $entry.children += (Read-Shape $g2.GroupItems.Item($i) ("child" + $i))
}
$log += $entry

# ---------------------------------------------------------------------------
# Slide 4 - a rotated, flipped group. What does GroupItems report for a child
# once the group is turned? The answer decides whether `Shape.Left` inside a
# group is slide space or group space, which is the whole method for C6's
# synthetic probes.
# ---------------------------------------------------------------------------
$slide4 = $pres.Slides.Add(4, $ppLayoutBlank)
$a = $slide4.Shapes.AddShape($msoShapeRectangle, 100, 100, 120, 60)
$a.Name = 'g3-child-a'
$b = $slide4.Shapes.AddShape($msoShapeRightTriangle, 300, 220, 200, 100)
$b.Name = 'g3-child-b'
$g3 = $slide4.Shapes.Range(@('g3-child-a', 'g3-child-b')).Group()
$g3.Name = 'g3'
$g3.Rotation = 30
$g3.Flip($msoFlipHorizontal)
$entry = [ordered]@{ kind = 'group-turned'; slide = 4; group = (Read-Shape $g3 'group'); children = @() }
for ($i = 1; $i -le $g3.GroupItems.Count; $i++) {
    $entry.children += (Read-Shape $g3.GroupItems.Item($i) ("child" + $i))
}
$log += $entry

# ---------------------------------------------------------------------------
# Slide 5 - a rotated group holding a rotated, flipped child. Ungrouping this
# is PowerPoint composing two transforms and writing the result down.
# ---------------------------------------------------------------------------
$slide5 = $pres.Slides.Add(5, $ppLayoutBlank)
$a = $slide5.Shapes.AddShape($msoShapeRightTriangle, 120, 120, 160, 80)
$a.Name = 'g4-child-a'
$a.Rotation = 40
$a.Flip($msoFlipHorizontal)
$b = $slide5.Shapes.AddShape($msoShapeRectangle, 400, 260, 140, 90)
$b.Name = 'g4-child-b'
$b.Rotation = 15
$g4 = $slide5.Shapes.Range(@('g4-child-a', 'g4-child-b')).Group()
$g4.Name = 'g4'
$g4.Rotation = 25
$entry = [ordered]@{ kind = 'group-nested-rot'; slide = 5; group = (Read-Shape $g4 'group'); children = @() }
for ($i = 1; $i -le $g4.GroupItems.Count; $i++) {
    $entry.children += (Read-Shape $g4.GroupItems.Item($i) ("child" + $i))
}
$log += $entry

# ---------------------------------------------------------------------------
# Slide 6 - a group inside a group, so the composition can be asked twice.
# ---------------------------------------------------------------------------
$slide6 = $pres.Slides.Add(6, $ppLayoutBlank)
$a = $slide6.Shapes.AddShape($msoShapeRectangle, 100, 100, 100, 60)
$a.Name = 'g5-a'
$b = $slide6.Shapes.AddShape($msoShapeRightTriangle, 240, 100, 100, 60)
$b.Name = 'g5-b'
$inner = $slide6.Shapes.Range(@('g5-a', 'g5-b')).Group()
$inner.Name = 'g5-inner'
$c = $slide6.Shapes.AddShape($msoShapeRectangle, 100, 300, 340, 80)
$c.Name = 'g5-c'
$outer = $slide6.Shapes.Range(@('g5-inner', 'g5-c')).Group()
$outer.Name = 'g5-outer'
$outer.LockAspectRatio = $msoFalse
$outer.Width = [double]$outer.Width * 1.5
$entry = [ordered]@{ kind = 'group-nested'; slide = 6; group = (Read-Shape $outer 'group'); children = @() }
for ($i = 1; $i -le $outer.GroupItems.Count; $i++) {
    $child = $outer.GroupItems.Item($i)
    $rec = Read-Shape $child ("child" + $i)
    try {
        if ([int]$child.Type -eq 6) {
            for ($j = 1; $j -le $child.GroupItems.Count; $j++) {
                $rec.children += (Read-Shape $child.GroupItems.Item($j) ("grandchild" + $j))
            }
        }
    }
    catch {}
    $entry.children += $rec
}
$log += $entry

# ---------------------------------------------------------------------------
# Save with the groups intact, then ungroup every one of them and save again.
# ---------------------------------------------------------------------------
$groupsPath = Join-Path $root 'pp-groups.pptx'
if (Test-Path -LiteralPath $groupsPath) { Remove-Item -LiteralPath $groupsPath -Force }
$pres.SaveAs($groupsPath, $ppSaveAsOpenXMLPresentation, $msoFalse)
Write-Host ("wrote {0}" -f $groupsPath)

$ungrouped = @()
foreach ($idx in 2, 3, 4, 5, 6) {
    $slide = $pres.Slides.Item($idx)
    # Ungroup returns the freed shapes. On the nested slide the inner group
    # survives one call and has to be asked again, which is itself a reading:
    # Ungroup is one level, not a flatten.
    $names = @()
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $sh = $slide.Shapes.Item($i)
        if ([int]$sh.Type -eq 6) { $names += [string]$sh.Name }
    }
    foreach ($name in $names) {
        try { $slide.Shapes.Item($name).Ungroup() | Out-Null } catch {}
    }
    $entry = [ordered]@{ kind = 'ungrouped'; slide = $idx; shapes = @() }
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $entry.shapes += (Read-Shape $slide.Shapes.Item($i) ("shape" + $i))
    }
    $ungrouped += $entry
}
$log += $ungrouped

$ungroupedPath = Join-Path $root 'pp-ungrouped.pptx'
if (Test-Path -LiteralPath $ungroupedPath) { Remove-Item -LiteralPath $ungroupedPath -Force }
$pres.SaveAs($ungroupedPath, $ppSaveAsOpenXMLPresentation, $msoFalse)
Write-Host ("wrote {0}" -f $ungroupedPath)

$pres.Close()
if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ entries = $log } | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'author-transforms-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("{0} entries" -f $log.Count)
Write-Host 'done'
