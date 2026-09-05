# Experiment T6, step 1 - make PowerPoint write the text frame down.
#
#   powershell -File tools/ground-truth/text/frames/author.ps1 -Dir <work-dir>
#
# Drives TextFrame2 through every orientation, anchor, margin, column count and
# autosize mode, then saves, so the a:bodyPr PowerPoint writes for each is in the
# file. Writes pp-frames.pptx and author-frames-log.json.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3
$ppLayoutBlank = 12
$ppLayoutText = 2
$ppSaveAsOpenXMLPresentation = 24
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1

$root = (Resolve-Path -LiteralPath $Dir).Path

# MsoTextOrientation. 7 is not defined by the enum; it is probed anyway, because
# ST_TextVerticalType has seven values and this enum has six usable ones, and
# finding out which one has no COM spelling is part of the question.
$orientations = @(
    @{ name = 'horizontal'; value = 1 },
    @{ name = 'upward'; value = 2 },
    @{ name = 'downward'; value = 3 },
    @{ name = 'verticalFarEast'; value = 4 },
    @{ name = 'vertical'; value = 5 },
    @{ name = 'horizontalRotatedFarEast'; value = 6 },
    @{ name = 'undocumented7'; value = 7 }
)

# MsoVerticalAnchor.
$vAnchors = @(
    @{ name = 'top'; value = 1 },
    @{ name = 'topBaseline'; value = 2 },
    @{ name = 'middle'; value = 3 },
    @{ name = 'bottom'; value = 4 },
    @{ name = 'bottomBaseline'; value = 5 }
)

# MsoHorizontalAnchor.
$hAnchors = @(
    @{ name = 'none'; value = 1 },
    @{ name = 'center'; value = 2 }
)

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

$log = [ordered]@{
    orientations = @()
    verticalAnchors = @()
    horizontalAnchors = @()
    defaults = @()
    columns = @()
    autoSize = @()
    bounds = @()
}

$pres = $app.Presentations.Add($msoFalse)
$pres.PageSetup.SlideSize = 3   # ppSlideSizeOnScreen: 10in x 7.5in, whatever; sizes are read back
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function New-Box($slide, $name, $x, $y, $w, $h, $text) {
    $sp = $slide.Shapes.AddShape($msoShapeRectangle, $x, $y, $w, $h)
    $sp.Name = $name
    $sp.TextFrame2.TextRange.Text = $text
    $sp.TextFrame2.TextRange.Font.Size = 18
    $sp.TextFrame2.TextRange.Font.Name = 'Arial'
    return $sp
}

function Read-Frame($sp, $kind = '') {
    $f = [ordered]@{
        kind = $kind
        name = ''
        orientation = $null
        verticalAnchor = $null
        horizontalAnchor = $null
        marginLeft = $null
        marginTop = $null
        marginRight = $null
        marginBottom = $null
        wordWrap = $null
        autoSize = $null
        columnNumber = $null
        columnSpacing = $null
        left = $null
        top = $null
        width = $null
        height = $null
        boundLeft = $null
        boundTop = $null
        boundWidth = $null
        boundHeight = $null
    }
    try { $f.name = [string]$sp.Name } catch {}
    try { $f.orientation = [int]$sp.TextFrame2.Orientation } catch {}
    try { $f.verticalAnchor = [int]$sp.TextFrame2.VerticalAnchor } catch {}
    try { $f.horizontalAnchor = [int]$sp.TextFrame2.HorizontalAnchor } catch {}
    try { $f.marginLeft = [double]$sp.TextFrame2.MarginLeft } catch {}
    try { $f.marginTop = [double]$sp.TextFrame2.MarginTop } catch {}
    try { $f.marginRight = [double]$sp.TextFrame2.MarginRight } catch {}
    try { $f.marginBottom = [double]$sp.TextFrame2.MarginBottom } catch {}
    try { $f.wordWrap = [int]$sp.TextFrame2.WordWrap } catch {}
    try { $f.autoSize = [int]$sp.TextFrame2.AutoSize } catch {}
    try { $f.columnNumber = [int]$sp.TextFrame2.Column.Number } catch {}
    try { $f.columnSpacing = [double]$sp.TextFrame2.Column.Spacing } catch {}
    try { $f.left = [double]$sp.Left } catch {}
    try { $f.top = [double]$sp.Top } catch {}
    try { $f.width = [double]$sp.Width } catch {}
    try { $f.height = [double]$sp.Height } catch {}
    try { $f.boundLeft = [double]$sp.TextFrame2.TextRange.BoundLeft } catch {}
    try { $f.boundTop = [double]$sp.TextFrame2.TextRange.BoundTop } catch {}
    try { $f.boundWidth = [double]$sp.TextFrame2.TextRange.BoundWidth } catch {}
    try { $f.boundHeight = [double]$sp.TextFrame2.TextRange.BoundHeight } catch {}
    return $f
}

# ---- slide 1: the defaults, three ways ------------------------------------
# A rectangle, a text box and a body placeholder do not have to agree, and if
# they do not then "the default inset" is a property of the shape kind rather
# than of the format.
$s1 = $pres.Slides.Add(1, $ppLayoutBlank)
$rect = New-Box $s1 'default-rect' 40 40 300 120 'Ag'
$log.defaults += Read-Frame $rect 'autoShape'

$tb = $s1.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 200, 300, 120)
$tb.Name = 'default-textbox'
$tb.TextFrame2.TextRange.Text = 'Ag'
$tb.TextFrame2.TextRange.Font.Size = 18
$tb.TextFrame2.TextRange.Font.Name = 'Arial'
$log.defaults += Read-Frame $tb 'textBox'

$s1b = $pres.Slides.Add(2, $ppLayoutText)
foreach ($sp in $s1b.Shapes) {
    $sp.TextFrame2.TextRange.Text = 'Ag'
    $log.defaults += Read-Frame $sp 'placeholder'
}

# ---- slide 3: orientation --------------------------------------------------
$s2 = $pres.Slides.Add(3, $ppLayoutBlank)
$i = 0
foreach ($o in $orientations) {
    $i++
    $entry = [ordered]@{ name = $o.name; requested = $o.value; applied = $null; error = $null; frame = $null }
    $sp = New-Box $s2 ("orient-" + $o.name) (20 + (($i - 1) % 7) * 130) 40 110 260 'Ag CJK'
    try {
        $sp.TextFrame2.Orientation = $o.value
        $entry.applied = [int]$sp.TextFrame2.Orientation
    }
    catch {
        $entry.error = $_.Exception.Message
    }
    $entry.frame = Read-Frame $sp
    $log.orientations += $entry
}

# ---- slide 4: vertical anchor ---------------------------------------------
$s3 = $pres.Slides.Add(4, $ppLayoutBlank)
$i = 0
foreach ($a in $vAnchors) {
    $i++
    $entry = [ordered]@{ name = $a.name; requested = $a.value; applied = $null; error = $null; frame = $null }
    $sp = New-Box $s3 ("vanchor-" + $a.name) (20 + ($i - 1) * 185) 40 170 300 'One'
    try {
        $sp.TextFrame2.VerticalAnchor = $a.value
        $entry.applied = [int]$sp.TextFrame2.VerticalAnchor
    }
    catch {
        $entry.error = $_.Exception.Message
    }
    $entry.frame = Read-Frame $sp
    $log.verticalAnchors += $entry
}

# ---- slide 5: horizontal anchor + word wrap --------------------------------
$s4 = $pres.Slides.Add(5, $ppLayoutBlank)
$i = 0
foreach ($a in $hAnchors) {
    $i++
    $entry = [ordered]@{ name = $a.name; requested = $a.value; applied = $null; error = $null; frame = $null }
    $sp = New-Box $s4 ("hanchor-" + $a.name) (20 + ($i - 1) * 300) 40 280 200 'One'
    try {
        $sp.TextFrame2.HorizontalAnchor = $a.value
        $entry.applied = [int]$sp.TextFrame2.HorizontalAnchor
    }
    catch {
        $entry.error = $_.Exception.Message
    }
    $entry.frame = Read-Frame $sp
    $log.horizontalAnchors += $entry
}

$wrapOff = New-Box $s4 'wrap-off' 640 40 280 200 'One'
$wrapOff.TextFrame2.WordWrap = $msoFalse
$log.horizontalAnchors += [ordered]@{ name = 'wordWrapOff'; requested = 0; applied = $null; error = $null; frame = (Read-Frame $wrapOff) }

# ---- slide 6: columns ------------------------------------------------------
$s5 = $pres.Slides.Add(6, $ppLayoutBlank)
$i = 0
foreach ($n in @(1, 2, 3)) {
    $i++
    $entry = [ordered]@{ columns = $n; spacing = $null; error = $null; frame = $null }
    $sp = New-Box $s5 ("cols-" + $n) (20 + ($i - 1) * 300) 40 280 300 'Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa'
    try {
        $sp.TextFrame2.Column.Number = $n
        if ($n -gt 1) { $sp.TextFrame2.Column.Spacing = 18 }
        $entry.spacing = [double]$sp.TextFrame2.Column.Spacing
    }
    catch {
        $entry.error = $_.Exception.Message
    }
    $entry.frame = Read-Frame $sp
    $log.columns += $entry
}

# ---- slide 7: autosize -----------------------------------------------------
$s6 = $pres.Slides.Add(7, $ppLayoutBlank)
$i = 0
foreach ($a in @(@{ name = 'none'; value = 0 }, @{ name = 'shapeToFitText'; value = 1 }, @{ name = 'textToFitShape'; value = 2 })) {
    $i++
    $entry = [ordered]@{ name = $a.name; requested = $a.value; applied = $null; error = $null; frame = $null }
    $sp = New-Box $s6 ("autosize-" + $a.name) (20 + ($i - 1) * 300) 40 280 120 'Alpha Beta Gamma Delta Epsilon Zeta Eta Theta'
    try {
        $sp.TextFrame2.AutoSize = $a.value
        $entry.applied = [int]$sp.TextFrame2.AutoSize
    }
    catch {
        $entry.error = $_.Exception.Message
    }
    $entry.frame = Read-Frame $sp
    $log.autoSize += $entry
}

# ---- slide 8: what PowerPoint reports for a frame it did not author --------
# Insets set to non-defaults through the object model, so that the saved file
# shows the units as well as the attribute names.
$s7 = $pres.Slides.Add(8, $ppLayoutBlank)
$ins = New-Box $s7 'insets-set' 40 40 400 200 'One'
$ins.TextFrame2.MarginLeft = 36
$ins.TextFrame2.MarginTop = 27
$ins.TextFrame2.MarginRight = 9
$ins.TextFrame2.MarginBottom = 0
$log.bounds += (Read-Frame $ins)

$neg = New-Box $s7 'insets-negative' 500 40 400 200 'One'
try { $neg.TextFrame2.MarginLeft = -18 } catch {}
$log.bounds += (Read-Frame $neg)

$out = Join-Path $root 'pp-frames.pptx'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
$pres.SaveAs($out, $ppSaveAsOpenXMLPresentation, $msoFalse)
$pres.Close()

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = $log | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'author-frames-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("wrote " + $out)
Write-Host ("wrote " + (Join-Path $root 'author-frames-log.json'))
