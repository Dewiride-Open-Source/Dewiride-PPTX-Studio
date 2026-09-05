# Experiment C4, step 0 - ask PowerPoint to AUTHOR lines and effects, then read
# what it wrote.
#
#   powershell -File tools/ground-truth/paint/lines/author.ps1 -Dir <out-dir>
#
# Writes several .pptx files and author-lines-log.json. Unzip
# ppt/slides/slide1.xml from each and read the markup PowerPoint chose.
#
# This is the same move that made C3 cheap, and lines have far more enumerations
# than fills did, so there is more of it to collect:
#
#   - Line.DashStyle over the MsoLineDashStyle range names the preset dashes
#     PowerPoint's own UI can produce. OOXML defines eleven; the object model
#     exposes fewer, and which four it cannot reach is worth knowing before
#     deciding what a renderer must get right.
#   - Line.Style names the five compound strokes, and Line.InsetPen the two
#     alignments.
#   - Line.BeginArrowheadStyle/Length/Width name the arrowhead vocabulary.
#   - Shadow.Type over msoShadow1..msoShadow43, Glow, SoftEdge and Reflection
#     make PowerPoint write the effect attribute sets itself. a:reflection alone
#     has thirteen attributes with no defaults worth guessing; a preset that
#     PowerPoint authored names all of them at once.
#
# Nothing here measures a pixel. It reads the file PowerPoint saved, which is a
# different and much cheaper question than what PowerPoint paints - and it is the
# question that tells the probe set what to ask.
#
# Creates presentations and saves them into <out-dir>; opens nothing it did not
# create. Attaches to a running PowerPoint if there is one and never quits one it
# did not start.
param([Parameter(Mandatory = $true)][string]$Dir)

$ErrorActionPreference = 'Stop'
$msoTrue = -1; $msoFalse = 0
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1

$root = (Resolve-Path -LiteralPath $Dir).Path

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
$app.DisplayAlerts = $ppAlertsNone

$log = @()
function New-Deck { $app.Presentations.Add($msoTrue) }

# ------------------------------------------------------------------ dashes ---
# A plain straight connector per dash style. Also records the line's default
# state before anything is set, which is the answer to "what does PowerPoint
# write when the user has expressed no opinion".
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$probe = $slide.Shapes.AddLine(20, 20, 300, 20)
$probe.Name = 'untouched'
$log += [ordered]@{ kind='default'; weight=[double]$probe.Line.Weight; dash=[int]$probe.Line.DashStyle;
                    style=[int]$probe.Line.Style; insetPen=[int]$probe.Line.InsetPen }
$n = 1
for ($i = 1; $i -le 12; $i++) {
    try {
        $y = 40 + $n * 22
        $sh = $slide.Shapes.AddLine(20, $y, 400, $y)
        $sh.Line.DashStyle = $i
        $sh.Line.Weight = 3
        $sh.Name = "dash{0:d2}" -f $i
        $n++
        $log += [ordered]@{ kind='dash'; index=$i; ok=$true; error=$null }
    } catch {
        $log += [ordered]@{ kind='dash'; index=$i; ok=$false; error=$_.Exception.Message }
    }
}
$p = Join-Path $root 'pp-dashes.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("dashes: {0} accepted of 12" -f ($n - 1))

# ------------------------------------------------- compound, cap, alignment ---
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
for ($i = 1; $i -le 6; $i++) {
    try {
        $y = 20 + $n * 30
        $sh = $slide.Shapes.AddLine(20, $y, 400, $y)
        $sh.Line.Style = $i
        $sh.Line.Weight = 9
        $sh.Name = "cmpd{0}" -f $i
        $n++
        $log += [ordered]@{ kind='cmpd'; index=$i; ok=$true; error=$null }
    } catch { $log += [ordered]@{ kind='cmpd'; index=$i; ok=$false; error=$_.Exception.Message } }
}
# InsetPen is the object model's name for a:ln/@algn. It is only meaningful on a
# closed shape, so this half uses rectangles.
foreach ($inset in @($msoFalse, $msoTrue)) {
    try {
        $sh = $slide.Shapes.AddShape(1, 450, 20 + $n * 10, 120, 80)
        $sh.Line.Weight = 12
        $sh.Line.InsetPen = $inset
        $sh.Name = "inset{0}" -f $inset
        $n++
        $log += [ordered]@{ kind='insetPen'; value=$inset; ok=$true; error=$null }
    } catch { $log += [ordered]@{ kind='insetPen'; value=$inset; ok=$false; error=$_.Exception.Message } }
}
$p = Join-Path $root 'pp-compound.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("compound/inset: {0}" -f $n)

# -------------------------------------------------------------- arrowheads ---
# Every head type at every length and width, one line each. 6 x 3 x 3 = 54 if
# they are all accepted.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
for ($type = 1; $type -le 7; $type++) {
    for ($len = 1; $len -le 3; $len++) {
        for ($wid = 1; $wid -le 3; $wid++) {
            try {
                $col = $n % 9; $row = [Math]::Floor($n / 9)
                $y = 15 + $row * 26
                $x = 15 + $col * 78
                $sh = $slide.Shapes.AddLine($x, $y, $x + 66, $y)
                $sh.Line.Weight = 4
                $sh.Line.EndArrowheadStyle = $type
                $sh.Line.EndArrowheadLength = $len
                $sh.Line.EndArrowheadWidth = $wid
                $sh.Name = "ah-t{0}-l{1}-w{2}" -f $type, $len, $wid
                $n++
                $log += [ordered]@{ kind='arrowhead'; type=$type; len=$len; width=$wid; ok=$true; error=$null }
            } catch {
                $log += [ordered]@{ kind='arrowhead'; type=$type; len=$len; width=$wid; ok=$false; error=$_.Exception.Message }
            }
        }
    }
}
$p = Join-Path $root 'pp-arrowheads.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("arrowheads: {0} accepted of 63" -f $n)

# ------------------------------------------------------------------ shadows ---
# msoShadow1 .. msoShadow43. Each writes an a:outerShdw (or a:prstShdw, or an
# inner shadow) with attributes PowerPoint chose - which is the whole point.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
for ($i = 1; $i -le 43; $i++) {
    try {
        $col = $n % 9; $row = [Math]::Floor($n / 9)
        $sh = $slide.Shapes.AddShape(1, 15 + $col * 74, 15 + $row * 62, 56, 46)
        $sh.Shadow.Type = $i
        $sh.Name = "shadow{0:d2}" -f $i
        $n++
        $log += [ordered]@{ kind='shadow'; index=$i; ok=$true; error=$null }
    } catch {
        $log += [ordered]@{ kind='shadow'; index=$i; ok=$false; error=$_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}
$p = Join-Path $root 'pp-shadows.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("shadows: {0} accepted of 43" -f $n)

# --------------------------------------------- glow, soft edge, reflection ---
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
foreach ($radius in @(4, 8, 16, 32)) {
    try {
        $sh = $slide.Shapes.AddShape(1, 15 + $n * 90, 20, 70, 50)
        $sh.Glow.Radius = $radius
        $sh.Glow.Color.RGB = 12874308
        $sh.Name = "glow{0}" -f $radius
        $n++
        $log += [ordered]@{ kind='glow'; radius=$radius; ok=$true; error=$null }
    } catch { $log += [ordered]@{ kind='glow'; radius=$radius; ok=$false; error=$_.Exception.Message } }
}
$m = 0
for ($i = 0; $i -le 6; $i++) {
    try {
        $sh = $slide.Shapes.AddShape(1, 15 + $m * 90, 100, 70, 50)
        $sh.SoftEdge.Type = $i
        $sh.Name = "soft{0}" -f $i
        $m++
        $log += [ordered]@{ kind='softEdge'; index=$i; radius=[double]$sh.SoftEdge.Radius; ok=$true; error=$null }
    } catch { $log += [ordered]@{ kind='softEdge'; index=$i; ok=$false; error=$_.Exception.Message } }
}
$k = 0
for ($i = 1; $i -le 10; $i++) {
    try {
        $sh = $slide.Shapes.AddShape(1, 15 + $k * 90, 190, 70, 50)
        $sh.Reflection.Type = $i
        $sh.Name = "refl{0:d2}" -f $i
        $k++
        $log += [ordered]@{ kind='reflection'; index=$i; ok=$true; error=$null }
    } catch {
        $log += [ordered]@{ kind='reflection'; index=$i; ok=$false; error=$_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}
$p = Join-Path $root 'pp-effects.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("glow {0}, softEdge {1}, reflection {2}" -f $n, $m, $k)

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{
    files = @('pp-dashes.pptx','pp-compound.pptx','pp-arrowheads.pptx','pp-shadows.pptx','pp-effects.pptx')
    log = $log
} | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText((Join-Path $root 'author-lines-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "done"
