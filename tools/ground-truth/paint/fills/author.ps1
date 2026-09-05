# Experiment C3, step 0 - ask PowerPoint to AUTHOR fills, then read what it wrote.
#
#   powershell -File tools/ground-truth/paint/fills/author.ps1 -Dir <out-dir>
#
# Writes three .pptx files and author-log.json. Unzip ppt/slides/slide1.xml from
# each and read the markup PowerPoint chose.
#
# This is the cheapest script in the directory and it settled three questions
# before a single probe existed:
#
#   - Fill.Patterned(i) over the whole MsoPatternType range is accepted for
#     exactly 54 values, and PowerPoint spells each one into a:pattFill/@prst
#     itself - so the enumeration in fills.ts is PowerPoint's own rather than a
#     transcription of a schema nobody here has read.
#   - a two-colour gradient is written with TWO stops. Not 33. The plan expected
#     a pre-sampled gamma ramp in the file; since the authoring UI does not put
#     one there, if the ramp is real it belongs to the renderer and only a bitmap
#     can say.
#   - PowerPoint writes a:gsLst OUT OF @pos ORDER. Its own from-centre variant is
#     pos="50000", pos="0", pos="100000". A renderer that trusts document order is
#     wrong on a file PowerPoint wrote.
#
# The general lesson is worth more than any of the three. When the question is
# "what does this format mean", the application that defines the format can often
# be made to answer in writing, and that answer costs minutes rather than a
# measurement rig.
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

# ---------------------------------------------------------------- patterns ---
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$ok = 0
for ($i = 1; $i -le 60; $i++) {
    $col = ($ok % 10); $row = [Math]::Floor($ok / 10)
    try {
        $sh = $slide.Shapes.AddShape(1, 10 + $col * 94, 10 + $row * 88, 90, 84)
        $sh.Fill.Patterned($i)
        $sh.Fill.ForeColor.RGB = 0          # black, BGR
        $sh.Fill.BackColor.RGB = 16777215   # white
        $sh.Line.Visible = $msoFalse
        $sh.Name = "pat{0:d2}" -f $i
        $ok++
        $log += [ordered]@{ kind='pattern'; index=$i; ok=$true; error=$null }
    } catch {
        $log += [ordered]@{ kind='pattern'; index=$i; ok=$false; error=$_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}
$p1 = Join-Path $root 'pp-patterns.pptx'
$pres.SaveAs($p1, $ppSaveAsOpenXMLPresentation)
$pres.Close()
Write-Host ("patterns: {0} accepted of 60" -f $ok)

# --------------------------------------------------------------- gradients ---
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
# TwoColorGradient(Style, Variant) across all 7 styles x 4 variants
for ($style = 1; $style -le 7; $style++) {
    for ($variant = 1; $variant -le 4; $variant++) {
        $col = ($n % 7); $row = [Math]::Floor($n / 7)
        try {
            $sh = $slide.Shapes.AddShape(1, 10 + $col * 134, 10 + $row * 100, 128, 94)
            $sh.Fill.TwoColorGradient($style, $variant)
            $sh.Fill.ForeColor.RGB = 0
            $sh.Fill.BackColor.RGB = 16777215
            $sh.Line.Visible = $msoFalse
            $sh.Name = "two-s{0}-v{1}" -f $style, $variant
            $n++
            $log += [ordered]@{ kind='twoColor'; style=$style; variant=$variant; ok=$true; error=$null }
        } catch {
            $log += [ordered]@{ kind='twoColor'; style=$style; variant=$variant; ok=$false; error=$_.Exception.Message }
            try { if ($sh) { $sh.Delete() } } catch {}
        }
    }
}
$p2 = Join-Path $root 'pp-gradients-two.pptx'
$pres.SaveAs($p2, $ppSaveAsOpenXMLPresentation)
$pres.Close()
Write-Host ("two-colour gradients: {0}" -f $n)

# ------------------------------------------------- one-colour + preset ramps ---
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
foreach ($degree in @(0.0, 0.25, 0.5, 0.75, 1.0)) {
    $col = ($n % 7); $row = [Math]::Floor($n / 7)
    try {
        $sh = $slide.Shapes.AddShape(1, 10 + $col * 134, 10 + $row * 100, 128, 94)
        $sh.Fill.ForeColor.RGB = 12874308   # 0x4472C4 in BGR = C47244
        $sh.Fill.OneColorGradient(1, 1, $degree)
        $sh.Line.Visible = $msoFalse
        $sh.Name = "one-d{0}" -f ($degree * 100)
        $n++
        $log += [ordered]@{ kind='oneColor'; degree=$degree; ok=$true; error=$null }
    } catch { $log += [ordered]@{ kind='oneColor'; degree=$degree; ok=$false; error=$_.Exception.Message } }
}
for ($pg = 1; $pg -le 24; $pg++) {
    $col = ($n % 7); $row = [Math]::Floor($n / 7)
    try {
        $sh = $slide.Shapes.AddShape(1, 10 + $col * 134, 10 + $row * 100, 128, 94)
        $sh.Fill.PresetGradient(1, 1, $pg)
        $sh.Line.Visible = $msoFalse
        $sh.Name = "prst{0:d2}" -f $pg
        $n++
        $log += [ordered]@{ kind='presetGradient'; index=$pg; ok=$true; error=$null }
    } catch {
        $log += [ordered]@{ kind='presetGradient'; index=$pg; ok=$false; error=$_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}
$p3 = Join-Path $root 'pp-gradients-preset.pptx'
$pres.SaveAs($p3, $ppSaveAsOpenXMLPresentation)
$pres.Close()
Write-Host ("one-colour + preset: {0}" -f $n)

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ files = @('pp-patterns.pptx','pp-gradients-two.pptx','pp-gradients-preset.pptx'); log = $log } | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText((Join-Path $root 'author-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "done"
