# Experiment T5, step 0 - ask PowerPoint to author the bullets and the fields.
#
#   powershell -File tools/ground-truth/text/bullets/author.ps1 -Dir <work-dir>
#
# Fourth sub-phase running that the cheapest script in this directory is the
# highest-yield, and this one settles more before a probe exists than any of
# them.
#
# ## What can be asked in writing here
#
# Three whole vocabularies, and all three are enumerations PowerPoint's own UI
# can walk:
#
# - `ST_TextAutonumberScheme` has 41 values and `PpNumberedBulletStyle` has 41
#   values. Setting each style and reading `a:buAutoNum/@type` back out of the
#   saved file is the mapping between them, written by the only party entitled
#   to define it.
# - A symbol bullet is picked from a font in the Symbol dialog, and the question
#   3.5 exists to answer is what PowerPoint writes into `a:buChar/@char` when it
#   does. `Bullet.Character` is an integer, so the round trip through it says
#   whether the U+F0xx private-use code point survives into the file or is
#   folded back to the U+00xx one it aliases.
# - A field is inserted by `InsertSlideNumber` and `InsertDateTime`, and what
#   comes out is the `@type` name *and* the cached text PowerPoint chose for
#   this machine's locale. Thirteen date formats for one loop.
#
# ## And one thing that cannot
#
# What a scheme *renders* is not in the file - `arabicPeriod` names a format,
# not a string. That is what the probe decks are for, and why this script also
# writes down the geometry PowerPoint chose (`marL`, `indent`, the bullet font)
# rather than only the vocabulary: those numbers say what a real deck looks
# like, so a probe that sets `marL="0"` knows it is asking an unusual question.
#
# Writes only into -Dir. Attaches to a running PowerPoint if there is one and
# never quits one it did not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$ppSaveAsOpenXMLPresentation = 24
$ppLayoutBlank = 12
$ppLayoutText = 2
$ppBulletUnnumbered = 1
$ppBulletNumbered = 2
$msoTextOrientationHorizontal = 1

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

$pres = $app.Presentations.Add($msoFalse)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function New-Box($slide, $name, $x, $y, $w, $h) {
    $sh = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
    $sh.Name = $name
    $tf = $sh.TextFrame2
    $tf.WordWrap = $msoFalse
    $tf.AutoSize = 0
    $tf.MarginLeft = 0
    $tf.MarginTop = 0
    $tf.MarginRight = 0
    $tf.MarginBottom = 0
    return $sh
}

# ---------------------------------------------------------------------------
# Slide 1 - the 41 numbered styles.
#
# `PpNumberedBulletStyle` is documented as 0..40 plus ppBulletMixed (-2). The
# sweep runs to 47 anyway: an accepted value outside the documented range would
# mean the enumeration in the object model is not the enumeration in the schema,
# and a rejected one inside it would mean the same thing the other way.
# ---------------------------------------------------------------------------
$slide1 = $pres.Slides.Add(1, $ppLayoutBlank)
for ($style = 0; $style -le 47; $style++) {
    $x = 20 + ($style % 6) * 155
    $y = 15 + [Math]::Floor($style / 6) * 62
    $sh = New-Box $slide1 ("num-$style") $x $y 140 55
    $tr = $sh.TextFrame2.TextRange
    $tr.Text = 'X'
    $tr.Font.Name = 'Arial'
    $tr.Font.Size = 24
    $entry = [ordered]@{ kind = 'numbered-style'; style = $style; shape = $sh.Name; accepted = $false; error = $null }
    try {
        $b = $tr.ParagraphFormat.Bullet
        $b.Visible = $msoTrue
        $b.Type = $ppBulletNumbered
        $b.Style = $style
        $b.StartValue = 1
        $entry.accepted = $true
        try { $entry.readBack = [int]$b.Style } catch {}
        try { $entry.font = [string]$b.Font.Name } catch {}
        try { $entry.relativeSize = [double]$b.RelativeSize } catch {}
        try { $entry.leftIndent = [double]$tr.ParagraphFormat.LeftIndent } catch {}
        try { $entry.firstLineIndent = [double]$tr.ParagraphFormat.FirstLineIndent } catch {}
    }
    catch {
        $entry.error = $_.Exception.Message
        # A style the object model refuses must not leave a half-set bullet in
        # the file, or the XML sweep reads the previous style under this name.
        try { $tr.ParagraphFormat.Bullet.Visible = $msoFalse } catch {}
    }
    $log += $entry
}

# ---------------------------------------------------------------------------
# Slide 2 - a symbol bullet, and the private-use question.
#
# The Symbol dialog offers Wingdings glyphs at U+F020..U+F0FF. Setting
# `Bullet.Character` to each of a handful of them and reading `a:buChar/@char`
# back out of the file says whether PowerPoint writes the private-use code point
# or the U+00xx one it aliases - which decides whether a renderer must add
# 0xF000 before looking a glyph up, or must not.
#
# Both spellings are asked of the same font, so a difference is attributable to
# the code point rather than to the face.
# ---------------------------------------------------------------------------
$slide2 = $pres.Slides.Add(2, $ppLayoutBlank)
$symbols = @(
    @{ label = 'wd-pua-a7'; font = 'Wingdings'; code = 0xF0A7 },
    @{ label = 'wd-low-a7'; font = 'Wingdings'; code = 0x00A7 },
    @{ label = 'wd-pua-6c'; font = 'Wingdings'; code = 0xF06C },
    @{ label = 'wd-low-6c'; font = 'Wingdings'; code = 0x006C },
    @{ label = 'sym-pua-b7'; font = 'Symbol'; code = 0xF0B7 },
    @{ label = 'sym-low-b7'; font = 'Symbol'; code = 0x00B7 },
    @{ label = 'arial-2022'; font = 'Arial'; code = 0x2022 },
    @{ label = 'wd2-pua-a2'; font = 'Wingdings 2'; code = 0xF0A2 },
    @{ label = 'wd3-pua-9f'; font = 'Wingdings 3'; code = 0xF09F },
    @{ label = 'webdings-a4'; font = 'Webdings'; code = 0xF0A4 }
)
$i = 0
foreach ($sym in $symbols) {
    $x = 20 + ($i % 5) * 185
    $y = 20 + [Math]::Floor($i / 5) * 90
    $sh = New-Box $slide2 ("sym-" + $sym.label) $x $y 170 70
    $tr = $sh.TextFrame2.TextRange
    $tr.Text = 'X'
    $tr.Font.Name = 'Arial'
    $tr.Font.Size = 24
    $entry = [ordered]@{ kind = 'symbol'; label = $sym.label; shape = $sh.Name; asked = $sym.code; askedFont = $sym.font; error = $null }
    try {
        $b = $tr.ParagraphFormat.Bullet
        $b.Visible = $msoTrue
        $b.Type = $ppBulletUnnumbered
        $b.Character = $sym.code
        $b.Font.Name = $sym.font
        try { $entry.readBack = [int]$b.Character } catch {}
        try { $entry.readFont = [string]$b.Font.Name } catch {}
    }
    catch { $entry.error = $_.Exception.Message }
    $log += $entry
    $i++
}

# ---------------------------------------------------------------------------
# Slide 3 - bullet size, colour and font, each set on its own.
#
# Three separate shapes rather than one, because a shape that sets all three
# cannot say which attribute PowerPoint wrote for which request.
# ---------------------------------------------------------------------------
$slide3 = $pres.Slides.Add(3, $ppLayoutBlank)
$decorations = @(
    @{ label = 'relsize-50'; act = 'relsize'; value = 0.5 },
    @{ label = 'relsize-200'; act = 'relsize'; value = 2.0 },
    @{ label = 'relsize-500'; act = 'relsize'; value = 5.0 },
    @{ label = 'colour'; act = 'colour'; value = 0x0000FF },
    @{ label = 'usetextcolor'; act = 'usetextcolor'; value = $msoTrue },
    @{ label = 'usetextfont'; act = 'usetextfont'; value = $msoTrue },
    @{ label = 'font-courier'; act = 'font'; value = 'Courier New' }
)
$i = 0
foreach ($d in $decorations) {
    $x = 20 + ($i % 4) * 230
    $y = 20 + [Math]::Floor($i / 4) * 120
    $sh = New-Box $slide3 ("dec-" + $d.label) $x $y 215 100
    $tr = $sh.TextFrame2.TextRange
    $tr.Text = 'X'
    $tr.Font.Name = 'Arial'
    $tr.Font.Size = 24
    $entry = [ordered]@{ kind = 'decoration'; label = $d.label; shape = $sh.Name; error = $null }
    try {
        $b = $tr.ParagraphFormat.Bullet
        $b.Visible = $msoTrue
        $b.Type = $ppBulletUnnumbered
        $b.Character = 0xF0B7
        $b.Font.Name = 'Symbol'
        switch ($d.act) {
            'relsize' { $b.RelativeSize = $d.value }
            'colour' { $b.Font.Fill.ForeColor.RGB = $d.value }
            'usetextcolor' { $b.UseTextColor = $d.value }
            'usetextfont' { $b.UseTextFont = $d.value }
            'font' { $b.Font.Name = $d.value }
        }
        try { $entry.relativeSize = [double]$b.RelativeSize } catch {}
        try { $entry.useTextColor = [int]$b.UseTextColor } catch {}
        try { $entry.useTextFont = [int]$b.UseTextFont } catch {}
        try { $entry.font = [string]$b.Font.Name } catch {}
    }
    catch { $entry.error = $_.Exception.Message }
    $log += $entry
    $i++
}

# ---------------------------------------------------------------------------
# Slide 4 - the fields.
#
# `InsertDateTime` and `InsertSlideNumber` live on the *legacy* `TextRange`, not
# on `TextRange2`; asking TextFrame2 for them throws. That is worth writing down
# because every other reading in this directory goes through TextFrame2.
#
# `ppDateTimeFormat` runs 1..14. The sweep runs to 16 for the same reason the
# style sweep runs past 40. Each format gets its own shape, and the cached text
# comes out with it - which is the format string for this machine's locale,
# expressed in the only way PowerPoint ever expresses it: an example.
# ---------------------------------------------------------------------------
$slide4 = $pres.Slides.Add(4, $ppLayoutBlank)
for ($fmt = 1; $fmt -le 16; $fmt++) {
    $x = 20 + (($fmt - 1) % 4) * 235
    $y = 20 + [Math]::Floor(($fmt - 1) / 4) * 90
    $sh = New-Box $slide4 ("fld-dt-$fmt") $x $y 225 70
    $tr2 = $sh.TextFrame2.TextRange
    $tr2.Font.Name = 'Arial'
    $tr2.Font.Size = 18
    $entry = [ordered]@{ kind = 'field-datetime'; format = $fmt; shape = $sh.Name; error = $null }
    try {
        $tr = $sh.TextFrame.TextRange
        $tr.Text = ''
        $null = $tr.InsertDateTime($fmt, $msoTrue)
        $entry.text = [string]$tr.Text
    }
    catch { $entry.error = $_.Exception.Message }
    $log += $entry
}
$sh = New-Box $slide4 'fld-slidenum' 20 400 225 70
$entry = [ordered]@{ kind = 'field-slidenum'; shape = $sh.Name; error = $null }
try {
    $tr = $sh.TextFrame.TextRange
    $tr.Text = ''
    $null = $tr.InsertSlideNumber()
    $entry.text = [string]$tr.Text
}
catch { $entry.error = $_.Exception.Message }
$log += $entry

# A date inserted as *text* rather than as a field, so the difference between
# the two is in one file and attributable to the flag that was flipped.
$sh = New-Box $slide4 'fld-astext' 260 400 225 70
$entry = [ordered]@{ kind = 'field-as-text'; shape = $sh.Name; error = $null }
try {
    $tr = $sh.TextFrame.TextRange
    $tr.Text = ''
    $null = $tr.InsertDateTime(2, $msoFalse)
    $entry.text = [string]$tr.Text
}
catch { $entry.error = $_.Exception.Message }
$log += $entry

# ---------------------------------------------------------------------------
# Slide 5 - the header/footer placeholders, which is where a real deck's
# `slidenum` and `datetime` fields actually live.
# ---------------------------------------------------------------------------
$slide5 = $pres.Slides.Add(5, $ppLayoutText)
$entry = [ordered]@{ kind = 'headers-footers'; slide = 5; error = $null }
try {
    $hf = $slide5.HeadersFooters
    $hf.SlideNumber.Visible = $msoTrue
    $hf.DateAndTime.Visible = $msoTrue
    $hf.DateAndTime.UseFormat = $msoTrue
    $hf.DateAndTime.Format = 2
    $hf.Footer.Visible = $msoTrue
    $hf.Footer.Text = 'footer text'
    $entry.dateText = [string]$hf.DateAndTime.Text
    $entry.footerText = [string]$hf.Footer.Text
}
catch { $entry.error = $_.Exception.Message }
$log += $entry

# ---------------------------------------------------------------------------
# Slide 6 - a multi-level numbered list, so PowerPoint writes down what it does
# with `startAt` on paragraphs after the first.
#
# The rendering is the question and this file cannot answer it. What it *can*
# answer is whether PowerPoint records a per-paragraph start value at all, or
# writes `startAt` once and leaves the sequence implicit. Those are two
# different formats, and only one of them needs a numbering pass at render time.
# ---------------------------------------------------------------------------
$slide6 = $pres.Slides.Add(6, $ppLayoutBlank)
$sh = New-Box $slide6 'list' 40 40 500 400
$tr = $sh.TextFrame2.TextRange
$tr.Text = ('one' + [char]13 + 'two' + [char]13 + 'three' + [char]13 + 'four' + [char]13 + 'five' + [char]13 + 'six')
$tr.Font.Name = 'Arial'
$tr.Font.Size = 20
for ($p = 1; $p -le 6; $p++) {
    $para = $tr.Paragraphs($p, 1)
    $b = $para.ParagraphFormat.Bullet
    $b.Visible = $msoTrue
    $b.Type = $ppBulletNumbered
    $b.Style = 3
    # `IndentLevel` is on `ParagraphFormat2`, not on the range - the legacy
    # `TextRange` carries it directly and TextRange2 does not.
    if ($p -eq 3 -or $p -eq 4) { $para.ParagraphFormat.IndentLevel = 2 }
}
$log += [ordered]@{ kind = 'multi-level-list'; shape = $sh.Name }

# ---------------------------------------------------------------------------
# Save.
# ---------------------------------------------------------------------------
$authored = Join-Path $root 'pp-bullets.pptx'
if (Test-Path -LiteralPath $authored) { Remove-Item -LiteralPath $authored -Force }
$pres.SaveAs($authored, $ppSaveAsOpenXMLPresentation, $msoFalse)
Write-Host ("wrote {0}" -f $authored)
$pres.Close()

# ---------------------------------------------------------------------------
# A stock layout, saved untouched, for the bullets PowerPoint's own masters
# carry. Only the `buChar`/`buFont`/`marL` values are read out of it, and those
# are facts about the format's conventions rather than anybody's creative work.
# ---------------------------------------------------------------------------
$pres2 = $app.Presentations.Add($msoFalse)
$pres2.PageSetup.SlideWidth = 960
$pres2.PageSetup.SlideHeight = 540
$null = $pres2.Slides.Add(1, $ppLayoutText)
$stock = Join-Path $root 'pp-stock.pptx'
if (Test-Path -LiteralPath $stock) { Remove-Item -LiteralPath $stock -Force }
$pres2.SaveAs($stock, $ppSaveAsOpenXMLPresentation, $msoFalse)
Write-Host ("wrote {0}" -f $stock)
$pres2.Close()

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ entries = $log } | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'author-bullets-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("{0} entries" -f $log.Count)
Write-Host 'done'
