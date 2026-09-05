# Experiment C5, step 0 - ask PowerPoint to AUTHOR the inheritance, then read
# what it wrote.
#
#   powershell -File tools/ground-truth/model/author.ps1 -Dir <out-dir>
#
# Writes several .pptx files and author-sheets-log.json. Unzip the slide, layout
# and master parts from each and read the markup PowerPoint chose.
#
# The same move that made C3 and C4 cheap, and for this sub-phase it is worth
# more than either, because the questions are structural rather than visual:
#
#   - The eleven built-in layouts are the matcher's real input. What (type, idx)
#     pairs do they carry, in what order, with what @sz and @orient, and what
#     does the master carry? Guessing this is how a five-tier matcher ends up
#     with tiers nobody can produce a case for.
#   - A slide placeholder that the user has not moved: does PowerPoint write an
#     `a:xfrm` for it or leave it absent? The whole architecture rests on
#     `xfrm === undefined` being a fact the file states, so this is the single
#     most load-bearing thing on this page. Two shapes, one moved and one not,
#     settle it.
#   - `Shape.ShapeStyle` over its range makes PowerPoint write `p:style` with
#     `lnRef`/`fillRef`/`effectRef`/`fontRef` idx values it chose, which names
#     the style-matrix indexing without a single pixel being measured.
#   - The master's own background is written by PowerPoint, not by us. Whatever
#     `p:bgRef/@idx` it puts there is the 1000-offset stated by its author.
#   - `Slide.FollowMasterBackground`, `Slide.DisplayMasterShapes` and
#     `Presentation.Designs.Add` each write one structural fact we would
#     otherwise have to infer.
#
# Nothing here measures a pixel. It reads the file PowerPoint saved.
#
# Creates presentations and saves them into <out-dir>; opens nothing it did not
# create. Attaches to a running PowerPoint if there is one and never quits one it
# did not start.
param([Parameter(Mandatory = $true)][string]$Dir)

$ErrorActionPreference = 'Stop'
$msoTrue = -1; $msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1
$ppLayoutBlank = 12

$root = (Resolve-Path -LiteralPath $Dir).Path

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
$app.DisplayAlerts = $ppAlertsNone

$log = @()
function New-Deck { $app.Presentations.Add($msoTrue) }

# --------------------------------------------------- the built-in layouts ---
# One slide per CustomLayout of the stock design. Nothing is typed and nothing
# is moved: the point is what PowerPoint writes for a slide the user has only
# created.
$pres = New-Deck
$master = $pres.SlideMaster
$layoutCount = [int]$master.CustomLayouts.Count
$log += [ordered]@{ kind = 'masterInfo'; layouts = $layoutCount
                    width = [double]$pres.PageSetup.SlideWidth
                    height = [double]$pres.PageSetup.SlideHeight }

foreach ($shape in $master.Shapes) {
    $ph = $null
    try { $ph = [ordered]@{ type = [int]$shape.PlaceholderFormat.Type
                            contained = [int]$shape.PlaceholderFormat.ContainedType } } catch {}
    $log += [ordered]@{ kind = 'masterShape'; name = [string]$shape.Name
                        left = [double]$shape.Left; top = [double]$shape.Top
                        width = [double]$shape.Width; height = [double]$shape.Height
                        placeholder = $ph }
}

for ($i = 1; $i -le $layoutCount; $i++) {
    $layout = $master.CustomLayouts.Item($i)
    $entry = [ordered]@{ kind = 'layout'; index = $i; name = [string]$layout.Name
                         shapes = @() }
    foreach ($shape in $layout.Shapes) {
        $ph = $null
        try { $ph = [ordered]@{ type = [int]$shape.PlaceholderFormat.Type
                                contained = [int]$shape.PlaceholderFormat.ContainedType } } catch {}
        $entry.shapes += [ordered]@{ name = [string]$shape.Name
                                     left = [double]$shape.Left; top = [double]$shape.Top
                                     width = [double]$shape.Width; height = [double]$shape.Height
                                     placeholder = $ph }
    }
    $log += $entry

    $slide = $pres.Slides.AddSlide($i, $layout)
    $slide.Name = "L{0:d2}" -f $i
}
$p = Join-Path $root 'pp-layouts.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("layouts: {0}" -f $layoutCount)

# --------------------------------------- moved and unmoved, side by side ---
# Two slides on the same layout. On the first nothing is touched. On the second
# the title is nudged one point to the right and the body is left alone. If
# `a:xfrm` appears only under the nudged title, then its absence really is the
# statement "this shape has no geometry of its own", and Change Layout has
# something to read.
$pres = New-Deck
$layout = $pres.SlideMaster.CustomLayouts.Item(2)   # Title and Content
$untouched = $pres.Slides.AddSlide(1, $layout)
$untouched.Name = 'untouched'
$nudged = $pres.Slides.AddSlide(2, $layout)
$nudged.Name = 'nudged'
$nudged.Shapes.Item(1).Left = [double]$nudged.Shapes.Item(1).Left + 1
foreach ($slide in @($untouched, $nudged)) {
    foreach ($shape in $slide.Shapes) {
        $log += [ordered]@{ kind = 'geometry'; slide = [string]$slide.Name
                            name = [string]$shape.Name
                            left = [double]$shape.Left; top = [double]$shape.Top
                            width = [double]$shape.Width; height = [double]$shape.Height }
    }
}
# A third slide where text is typed but nothing is moved, because "has content"
# and "has geometry" are different questions and a renderer must not conflate
# them.
$typed = $pres.Slides.AddSlide(3, $layout)
$typed.Name = 'typed'
$typed.Shapes.Item(1).TextFrame.TextRange.Text = 'Typed, not moved'
$p = Join-Path $root 'pp-geometry.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host 'geometry: 3 slides'

# ------------------------------------------------------------ shape styles ---
# The theme's style matrix, as PowerPoint's own gallery reaches it. Each of
# these writes a `p:style` whose four refs carry the indices we would otherwise
# have to infer from ECMA's prose.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
for ($i = 1; $i -le 60; $i++) {
    try {
        $col = $n % 10; $row = [Math]::Floor($n / 10)
        $sh = $slide.Shapes.AddShape(1, 12 + $col * 66, 12 + $row * 52, 52, 40)
        $sh.ShapeStyle = $i
        $sh.Name = "style{0:d2}" -f $i
        $n++
        $log += [ordered]@{ kind = 'shapeStyle'; index = $i; ok = $true; error = $null }
    }
    catch {
        $log += [ordered]@{ kind = 'shapeStyle'; index = $i; ok = $false; error = $_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}
$p = Join-Path $root 'pp-shapestyles.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host ("shape styles: {0} accepted of 60" -f $n)

# ------------------------------------------------ backgrounds and master sp ---
# Four slides: inherit, own solid, own gradient, master shapes hidden. The
# first is the control - a slide with no opinion should carry no `p:bg` at all.
$pres = New-Deck
$s1 = $pres.Slides.Add(1, $ppLayoutBlank); $s1.Name = 'inherit'
$s2 = $pres.Slides.Add(2, $ppLayoutBlank); $s2.Name = 'ownSolid'
$s2.FollowMasterBackground = $msoFalse
$s2.Background.Fill.Solid()
$s2.Background.Fill.ForeColor.RGB = 3243501
$s3 = $pres.Slides.Add(3, $ppLayoutBlank); $s3.Name = 'ownGradient'
$s3.FollowMasterBackground = $msoFalse
$s3.Background.Fill.TwoColorGradient(1, 1)
$s4 = $pres.Slides.Add(4, $ppLayoutBlank); $s4.Name = 'noMasterShapes'
$s4.DisplayMasterShapes = $msoFalse
foreach ($slide in @($s1, $s2, $s3, $s4)) {
    $log += [ordered]@{ kind = 'background'; slide = [string]$slide.Name
                        follow = [int]$slide.FollowMasterBackground
                        displayMasterShapes = [int]$slide.DisplayMasterShapes
                        fillType = [int]$slide.Background.Fill.Type }
}
$p = Join-Path $root 'pp-backgrounds.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host 'backgrounds: 4 slides'

# ------------------------------------------------------------ two masters ---
# `Designs.Add` is how PowerPoint's UI produces a second master. What we want
# from it is the shape of the package: how many themes, which part owns which
# relationship, and whether the second master's layouts are numbered from one
# or continue the first's.
$pres = New-Deck
$pres.Slides.Add(1, $ppLayoutBlank).Name = 'onMasterOne'
$design = $pres.Designs.Add('Second')
$second = $pres.Slides.Add(2, $ppLayoutBlank)
$second.Name = 'onMasterTwo'
$second.Design = $design
$log += [ordered]@{ kind = 'designs'; count = [int]$pres.Designs.Count
                    masters = [int]$pres.Designs.Count
                    layoutsOne = [int]$pres.Designs.Item(1).SlideMaster.CustomLayouts.Count
                    layoutsTwo = [int]$pres.Designs.Item(2).SlideMaster.CustomLayouts.Count }
$p = Join-Path $root 'pp-two-masters.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host 'two masters'

# ---------------------------------------------------------- change layout ---
# Not built in this sub-phase - Change Layout is 7.4 - but what PowerPoint
# rewrites when it happens is cheap to capture now and expensive to reproduce
# later. Two slides, identical content, one switched to a layout whose body
# placeholder has a different @idx.
$pres = New-Deck
$layoutA = $pres.SlideMaster.CustomLayouts.Item(2)   # Title and Content
$layoutB = $pres.SlideMaster.CustomLayouts.Item(4)   # Two Content
$before = $pres.Slides.AddSlide(1, $layoutA); $before.Name = 'before'
$before.Shapes.Item(1).TextFrame.TextRange.Text = 'Title'
$before.Shapes.Item(2).TextFrame.TextRange.Text = 'Body'
$after = $pres.Slides.AddSlide(2, $layoutA); $after.Name = 'after'
$after.Shapes.Item(1).TextFrame.TextRange.Text = 'Title'
$after.Shapes.Item(2).TextFrame.TextRange.Text = 'Body'
$after.CustomLayout = $layoutB
foreach ($slide in @($before, $after)) {
    foreach ($shape in $slide.Shapes) {
        $ph = $null
        try { $ph = [int]$shape.PlaceholderFormat.Type } catch {}
        $log += [ordered]@{ kind = 'changeLayout'; slide = [string]$slide.Name
                            name = [string]$shape.Name; placeholderType = $ph
                            left = [double]$shape.Left; top = [double]$shape.Top }
    }
}
$p = Join-Path $root 'pp-change-layout.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host 'change layout: 2 slides'

# ------------------------------------------------------ headers and footers ---
# `p:hf` lives on the master and the date/footer/slide-number placeholders live
# on every layout. Turning them on is the only way to see which of the two
# actually gates the rendering.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank); $slide.Name = 'hf'
$slide.HeadersFooters.Footer.Visible = $msoTrue
$slide.HeadersFooters.Footer.Text = 'footer text'
$slide.HeadersFooters.SlideNumber.Visible = $msoTrue
$slide.HeadersFooters.DateAndTime.Visible = $msoTrue
$slide.HeadersFooters.DateAndTime.UseFormat = $msoTrue
$log += [ordered]@{ kind = 'headerFooter'; shapes = [int]$slide.Shapes.Count }
$p = Join-Path $root 'pp-headers.pptx'
$pres.SaveAs($p, $ppSaveAsOpenXMLPresentation); $pres.Close()
Write-Host 'headers/footers'

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{
    files = @('pp-layouts.pptx', 'pp-geometry.pptx', 'pp-shapestyles.pptx', 'pp-backgrounds.pptx',
              'pp-two-masters.pptx', 'pp-change-layout.pptx', 'pp-headers.pptx')
    log   = $log
} | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $root 'author-sheets-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
