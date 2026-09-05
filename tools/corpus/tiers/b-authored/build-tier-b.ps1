# Tier B of the sub-phase 1.1 corpus: the nine decks Microsoft PowerPoint wrote.
#
#   powershell -File tools/corpus/tiers/b-authored/build-tier-b.ps1 -Out <dir>
#   powershell -File tools/corpus/tiers/b-authored/build-tier-b.ps1 -Out <dir> -Id b05-chart
#
# Why this script is committed
# ----------------------------
# Tier A's decks are reproducible: `build-probes.ts --check` rebuilds every one
# byte-for-byte, so the recipe *is* the provenance. Nothing of the sort is
# possible here. PowerPoint stamps `dcterms:created` and `dcterms:modified` into
# `docProps/core.xml` on every save, so two runs one second apart differ, and
# the monthly build changes the markup underneath. A `recipe` claiming to
# reproduce these bytes would be a lie with a shelf life.
#
# What is auditable instead is the *authoring*: this file. Every deck below is a
# sequence of PowerPoint object-model calls anyone can read, re-run and compare
# against what landed in `corpus/authored/`. That is the whole reason Tier B is
# scripted rather than hand-made in the UI.
#
# What this script reads
# ----------------------
# Nothing outside `-Out`. Every presentation is created in memory with
# `Presentations.Add`; no file on this machine is opened. `b07`'s OLE object is a
# *new* empty workbook created in place. `b09`'s image is written into `-Out`
# first by `tools/corpus/tiers/b-authored/make-assets.ts`, from this repository's own
# PNG and JPEG encoders - there is no stock imagery anywhere in the corpus.
#
# The four constraints, enforced rather than trusted
# --------------------------------------------------
# 1. Blank Presentation only. `Presentations.Add` takes the built-in Office
#    Theme; no shipped design template is ever applied.
# 2. No embedded fonts. `SaveAs`'s third argument is `EmbedTrueTypeFonts`, and
#    it is 0 (msoFalse) on every call in this file.
# 3. No Designer suggestion and no stock imagery. Nothing here opens the
#    Designer pane or calls any content-insertion API that reaches a service.
# 4. No author name, company or machine name in `docProps`. See below - this
#    turned out not to work the way the roster said it would.
#
# Scrubbing personal information: measured, not assumed
# -----------------------------------------------------
# `ROSTER.md` said to scrub via `BuiltInDocumentProperties`. That is not
# reachable from PowerShell on this machine, and the failure is quiet enough to
# be worth writing down. `$pres.BuiltInDocumentProperties` returns a non-null
# object that enumerates to 34 items, but it has no usable CLR type: `.GetType()`
# throws NullReferenceException, `.Item('Author')` throws the same, and so does
# `.Name` or `.Value` on any element. `[System.__ComObject].InvokeMember` does
# not help either - it reports "Method 'System.__ComObject.Item' not found",
# because the Office type library is not bound in this host. A scrub written that
# way *appears* to run and silently changes nothing: the first version of this
# script saved a deck still carrying the author's name twice.
#
# `RemoveDocumentInformation(ppRDIRemovePersonalInformation)` does work, in one
# documented call, before the save and without editing the saved file. Measured
# on 2026-08-28 against build 16.0.20326: `dc:creator` and `cp:lastModifiedBy`
# both come back empty. It has one visible consequence, which is recorded rather
# than hidden - it writes `removePersonalInfo="1"` onto `p:presentation`, an
# attribute Tier A's chassis never emits. `decks.test.ts` asserts the result
# rather than the method, so this stays honest if the method ever changes.
#
# `dcterms:created` and `dcterms:modified` survive, and they are the reason a
# re-run does not reproduce the committed bytes. That is expected and is why
# these entries carry no `recipe`.

param(
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Id = ''
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Out)) { New-Item -ItemType Directory -Force $Out | Out-Null }
$Out = (Resolve-Path $Out).Path

# --------------------------------------------------------------- COM constants
# Named rather than inlined, because a wrong one of these is the difference
# between a deck and a two-hour bisection. `ppEffectFlashOnceFast` (3841) was the
# first value tried for a transition and PowerPoint rejected it with "this
# enumeration value is not valid for transitions": PpEntryEffect is one
# enumeration serving both shape entry animations and slide transitions, and
# neither half accepts all of it.
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1
$msoFalse = 0
$msoTrue = -1
$ppRDIRemovePersonalInformation = 4


$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5

# PpEntryEffect, transition half only.
$ppEffectFade = 1793
$ppEffectDissolve = 1537
$ppEffectWipeRight = 2819
$ppEffectPushUp = 3855
$ppEffectHoneycomb = 3898

# MsoAnimEffect / MsoAnimTriggerType.
$msoAnimEffectFade = 10
$msoAnimEffectSpin = 26
$msoAnimEffectFly = 2
$msoAnimTriggerOnPageClick = 1
$msoAnimTriggerWithPrevious = 2
$msoAnimateLevelNone = 0

# XlChartType.
$xlColumnClustered = 51
$xlLine = 4

# MsoAutoSize / MsoTriState for TextFrame2.
$msoAutoSizeTextToFitShape = 2
$msoAutoSizeNone = 0

$built = @()

function Save-Deck($pres, $name) {
    # The scrub has to happen before the save, and it is the only thing standing
    # between this corpus and the author's name in a public repository.
    $pres.RemoveDocumentInformation($script:ppRDIRemovePersonalInformation)
    $path = Join-Path $script:Out ($name + '.pptx')
    if (Test-Path $path) { Remove-Item $path -Force }
    # Third argument is EmbedTrueTypeFonts. Never anything but msoFalse here.
    $pres.SaveAs($path, $script:ppSaveAsOpenXMLPresentation, $script:msoFalse)
    $slides = $pres.Slides.Count
    $pres.Close()
    $script:built += [ordered]@{ id = $name; path = $path; slides = $slides }
    Write-Host ('  wrote ' + $name + '.pptx  (' + $slides + ' slide(s), ' +
        [math]::Round((Get-Item $path).Length / 1024, 1) + ' KiB)')
}

function Want($name) {
    return ($script:Id -eq '' -or $script:Id -eq $name)
}

$app = New-Object -ComObject PowerPoint.Application
# A repair prompt is a modal dialog: it blocks the COM call rather than
# returning an error, so this is not optional.
$app.DisplayAlerts = $ppAlertsNone

try {

    # ===================================================================== b01
    if (Want 'b01-blank') {
        Write-Host 'b01-blank'
        $p = $app.Presentations.Add($msoTrue)
        # Exactly one slide, on the Title Slide layout, with the two
        # placeholders that layout owns and nothing else. This is Tier A's
        # `a01-minimal` as Microsoft writes it, and the pair is the whole point:
        # every lexical convention that differs between the two producers shows
        # up in a diff of these two files and nowhere smaller.
        $s = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(1))
        $s.Shapes.Item(1).TextFrame.TextRange.Text = 'PPTX Studio corpus'
        $s.Shapes.Item(2).TextFrame.TextRange.Text = 'b01 - what PowerPoint writes when asked for almost nothing'
        Save-Deck $p 'b01-blank'
    }

    # ===================================================================== b02
    if (Want 'b02-layouts') {
        Write-Host 'b02-layouts'
        # One slide per built-in layout, in collection order, each labelled with
        # the layout's own name. This is the fixture sub-phase 7.1's 121-case
        # matcher matrix is built from: the eleven layouts crossed with
        # themselves, with real (type, idx) pairs rather than invented ones.
        #
        # Note there is no slide to delete first. `Presentations.Add` returns a
        # presentation with ZERO slides, not one - deleting "the auto-created
        # first slide" costs a real layout and is how this deck first came out
        # with ten slides instead of eleven.
        $p = $app.Presentations.Add($msoTrue)
        $m = $p.SlideMaster
        for ($i = 1; $i -le $m.CustomLayouts.Count; $i++) {
            $cl = $m.CustomLayouts.Item($i)
            $s = $p.Slides.AddSlide($i, $cl)
            foreach ($ph in $s.Shapes.Placeholders) {
                $t = $ph.PlaceholderFormat.Type
                # 1 title, 2 body, 3 ctrTitle, 4 subTitle, 5 vertTitle, 6 vertBody, 7 obj.
                if ($t -in @(1, 3, 5)) {
                    $ph.TextFrame.TextRange.Text = ('' + $i + '. ' + $cl.Name)
                }
                elseif ($t -in @(2, 4, 6, 7)) {
                    $ph.TextFrame.TextRange.Text =
                    ('idx ' + $ph.PlaceholderFormat.ContainedType + ', type ' + $t + "`r" +
                        'Placed by PowerPoint, not by us.')
                }
            }
        }
        Save-Deck $p 'b02-layouts'
    }

    # ===================================================================== b03
    if (Want 'b03-text') {
        Write-Host 'b03-text'
        # Autofit as PowerPoint computes it. The point of this deck is the
        # `a:normAutofit/@fontScale` and `@lnSpcReduction` pair: sub-phase 3.4
        # replays those verbatim in view mode and re-derives them in edit mode,
        # and the ladder it re-derives has to land on the same discrete steps
        # PowerPoint chose here. No synthetic deck can settle that, because the
        # ladder is what is being checked.
        $p = $app.Presentations.Add($msoTrue)

        $s1 = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(2))
        $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'b03 - autofit, shrink on overflow'
        $body = $s1.Shapes.Item(2)
        $body.TextFrame2.AutoSize = $msoAutoSizeTextToFitShape
        $body.TextFrame.TextRange.Text = (
            "A paragraph long enough to overflow the placeholder it is in, so that PowerPoint has to shrink it.`r" +
            "A second one, for the same reason.`r" +
            "A third, because the ladder has fourteen steps and one overflow only exercises the first of them.`r" +
            "A fourth, so the scale lands somewhere in the middle of the ladder rather than at its top.`r" +
            "A fifth, and by now the font scale and the line-space reduction are both in play.")

        $s2 = $p.Slides.AddSlide(2, $p.SlideMaster.CustomLayouts.Item(2))
        $s2.Shapes.Item(1).TextFrame.TextRange.Text = 'b03 - runs, bullets and a field'
        $t2 = $s2.Shapes.Item(2).TextFrame.TextRange
        $t2.Text = ("Plain.`rBold and italic in one paragraph.`rA second bullet level.`rColoured.")
        $t2.Paragraphs(2).Words(1).Font.Bold = $msoTrue
        $t2.Paragraphs(2).Words(3).Font.Italic = $msoTrue
        $t2.Paragraphs(3).IndentLevel = 2
        $t2.Paragraphs(4).Font.Color.RGB = 0x9933CC
        # A slide-number field, so `a:fld` arrives with the GUID PowerPoint
        # allocated rather than one we invented. `a13-fields` writes its own.
        $sn = $s2.HeadersFooters.SlideNumber
        $sn.Visible = $msoTrue
        $s2.HeadersFooters.Footer.Visible = $msoTrue
        $s2.HeadersFooters.Footer.Text = 'PPTX Studio corpus'

        $s3 = $p.Slides.AddSlide(3, $p.SlideMaster.CustomLayouts.Item(7))
        $tb = $s3.Shapes.AddTextbox(1, 60, 60, 400, 200)   # msoTextOrientationHorizontal
        $tb.TextFrame2.AutoSize = $msoAutoSizeNone
        $tb.TextFrame2.WordWrap = $msoTrue
        $tb.TextFrame.TextRange.Text =
        'A plain text box, which is not a placeholder and so inherits from otherStyle rather than bodyStyle.'
        Save-Deck $p 'b03-text'
    }

    # ===================================================================== b04
    if (Want 'b04-table') {
        Write-Host 'b04-table'
        # `ppt/tableStyles.xml` in every stock template is an EMPTY
        # `<a:tblStyleLst def="{5C22544A-...}"/>`, so a renderer that reads only
        # that part draws every real table white and borderless. This deck is
        # the evidence: a table carrying a built-in style id that resolves to
        # nothing in the package, plus the six banding flags that drive
        # sub-phase 4.3's thirteen-layer cascade.
        $p = $app.Presentations.Add($msoTrue)
        $s = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(6))
        $s.Shapes.Item(1).TextFrame.TextRange.Text = 'b04 - a built-in table style, and merged cells'
        $shape = $s.Shapes.AddTable(4, 4, 60, 140, 840, 300)
        $tbl = $shape.Table
        $tbl.FirstRow = $msoTrue
        $tbl.LastRow = $msoTrue
        $tbl.FirstCol = $msoTrue
        $tbl.LastCol = $msoFalse
        $tbl.HorizBanding = $msoTrue
        $tbl.VertBanding = $msoFalse
        for ($r = 1; $r -le 4; $r++) {
            for ($c = 1; $c -le 4; $c++) {
                $tbl.Cell($r, $c).Shape.TextFrame.TextRange.Text = ('r' + $r + 'c' + $c)
            }
        }
        # A horizontal merge and a vertical one, so `gridSpan`/`hMerge` and
        # `rowSpan`/`vMerge` both appear.
        $tbl.Cell(2, 2).Merge($tbl.Cell(2, 3))
        $tbl.Cell(3, 1).Merge($tbl.Cell(4, 1))
        Save-Deck $p 'b04-table'
    }

    # ===================================================================== b05
    if (Want 'b05-chart') {
        Write-Host 'b05-chart'
        # Two chart groups in one `c:plotArea` on slide 2, because
        # `firstElementChild` is the bug sub-phase 9.8 exists to catch. Also the
        # five parts a chart really costs: chart1.xml, its rels, colors1.xml,
        # style1.xml and the embedded workbook - `c:ser/c:spPr` is usually
        # absent, so a renderer that skips colors1.xml gets every series wrong.
        $p = $app.Presentations.Add($msoTrue)

        $s1 = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(6))
        $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'b05 - a clustered column chart'
        $c1 = $s1.Shapes.AddChart2(-1, $xlColumnClustered, 80, 140, 800, 340, $msoTrue)
        try { $c1.Chart.ChartData.Workbook.Application.Visible = $false } catch {}
        try { $c1.Chart.ChartData.Workbook.Close() } catch {}

        $s2 = $p.Slides.AddSlide(2, $p.SlideMaster.CustomLayouts.Item(6))
        $s2.Shapes.Item(1).TextFrame.TextRange.Text = 'b05 - a combo, so c:plotArea holds two chart groups'
        $c2 = $s2.Shapes.AddChart2(-1, $xlColumnClustered, 80, 140, 800, 340, $msoTrue)
        try {
            # Move the last series onto a line group. This is what makes the
            # plot area hold two chart-group elements rather than one.
            $c2.Chart.SeriesCollection($c2.Chart.SeriesCollection().Count).ChartType = $xlLine
        }
        catch { Write-Host ('    combo failed: ' + $_.Exception.Message.Split("`n")[0]) }
        try { $c2.Chart.ChartData.Workbook.Application.Visible = $false } catch {}
        try { $c2.Chart.ChartData.Workbook.Close() } catch {}

        Save-Deck $p 'b05-chart'
    }

    # ===================================================================== b06
    if (Want 'b06-smartart') {
        Write-Host 'b06-smartart'
        # Three families, because `dsp:` fallback resolution and `dsp:txXfrm`
        # behave differently across them, and because SmartArt is the content
        # type most likely to appear in a real corporate deck.
        $p = $app.Presentations.Add($msoTrue)
        $wanted = @(
            @{ urn = 'urn:microsoft.com/office/officeart/2005/8/layout/default'; label = 'Basic Block List' },
            @{ urn = 'urn:microsoft.com/office/officeart/2005/8/layout/orgChart1'; label = 'Organization Chart' },
            @{ urn = 'urn:microsoft.com/office/officeart/2005/8/layout/process1'; label = 'Basic Process' }
        )
        $n = 0
        foreach ($w in $wanted) {
            $n++
            $lay = $null
            for ($i = 1; $i -le $app.SmartArtLayouts.Count; $i++) {
                if ($app.SmartArtLayouts.Item($i).Id -eq $w.urn) { $lay = $app.SmartArtLayouts.Item($i); break }
            }
            if ($null -eq $lay) { Write-Host ('    no layout ' + $w.urn); $n--; continue }
            $s = $p.Slides.AddSlide($n, $p.SlideMaster.CustomLayouts.Item(6))
            $s.Shapes.Item(1).TextFrame.TextRange.Text = ('b06 - ' + $w.label)
            $sa = $s.Shapes.AddSmartArt($lay, 80, 140, 800, 340)
            try {
                $k = 0
                foreach ($node in $sa.SmartArt.AllNodes) {
                    $k++
                    $node.TextFrame2.TextRange.Text = ($w.label + ' ' + $k)
                }
            }
            catch { Write-Host ('    node text failed: ' + $_.Exception.Message.Split("`n")[0]) }
        }
        Save-Deck $p 'b06-smartart'
    }

    # ===================================================================== b07
    if (Want 'b07-ole') {
        Write-Host 'b07-ole'
        # Experiment E6 measured that this build still writes the
        # `mc:AlternateContent` wrapper but no `vmlDrawing` part and no `@spid`.
        # `a25-ole` was written from that measurement; this is the deck the
        # measurement came from, committed, so the two can be diffed rather than
        # trusted. The workbook is created empty in place - no file on this
        # machine is opened to make it.
        $p = $app.Presentations.Add($msoTrue)
        $s = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(6))
        $s.Shapes.Item(1).TextFrame.TextRange.Text = 'b07 - an embedded Excel worksheet'
        $ole = $s.Shapes.AddOLEObject(80, 140, 700, 320, 'Excel.Sheet.12')
        Write-Host ('    progId = ' + $ole.OLEFormat.ProgID)
        Save-Deck $p 'b07-ole'
    }

    # ===================================================================== b08
    if (Want 'b08-transitions') {
        Write-Host 'b08-transitions'
        # Five transitions and two animation effects a slide. The reason this
        # deck matters more than `a26-transitions` does: PowerPoint wraps
        # `p:transition` in `mc:AlternateContent`, and for a post-2010 effect the
        # `mc:Choice` holds `p14:honeycomb` while the `mc:Fallback` holds a plain
        # `p:fade`. The two branches are not the same transition. A consumer that
        # takes the fallback shows something different rather than something
        # degraded, which is exactly the failure `mc:Choice/@Requires` resolution
        # exists to prevent, written by Microsoft rather than by us.
        $p = $app.Presentations.Add($msoTrue)
        $effects = @(
            @{ n = 'Fade'; v = $ppEffectFade },
            @{ n = 'Dissolve'; v = $ppEffectDissolve },
            @{ n = 'Wipe right'; v = $ppEffectWipeRight },
            @{ n = 'Push up'; v = $ppEffectPushUp },
            @{ n = 'Honeycomb'; v = $ppEffectHoneycomb }
        )
        $i = 0
        foreach ($e in $effects) {
            $i++
            $s = $p.Slides.AddSlide($i, $p.SlideMaster.CustomLayouts.Item(6))
            $s.Shapes.Item(1).TextFrame.TextRange.Text = ('b08 - transition: ' + $e.n)
            $box = $s.Shapes.AddShape($msoShapeRoundedRectangle, 120, 200, 320, 180)
            $box.TextFrame.TextRange.Text = 'entrance'
            $box2 = $s.Shapes.AddShape($msoShapeRectangle, 520, 200, 320, 180)
            $box2.TextFrame.TextRange.Text = 'with previous'

            $s.SlideShowTransition.EntryEffect = $e.v
            $s.SlideShowTransition.Duration = 1.25
            $s.SlideShowTransition.AdvanceOnTime = $msoTrue
            $s.SlideShowTransition.AdvanceTime = 3

            $null = $s.TimeLine.MainSequence.AddEffect(
                $box, $msoAnimEffectFade, $msoAnimateLevelNone, $msoAnimTriggerOnPageClick)
            $spin = $s.TimeLine.MainSequence.AddEffect(
                $box2, $msoAnimEffectSpin, $msoAnimateLevelNone, $msoAnimTriggerWithPrevious)
            $spin.Timing.Duration = 2
            if ($i -eq 1) {
                # One fly-in too, so `p:animMotion` is not the only motion form
                # in the corpus and the main sequence has three effects on at
                # least one slide.
                $null = $s.TimeLine.MainSequence.AddEffect(
                    $s.Shapes.Item(1), $msoAnimEffectFly, $msoAnimateLevelNone, $msoAnimTriggerOnPageClick)
            }
        }
        Save-Deck $p 'b08-transitions'
    }

    # ===================================================================== b09
    if (Want 'b09-picture') {
        Write-Host 'b09-picture'
        # The image is this repository's own, written by make-assets.ts from
        # tools/corpus/tiers/a-generated/assets/png.ts and jpeg.ts. There is no stock imagery in this
        # corpus and no file outside -Out is read to make this deck.
        $pngPath = Join-Path $Out 'b09-source.png'
        $jpgPath = Join-Path $Out 'b09-source.jpeg'
        foreach ($needed in $pngPath, $jpgPath) {
            if (-not (Test-Path $needed)) {
                throw ('missing ' + $needed + ' - run `node tools/corpus/tiers/b-authored/make-assets.ts ' + $Out + '` first')
            }
        }
        $p = $app.Presentations.Add($msoTrue)

        $s1 = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(6))
        $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'b09 - crop, and the effects PowerPoint writes'
        $pic = $s1.Shapes.AddPicture($pngPath, $msoFalse, $msoTrue, 80, 150, 360, 270)
        # A crop on all four sides, so `a:srcRect` carries four non-zero
        # percentages rather than the one a simpler deck would produce.
        $pic.PictureFormat.CropLeft = 20
        $pic.PictureFormat.CropRight = 30
        $pic.PictureFormat.CropTop = 15
        $pic.PictureFormat.CropBottom = 25
        $pic2 = $s1.Shapes.AddPicture($jpgPath, $msoFalse, $msoTrue, 520, 150, 360, 270)
        $pic2.Shadow.Visible = $msoTrue
        try { $pic2.Reflection.Type = 2 } catch {}
        try { $pic2.Glow.Radius = 8 } catch {}
        try { $pic2.SoftEdge.Radius = 6 } catch {}

        $s2 = $p.Slides.AddSlide(2, $p.SlideMaster.CustomLayouts.Item(9))
        # The Picture with Caption layout, so a picture arrives in a
        # `p:ph type="pic"` rather than as a bare `p:pic` - the two inherit
        # differently and only one of them survives a layout change.
        foreach ($ph in $s2.Shapes.Placeholders) {
            $t = $ph.PlaceholderFormat.Type
            if ($t -eq 18) { $null = $ph.Fill.UserPicture($pngPath) }
            elseif ($t -eq 1) { $ph.TextFrame.TextRange.Text = 'b09 - a picture placeholder' }
            elseif ($t -eq 2) { $ph.TextFrame.TextRange.Text = 'Filled through the layout, not placed on the slide.' }
        }
        Save-Deck $p 'b09-picture'
    }

}
finally {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
}

Write-Host ''
Write-Host ('built ' + $built.Count + ' deck(s) into ' + $Out)
$built | ConvertTo-Json -Depth 3
