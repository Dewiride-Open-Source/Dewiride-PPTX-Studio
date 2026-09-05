# Experiment T4, step 3 - make PowerPoint compute an autofit, then read what it
# wrote.
#
#   powershell -File tools/ground-truth/text/autofit/run.ps1 -Dir <work-dir>
#
# Reads `autofit-inputs.json`, and for each deck does one of two things.
#
# ## Recompute decks
#
# `TextFrame2.AutoSize = msoAutoSizeNone` and then back to
# `msoAutoSizeTextToFitShape` is the object model's way of saying "do the thing
# a user does by typing". Setting it to None first matters: assigning the value
# a frame already holds is a no-op, and the deck is authored with
# `<a:normAutofit/>` already in place, so without the round trip through None
# nothing would be recomputed and every probe would come back at 100% - a
# result that looks like a measurement and is an artefact of the trigger.
#
# The deck is then saved under a new name and the analysis reads the emitted
# `@fontScale` and `@lnSpcReduction` out of the XML. That is the whole point:
# the numbers are PowerPoint's own, in PowerPoint's own file, and nothing here
# has to infer them from a rendering.
#
# ## View decks
#
# Opened read-only, read, and closed without saving. These carry stored scales
# that no recomputation would produce, and the question is what PowerPoint does
# with a file it did not write. Touching them at all would destroy the
# measurement, so the AutoSize round trip is skipped and the file is never
# saved.
#
# ## What is read
#
# Line count and the text block's bounds everywhere; per-line positions only on
# the view decks, where the baseline-to-baseline advance is the measurement.
# Reading 20 lines from each of 429 ladder probes over COM would cost minutes
# and answer nothing the emitted scale does not answer better.
#
# Attaches to a running PowerPoint if there is one and never quits one it did
# not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Only = '',
    # Open without a window. Kept because it is the measurement that found the
    # trigger: with `-NoWindow` every one of the 20 inset probes came back
    # `<a:normAutofit/>`, unshrunk, however hard the frame was poked. Autofit
    # runs in PowerPoint's layout path, and the layout path does not exist
    # without a window. A headless COM harness measures nothing here and looks
    # like it measured "PowerPoint never shrinks".
    [switch]$NoWindow,
    # Append a character and delete it again. The frame ends up holding exactly
    # the text it started with, having been edited twice, which is the event a
    # user generates by typing. Measured to change nothing either way, so it is
    # off by default - the AutoSize round trip below is what marks the frame,
    # and the window is what lets the mark be acted on.
    [switch]$TouchText,
    # Minimise PowerPoint while it works. Verified not to change a single
    # emitted scale on the inset deck, and it keeps a 1107-probe run from
    # throwing windows at whoever is at the keyboard.
    [switch]$Restore
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3
$ppSaveAsOpenXMLPresentation = 24

$msoAutoSizeNone = 0
$msoAutoSizeShapeToFitText = 1
$msoAutoSizeTextToFitShape = 2

# Every output of a variant run is named after it, so the four combinations of
# the two switches can sit in one directory and be compared rather than
# overwriting each other.
$variant = ''
if ($NoWindow) { $variant = $variant + '-nowin' }
if ($TouchText) { $variant = $variant + '-touch' }
if ($Restore) { $variant = $variant + '-restore' }

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'autofit-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no autofit-inputs.json in $root - run tools/ground-truth/text/autofit/build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

# Which autofit each probe asks for, so the runner knows whether to drive the
# frame to TextToFitShape or to ShapeToFitText. Keyed by shape name.
$wanted = @{}
foreach ($p in $inputs.probes) { $wanted[$p.id] = $p.autofit }

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

# ppWindowMinimized. Set before any deck is opened so no window is ever painted
# at full size. The application still lays out - which is the only thing the
# window is here for - and the emitted scales were checked against a run without
# this line.
$ppWindowMinimized = 2
if (-not $NoWindow -and -not $Restore) {
    try { $app.WindowState = $ppWindowMinimized } catch {}
}

function Read-Shape($shape, [bool]$withLines) {
    $name = ''
    try { $name = [string]$shape.Name } catch {}

    $s = [ordered]@{
        id          = $name
        shapeWidth  = $null
        shapeHeight = $null
        autoSize    = $null
        wordWrap    = $null
        lineCount   = $null
        paraCount   = $null
        fontSize    = $null
        fontName    = $null
        boundTop    = $null
        boundHeight = $null
        boundWidth  = $null
        lines       = @()
    }

    try { $s.shapeWidth = [double]$shape.Width } catch {}
    try { $s.shapeHeight = [double]$shape.Height } catch {}

    $frame = $null
    try { $frame = $shape.TextFrame2 } catch {}
    if ($null -eq $frame) { return $s }

    try { $s.autoSize = [int]$frame.AutoSize } catch {}
    try { $s.wordWrap = [int]$frame.WordWrap } catch {}

    $range = $null
    try { $range = $frame.TextRange } catch {}
    if ($null -eq $range) { return $s }

    try { $s.boundTop = [double]$range.BoundTop } catch {}
    try { $s.boundHeight = [double]$range.BoundHeight } catch {}
    try { $s.boundWidth = [double]$range.BoundWidth } catch {}
    try { $s.paraCount = [int]$range.Paragraphs().Count } catch {}
    try { $s.fontName = [string]$range.Font.Name } catch {}
    # The nominal size, not the shrunk one - PowerPoint reports what the run
    # states. Recorded so that a probe whose run did not say what it meant to
    # say is visible rather than absorbed into the ladder.
    try { $s.fontSize = [double]$range.Font.Size } catch {}

    $lineCount = 0
    try { $lineCount = [int]$range.Lines().Count } catch {}
    $s.lineCount = $lineCount

    if ($withLines) {
        for ($i = 1; $i -le $lineCount; $i++) {
            $line = $null
            try { $line = $range.Lines($i, 1) } catch {}
            if ($null -eq $line) { continue }
            $entry = [ordered]@{ index = $i; text = $null; top = $null; height = $null; length = $null }
            try { $entry.text = [string]$line.Text } catch {}
            try { $entry.top = [double]$line.BoundTop } catch {}
            try { $entry.height = [double]$line.BoundHeight } catch {}
            try { $entry.length = [int]$line.Length } catch {}
            $s.lines += $entry
        }
    }

    return $s
}

$decks = @()

foreach ($deck in $inputs.decks) {
    if ($Only -ne '' -and $deck.deck -ne $Only) { continue }

    $file = Join-Path $root $deck.file
    $recompute = [bool]$deck.recompute

    $record = [ordered]@{
        deck      = $deck.deck
        file      = $deck.file
        savedAs   = $null
        recompute = $recompute
        opened    = $false
        repaired  = $null
        error     = $null
        shapes    = @()
    }

    $pres = $null
    $readOnly = if ($recompute) { $msoFalse } else { $msoTrue }
    $window = if ($NoWindow) { $msoFalse } else { $msoTrue }
    try {
        # FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair
        $pres = $app.Presentations.Open2007($file, $readOnly, $msoFalse, $window, $msoFalse)
        $record.opened = $true
        $record.repaired = $false
    }
    catch {
        $record.error = $_.Exception.Message
        try {
            $pres = $app.Presentations.Open2007($file, $readOnly, $msoFalse, $window, $msoTrue)
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
                foreach ($shape in $slide.Shapes) {
                    if ($recompute) {
                        $name = ''
                        try { $name = [string]$shape.Name } catch {}
                        $want = $wanted[$name]
                        $target = if ($want -eq 'sp') { $msoAutoSizeShapeToFitText } else { $msoAutoSizeTextToFitShape }
                        if ($want -ne 'none') {
                            try {
                                $shape.TextFrame2.AutoSize = $msoAutoSizeNone
                                $shape.TextFrame2.AutoSize = $target
                            }
                            catch {}
                            if ($TouchText) {
                                try {
                                    $shape.TextFrame2.TextRange.InsertAfter('x') | Out-Null
                                    $r2 = $shape.TextFrame2.TextRange
                                    $r2.Characters($r2.Length, 1).Delete()
                                }
                                catch {}
                            }
                        }
                    }
                    $record.shapes += Read-Shape $shape (-not $recompute)
                }
            }

            if ($recompute) {
                $out = Join-Path $root ($deck.deck + $variant + '-out.pptx')
                if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
                $pres.SaveAs($out, $ppSaveAsOpenXMLPresentation, $msoFalse)
                $record.savedAs = ($deck.deck + $variant + '-out.pptx')
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-16} {1,-9} {2} shape(s)" -f $deck.deck, $state, $record.shapes.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$suffix = if ($Only -ne '') { '-' + $Only } else { '' }
$suffix = $suffix + $variant
$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root ('autofit-readings' + $suffix + '.json')), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
