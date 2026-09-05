# Experiment T6, step 3 - ask PowerPoint where the text landed.
#
#   powershell -File tools/ground-truth/text/frames/read.ps1 -Dir <work-dir>
#
# Three instruments: TextRange2.BoundTop/BoundLeft per paragraph and per line,
# TextFrame2 own resolved properties, and an EMF export of every slide.
# BoundWidth is recorded and not trusted - see ADR 0032.
#
# Opens with OpenAndRepair:=msoFalse first so REFUSED and REPAIRED stay apart.
# Read-only; never quits a PowerPoint it did not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Only = '',
    [switch]$NoEmf
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'frame-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no frame-inputs.json in $root - run build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$emfDir = Join-Path $root 'emf'
New-Item -ItemType Directory -Force -Path $emfDir | Out-Null

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

# Every read is wrapped. A property that throws on one shape kind is a missing
# field the analysis can score; a crashed script is not a measurement.
function Read-Shape($shape) {
    $s = [ordered]@{
        id            = ''
        slide         = 0
        left          = $null
        top           = $null
        width         = $null
        height        = $null
        orientation   = $null
        vAnchor       = $null
        hAnchor       = $null
        marginLeft    = $null
        marginTop     = $null
        marginRight   = $null
        marginBottom  = $null
        wordWrap      = $null
        autoSize      = $null
        columnNumber  = $null
        columnSpacing = $null
        textLeft      = $null
        textTop       = $null
        textWidth     = $null
        textHeight    = $null
        text          = $null
        paragraphs    = @()
        lines         = @()
    }
    try { $s.id = [string]$shape.Name } catch {}
    try { $s.left = [double]$shape.Left } catch {}
    try { $s.top = [double]$shape.Top } catch {}
    try { $s.width = [double]$shape.Width } catch {}
    try { $s.height = [double]$shape.Height } catch {}

    $frame = $null
    try { $frame = $shape.TextFrame2 } catch {}
    if ($null -eq $frame) { return $s }

    try { $s.orientation = [int]$frame.Orientation } catch {}
    try { $s.vAnchor = [int]$frame.VerticalAnchor } catch {}
    try { $s.hAnchor = [int]$frame.HorizontalAnchor } catch {}
    try { $s.marginLeft = [double]$frame.MarginLeft } catch {}
    try { $s.marginTop = [double]$frame.MarginTop } catch {}
    try { $s.marginRight = [double]$frame.MarginRight } catch {}
    try { $s.marginBottom = [double]$frame.MarginBottom } catch {}
    try { $s.wordWrap = [int]$frame.WordWrap } catch {}
    try { $s.autoSize = [int]$frame.AutoSize } catch {}
    try { $s.columnNumber = [int]$frame.Column.Number } catch {}
    try { $s.columnSpacing = [double]$frame.Column.Spacing } catch {}

    $range = $null
    try { $range = $frame.TextRange } catch {}
    if ($null -eq $range) { return $s }

    try { $s.textLeft = [double]$range.BoundLeft } catch {}
    try { $s.textTop = [double]$range.BoundTop } catch {}
    try { $s.textWidth = [double]$range.BoundWidth } catch {}
    try { $s.textHeight = [double]$range.BoundHeight } catch {}
    try { $s.text = [string]$range.Text } catch {}

    $paraCount = 0
    try { $paraCount = [int]$range.Paragraphs().Count } catch {}
    for ($i = 1; $i -le $paraCount; $i++) {
        $p = $null
        try { $p = $range.Paragraphs($i, 1) } catch {}
        if ($null -eq $p) { continue }
        $entry = [ordered]@{ index = $i; left = $null; top = $null; width = $null; height = $null; text = $null }
        try { $entry.left = [double]$p.BoundLeft } catch {}
        try { $entry.top = [double]$p.BoundTop } catch {}
        try { $entry.width = [double]$p.BoundWidth } catch {}
        try { $entry.height = [double]$p.BoundHeight } catch {}
        try { $entry.text = [string]$p.Text } catch {}
        $s.paragraphs += $entry
    }

    # Lines, not paragraphs, are what an `a:br` makes and what a column break
    # moves. The two readings disagree exactly where this sub-phase is looking.
    $lineCount = 0
    try { $lineCount = [int]$range.Lines().Count } catch {}
    for ($i = 1; $i -le $lineCount; $i++) {
        $line = $null
        try { $line = $range.Lines($i, 1) } catch {}
        if ($null -eq $line) { continue }
        $entry = [ordered]@{ index = $i; left = $null; top = $null; width = $null; height = $null; text = $null }
        try { $entry.left = [double]$line.BoundLeft } catch {}
        try { $entry.top = [double]$line.BoundTop } catch {}
        try { $entry.width = [double]$line.BoundWidth } catch {}
        try { $entry.height = [double]$line.BoundHeight } catch {}
        try { $entry.text = [string]$line.Text } catch {}
        $s.lines += $entry
    }

    return $s
}

$decks = @()

foreach ($deck in $inputs.decks) {
    if ($Only -ne '' -and $deck.deck -notlike $Only) { continue }
    $file = Join-Path $root $deck.file

    $record = [ordered]@{
        deck     = $deck.deck
        file     = $deck.file
        opened   = $false
        repaired = $null
        error    = $null
        emf      = @()
        shapes   = @()
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
                foreach ($shape in $slide.Shapes) {
                    $entry = Read-Shape $shape
                    $entry.slide = $i
                    $record.shapes += $entry
                }
                if (-not $NoEmf) {
                    $emfName = ("{0}-s{1}.emf" -f $deck.deck, $i)
                    $emfPath = Join-Path $emfDir $emfName
                    if (Test-Path -LiteralPath $emfPath) { Remove-Item -LiteralPath $emfPath -Force }
                    try {
                        $slide.Export($emfPath, 'EMF', 1920, 1080)
                        $record.emf += [ordered]@{ slide = $i; file = ("emf/" + $emfName); error = $null }
                    }
                    catch {
                        $record.emf += [ordered]@{ slide = $i; file = $null; error = $_.Exception.Message }
                    }
                }
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-30} {1,-9} {2,4} shape(s) {3,3} emf" -f $deck.deck, $state, $record.shapes.Count, $record.emf.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'frame-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
