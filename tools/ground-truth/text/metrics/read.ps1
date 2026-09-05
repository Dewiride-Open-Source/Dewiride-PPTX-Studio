# Experiment T2, step 2 - ask PowerPoint how wide and how tall.
#
#   powershell -File tools/ground-truth/text/metrics/read.ps1 -Dir <work-dir>
#
# Reads `metric-inputs.json`, opens each probe deck, and for every shape records
# what the object model says the laid-out text measures.
#
# ## Why `TextRange2` and not a bitmap
#
# `TextRange2.BoundWidth` and `.BoundHeight` are the numbers PowerPoint's own
# layout arrived at, reported in points as a float. A screenshot would give the
# same information at a worse resolution, through a rasteriser, after
# anti-aliasing - and would need a threshold to be read back, which is one more
# thing to get wrong. There is nothing here a pixel could answer better.
#
# ## Why per line
#
# `Lines($i, 1).BoundTop` is the top of line i. The difference between
# consecutive tops is the line advance *directly*, with no first-line padding to
# cancel and no assumption that the advance is even constant. The whole-range
# height is recorded too, so the two can be checked against each other rather
# than one being trusted.
#
# `Open2007` with `OpenAndRepair:=msoFalse` first, then with repair allowed, so
# REFUSED and REPAIRED are distinguishable - C2 found that plain `Open` repairs
# silently and reports success on exactly the files being asked about.
#
# Read-only. Nothing is written back to any deck. Attaches to a running
# PowerPoint if there is one and never quits one it did not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'metric-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no metric-inputs.json in $root - run tools/ground-truth/text/metrics/build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

# Which probes want the expensive per-character walk. A hash rather than a
# linear search, because it is consulted once per shape.
$perChar = @{}
foreach ($p in $inputs.probes) {
    if ($p.perChar) { $perChar[[string]$p.id] = $true }
}

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
    $name = ''
    try { $name = [string]$shape.Name } catch {}

    $s = [ordered]@{
        id          = $name
        shapeLeft   = $null
        shapeTop    = $null
        shapeWidth  = $null
        shapeHeight = $null
        width       = $null
        height      = $null
        left        = $null
        top         = $null
        lineCount   = $null
        paraCount   = $null
        fontName    = $null
        fontSize    = $null
        lines       = @()
        chars       = @()
    }

    try { $s.shapeLeft = [double]$shape.Left } catch {}
    try { $s.shapeTop = [double]$shape.Top } catch {}
    try { $s.shapeWidth = [double]$shape.Width } catch {}
    try { $s.shapeHeight = [double]$shape.Height } catch {}

    $range = $null
    try { $range = $shape.TextFrame2.TextRange } catch {}
    if ($null -eq $range) { return $s }

    try { $s.width = [double]$range.BoundWidth } catch {}
    try { $s.height = [double]$range.BoundHeight } catch {}
    try { $s.left = [double]$range.BoundLeft } catch {}
    try { $s.top = [double]$range.BoundTop } catch {}
    try { $s.paraCount = [int]$range.Paragraphs().Count } catch {}

    # `Font.Name` is the resolved face. It does not report a substitution - a
    # missing face still reports the name that was asked for - which is why the
    # analysis detects substitution from the widths instead. It is recorded
    # anyway, because a face that reports something *else* would mean the probe
    # never said what it meant to say.
    try { $s.fontName = [string]$range.Font.Name } catch {}
    try { $s.fontSize = [double]$range.Font.Size } catch {}

    $lineCount = 0
    try { $lineCount = [int]$range.Lines().Count } catch {}
    $s.lineCount = $lineCount
    for ($i = 1; $i -le $lineCount; $i++) {
        $line = $null
        try { $line = $range.Lines($i, 1) } catch {}
        if ($null -eq $line) { continue }
        $entry = [ordered]@{ index = $i; top = $null; height = $null; left = $null; width = $null; text = $null }
        try { $entry.top = [double]$line.BoundTop } catch {}
        try { $entry.height = [double]$line.BoundHeight } catch {}
        try { $entry.left = [double]$line.BoundLeft } catch {}
        try { $entry.width = [double]$line.BoundWidth } catch {}
        try { $entry.text = [string]$line.Text } catch {}
        $s.lines += $entry
    }

    if ($perChar.ContainsKey($name)) {
        $n = 0
        try { $n = [int]$range.Characters().Count } catch {}
        for ($i = 1; $i -le $n; $i++) {
            $ch = $null
            try { $ch = $range.Characters($i, 1) } catch {}
            if ($null -eq $ch) { continue }
            $entry = [ordered]@{ index = $i; text = $null; left = $null; width = $null }
            try { $entry.text = [string]$ch.Text } catch {}
            try { $entry.left = [double]$ch.BoundLeft } catch {}
            try { $entry.width = [double]$ch.BoundWidth } catch {}
            $s.chars += $entry
        }
    }

    return $s
}

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file

    $record = [ordered]@{
        deck     = $deck.deck
        file     = $deck.file
        opened   = $false
        repaired = $null
        error    = $null
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
                    $record.shapes += Read-Shape $shape
                }
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-14} {1,-9} {2} shape(s)" -f $deck.deck, $state, $record.shapes.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'metric-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
