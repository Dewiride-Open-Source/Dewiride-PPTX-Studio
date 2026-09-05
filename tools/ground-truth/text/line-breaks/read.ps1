# Experiment T3, step 3 - ask PowerPoint where it broke the line.
#
#   powershell -File tools/ground-truth/read-breaks.ps1 -Dir <work-dir>
#
# Reads `break-inputs.json`, opens each probe deck, and records the *text* of
# every laid-out line of every shape.
#
# ## Why this is a stronger measurement than T2's
#
# `TextRange2.Lines($i, 1).Text` returns the characters on line i. T2 had to
# infer a rule from a float; here the answer comes back as the string itself, so
# a break position is read rather than reconstructed. There is no threshold, no
# rounding, and nothing to model in between.
#
# The widths are recorded too - `BoundWidth` per line and the shape's own width -
# because the *fit test* is a separate question from the *opportunity set*, and
# the trailing-space and hangingPunct probes are asking about the fit test.
#
# ## Encoding
#
# The probe strings are Japanese, Thai and Arabic as well as English. The inputs
# are read as UTF-8 explicitly, because a file without a BOM is otherwise
# decoded in the console codepage and every non-ASCII probe id would silently
# stop matching. `ConvertTo-Json` escapes non-ASCII on output, which is what the
# analysis wants: `。` and `｡` are indistinguishable on screen and must
# not be indistinguishable in the fixture.
#
# `Open2007` with `OpenAndRepair:=msoFalse` first, then with repair allowed, so
# REFUSED and REPAIRED stay distinguishable - C2 found that plain `Open` repairs
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
$inputsPath = Join-Path $root 'break-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no break-inputs.json in $root - run build-break-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

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
        id         = $name
        shapeWidth = $null
        width      = $null
        height     = $null
        left       = $null
        top        = $null
        lineCount  = $null
        paraCount  = $null
        fontName   = $null
        fontSize   = $null
        text       = $null
        lines      = @()
    }

    try { $s.shapeWidth = [double]$shape.Width } catch {}

    $range = $null
    try { $range = $shape.TextFrame2.TextRange } catch {}
    if ($null -eq $range) { return $s }

    try { $s.width = [double]$range.BoundWidth } catch {}
    try { $s.height = [double]$range.BoundHeight } catch {}
    try { $s.left = [double]$range.BoundLeft } catch {}
    try { $s.top = [double]$range.BoundTop } catch {}
    try { $s.paraCount = [int]$range.Paragraphs().Count } catch {}
    try { $s.text = [string]$range.Text } catch {}

    # The resolved face. It does not report a substitution - a missing face
    # still reports the name that was asked for - so the analysis detects that
    # from the widths instead. Recorded anyway, because a face reporting
    # something *else* would mean the probe never said what it meant to say.
    try { $s.fontName = [string]$range.Font.Name } catch {}
    try { $s.fontSize = [double]$range.Font.Size } catch {}

    $lineCount = 0
    try { $lineCount = [int]$range.Lines().Count } catch {}
    $s.lineCount = $lineCount
    for ($i = 1; $i -le $lineCount; $i++) {
        $line = $null
        try { $line = $range.Lines($i, 1) } catch {}
        if ($null -eq $line) { continue }
        $entry = [ordered]@{ index = $i; text = $null; left = $null; top = $null; width = $null; length = $null }
        try { $entry.text = [string]$line.Text } catch {}
        try { $entry.left = [double]$line.BoundLeft } catch {}
        try { $entry.top = [double]$line.BoundTop } catch {}
        try { $entry.width = [double]$line.BoundWidth } catch {}
        # `Length` is PowerPoint's own count of characters in the line, which is
        # the answer without having to trust the round trip through `.Text` -
        # two independent readings of the same fact, and the analysis asserts
        # they agree.
        try { $entry.length = [int]$line.Length } catch {}
        $s.lines += $entry
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
    Write-Host ("{0,-16} {1,-9} {2} shape(s)" -f $deck.deck, $state, $record.shapes.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'break-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
