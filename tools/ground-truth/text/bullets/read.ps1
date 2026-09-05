# Experiment T5, step 3 - ask PowerPoint what it drew, and where.
#
#   powershell -File tools/ground-truth/text/bullets/read.ps1 -Dir <work-dir>
#
# Reads `bullet-inputs.json`, opens each probe deck read-only, and takes two
# readings of every shape.
#
# ## Two instruments, and they fail differently
#
# **`Slide.Export(path, "EMF")` says what was drawn.** An EMF is a recording of
# GDI calls, so a bullet comes back as the characters PowerPoint asked GDI to
# draw, with the face, size and colour it asked for. That is the whole
# identification problem solved by reading rather than by fitting: T5 never has
# to score a candidate string against a measured width to find out that
# `romanLcParenBoth` at 4 renders `(iv)`.
#
# PowerPoint has no SVG converter installed and refuses that filter by name;
# EMF and WMF both work, and EMF is the one with 32-bit coordinates and
# `ExtTextOutW` rather than the ANSI-only 16-bit original.
#
# **`Paragraphs(i).BoundLeft` says where it landed.** The EMF cannot answer
# that: its positions are in device units under a world transform, and a second
# exact instrument already exists. With `marL="0" indent="0"` and zero insets,
# the paragraph's left edge minus the shape's own left edge *is* the bullet's
# advance width, in points, to a quarter of one.
#
# Characters are read only where a probe asks for it. `Characters(i)` is one COM
# round trip per character and the script-run family is the only one that needs
# them; reading them everywhere would multiply the run time by twenty for
# nothing.
#
# `Open2007` with `OpenAndRepair:=msoFalse` first, then with repair allowed, so
# REFUSED and REPAIRED stay distinguishable - C2 found that plain `Open` repairs
# silently and reports success on exactly the files being asked about.
#
# Read-only. Nothing is written back to any deck. Attaches to a running
# PowerPoint if there is one and never quits one it did not start.

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
$inputsPath = Join-Path $root 'bullet-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no bullet-inputs.json in $root - run tools/ground-truth/text/bullets/build-deck.ts first"
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
function Read-Shape($shape, $wantChars) {
    $s = [ordered]@{
        id         = ''
        left       = $null
        top        = $null
        width      = $null
        height     = $null
        textLeft   = $null
        text       = $null
        paragraphs = @()
        lines      = @()
        characters = @()
    }
    try { $s.id = [string]$shape.Name } catch {}
    try { $s.left = [double]$shape.Left } catch {}
    try { $s.top = [double]$shape.Top } catch {}
    try { $s.width = [double]$shape.Width } catch {}
    try { $s.height = [double]$shape.Height } catch {}

    $range = $null
    try { $range = $shape.TextFrame2.TextRange } catch {}
    if ($null -eq $range) { return $s }

    try { $s.textLeft = [double]$range.BoundLeft } catch {}
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

    if ($wantChars) {
        $len = 0
        try { $len = [int]$range.Length } catch {}
        for ($i = 1; $i -le $len; $i++) {
            $c = $null
            try { $c = $range.Characters($i, 1) } catch {}
            if ($null -eq $c) { continue }
            $entry = [ordered]@{ index = $i; text = $null; left = $null; width = $null }
            try { $entry.text = [string]$c.Text } catch {}
            try { $entry.left = [double]$c.BoundLeft } catch {}
            try { $entry.width = [double]$c.BoundWidth } catch {}
            $s.characters += $entry
        }
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
        # A field is computed when PowerPoint lays the slide out, which is
        # somewhere inside this deck's window. The seconds move during a
        # two-minute run, so a date field cannot be turned back into a format
        # pattern without knowing when it was read.
        openedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        closedAt = $null
        opened   = $false
        repaired = $null
        error    = $null
        emf      = @()
        shapes   = @()
    }

    # Which shapes want per-character readings, by name.
    $wantChars = @{}
    foreach ($s in $deck.shapes) { if ($s.readChars) { $wantChars[[string]$s.id] = $true } }

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
                    $name = ''
                    try { $name = [string]$shape.Name } catch {}
                    $entry = Read-Shape $shape ($wantChars.ContainsKey($name))
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

    $record.closedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-24} {1,-9} {2,4} shape(s) {3,3} emf" -f $deck.deck, $state, $record.shapes.Count, $record.emf.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'bullet-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
