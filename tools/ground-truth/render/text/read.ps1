# Experiment T8, step 2 - ask PowerPoint where it drew the glyphs.
#
#   powershell -File tools/ground-truth/render/text/read.ps1 -Dir <work-dir>
#
# Three instruments, and each answers something the others cannot. The object
# model gives every line's box in slide points. The EMF gives the world
# transform, the baseline origin and the per-character advances GDI was handed -
# which is the only place a flip or a rotation is visible as a number. The
# bitmap gives ink, for the questions that are about a rule nobody records.
#
# Opens with OpenAndRepair:=msoFalse first so REFUSED and REPAIRED stay apart.
# Read-only; never quits a PowerPoint it did not start.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Only = '',
    [switch]$NoBitmap
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'text-render-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no text-render-inputs.json in $root - run build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$wantsBitmap = @{}
foreach ($p in $inputs.probes) { $wantsBitmap[$p.id] = [bool]$p.bitmap }

$emfDir = Join-Path $root 'emf'
$bmpDir = Join-Path $root 'bmp'
New-Item -ItemType Directory -Force -Path $emfDir | Out-Null
New-Item -ItemType Directory -Force -Path $bmpDir | Out-Null

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
function Read-Bounds($range) {
    $b = [ordered]@{ left = $null; top = $null; width = $null; height = $null; text = $null }
    try { $b.left = [double]$range.BoundLeft } catch {}
    try { $b.top = [double]$range.BoundTop } catch {}
    try { $b.width = [double]$range.BoundWidth } catch {}
    try { $b.height = [double]$range.BoundHeight } catch {}
    try { $b.text = [string]$range.Text } catch {}
    return $b
}

function Read-Font($range) {
    $f = [ordered]@{
        name           = $null
        size           = $null
        bold           = $null
        italic         = $null
        underlineStyle = $null
        strike         = $null
        baselineOffset = $null
        allCaps        = $null
        smallCaps      = $null
        spacing        = $null
        kerning        = $null
    }
    $font = $null
    try { $font = $range.Font } catch {}
    if ($null -eq $font) { return $f }
    try { $f.name = [string]$font.Name } catch {}
    try { $f.size = [double]$font.Size } catch {}
    try { $f.bold = [int]$font.Bold } catch {}
    try { $f.italic = [int]$font.Italic } catch {}
    try { $f.underlineStyle = [int]$font.UnderlineStyle } catch {}
    try { $f.strike = [int]$font.Strike } catch {}
    try { $f.baselineOffset = [double]$font.BaselineOffset } catch {}
    try { $f.allCaps = [int]$font.Allcaps } catch {}
    try { $f.smallCaps = [int]$font.Smallcaps } catch {}
    try { $f.spacing = [double]$font.Spacing } catch {}
    try { $f.kerning = [double]$font.Kerning } catch {}
    return $f
}

function Read-Shape($shape) {
    $s = [ordered]@{
        name        = ''
        slide       = 0
        left        = $null
        top         = $null
        width       = $null
        height      = $null
        rotation    = $null
        orientation = $null
        vAnchor     = $null
        textLeft    = $null
        textTop     = $null
        textWidth   = $null
        textHeight  = $null
        text        = $null
        font        = $null
        paragraphs  = @()
        lines       = @()
        runs        = @()
    }
    try { $s.name = [string]$shape.Name } catch {}
    try { $s.left = [double]$shape.Left } catch {}
    try { $s.top = [double]$shape.Top } catch {}
    try { $s.width = [double]$shape.Width } catch {}
    try { $s.height = [double]$shape.Height } catch {}
    try { $s.rotation = [double]$shape.Rotation } catch {}

    $frame = $null
    try { $frame = $shape.TextFrame2 } catch {}
    if ($null -eq $frame) { return $s }
    try { $s.orientation = [int]$frame.Orientation } catch {}
    try { $s.vAnchor = [int]$frame.VerticalAnchor } catch {}

    $range = $null
    try { $range = $frame.TextRange } catch {}
    if ($null -eq $range) { return $s }

    $whole = Read-Bounds $range
    $s.textLeft = $whole.left
    $s.textTop = $whole.top
    $s.textWidth = $whole.width
    $s.textHeight = $whole.height
    $s.text = $whole.text
    $s.font = Read-Font $range

    $n = 0
    try { $n = [int]$range.Paragraphs().Count } catch {}
    for ($i = 1; $i -le $n; $i++) {
        $p = $null
        try { $p = $range.Paragraphs($i, 1) } catch {}
        if ($null -eq $p) { continue }
        $entry = Read-Bounds $p
        $entry.Insert(0, 'index', $i)
        try { $entry['align'] = [int]$p.ParagraphFormat.Alignment } catch { $entry['align'] = $null }
        $s.paragraphs += $entry
    }

    $n = 0
    try { $n = [int]$range.Lines().Count } catch {}
    for ($i = 1; $i -le $n; $i++) {
        $line = $null
        try { $line = $range.Lines($i, 1) } catch {}
        if ($null -eq $line) { continue }
        $entry = Read-Bounds $line
        $entry.Insert(0, 'index', $i)
        $s.lines += $entry
    }

    # A PowerPoint "run" is a maximal span of uniform formatting, which is not
    # necessarily an `a:r`. Q7 is exactly the difference between the two.
    $n = 0
    try { $n = [int]$range.Runs().Count } catch {}
    for ($i = 1; $i -le $n; $i++) {
        $run = $null
        try { $run = $range.Runs($i, 1) } catch {}
        if ($null -eq $run) { continue }
        $entry = Read-Bounds $run
        $entry.Insert(0, 'index', $i)
        $entry['font'] = Read-Font $run
        $s.runs += $entry
    }

    return $s
}

# -Only re-reads one package. The rest are carried over from the previous run
# rather than dropped, so a narrow re-read cannot silently shrink the readings.
$decks = @()
$readingsPath = Join-Path $root 'text-render-readings.json'
$previous = @{}
if ($Only -ne '' -and (Test-Path -LiteralPath $readingsPath)) {
    $old = (Get-Content -LiteralPath $readingsPath -Raw -Encoding UTF8) | ConvertFrom-Json
    foreach ($d in $old.decks) { $previous[[string]$d.deck] = $d }
}

foreach ($deck in $inputs.decks) {
    if ($Only -ne '' -and $deck.deck -notlike $Only) {
        if ($previous.ContainsKey([string]$deck.deck)) { $decks += $previous[[string]$deck.deck] }
        continue
    }
    $file = Join-Path $root $deck.file

    $byIndex = @{}
    foreach ($s in $deck.slides) { $byIndex[[int]$s.slide] = [string]$s.probe }

    $record = [ordered]@{
        deck     = $deck.deck
        question = $deck.question
        file     = $deck.file
        opened   = $false
        repaired = $null
        error    = $null
        slides   = @()
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
                $probe = $byIndex[$i]
                $entry = [ordered]@{
                    slide  = $i
                    probe  = $probe
                    shapes = @()
                    emf    = $null
                    bmp    = $null
                    error  = $null
                }
                foreach ($shape in $slide.Shapes) {
                    $read = Read-Shape $shape
                    $read.slide = $i
                    $entry.shapes += $read
                }

                $emfName = ("{0}-{1}.emf" -f $deck.deck, $probe)
                $emfPath = Join-Path $emfDir $emfName
                if (Test-Path -LiteralPath $emfPath) { Remove-Item -LiteralPath $emfPath -Force }
                try {
                    $slide.Export($emfPath, 'EMF', 1920, 1080)
                    $entry.emf = "emf/" + $emfName
                }
                catch { $entry.error = $_.Exception.Message }

                if (-not $NoBitmap -and $wantsBitmap[$probe]) {
                    $bmpName = ("{0}-{1}.bmp" -f $deck.deck, $probe)
                    $bmpPath = Join-Path $bmpDir $bmpName
                    if (Test-Path -LiteralPath $bmpPath) { Remove-Item -LiteralPath $bmpPath -Force }
                    try {
                        $slide.Export($bmpPath, 'BMP', 1920, 1080)
                        $entry.bmp = "bmp/" + $bmpName
                    }
                    catch { $entry.error = $_.Exception.Message }
                }

                $record.slides += $entry
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-12} {1,-9} {2,4} slide(s)" -f $deck.deck, $state, $record.slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText($readingsPath, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
