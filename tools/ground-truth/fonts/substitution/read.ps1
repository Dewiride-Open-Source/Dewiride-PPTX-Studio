# Experiment T7, step 2 - ask PowerPoint which face it drew, and how wide.
#
#   powershell -File tools/ground-truth/fonts/substitution/read.ps1 -Dir <work-dir>
#
# Three instruments: the EMF's LOGFONTW face name, TextRange2.Font.Name as
# PowerPoint resolves it, and BoundWidth - compared only against another probe
# read the same way, so the reader's own bias cancels. ADR 0033.
#
# Also records the installed font families, by name only, from GDI+.
# Opens with OpenAndRepair:=msoFalse first so REFUSED and REPAIRED stay apart.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Only = ''
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'substitution-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no substitution-inputs.json in $root - run build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$emfDir = Join-Path $root 'emf'
New-Item -ItemType Directory -Force -Path $emfDir | Out-Null

# Family names only. Nothing here opens a font file or walks a directory.
Add-Type -AssemblyName System.Drawing
$installed = @()
try {
    $collection = New-Object System.Drawing.Text.InstalledFontCollection
    foreach ($family in $collection.Families) { $installed += $family.Name }
}
catch {
    Write-Host "could not enumerate installed families: $($_.Exception.Message)"
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

# Every read is wrapped: a property that throws is a field the analysis can
# score, a crashed script is not a measurement.
function Read-Shape($shape) {
    $s = [ordered]@{
        name       = ''
        fontName   = $null
        fontNameFE = $null
        fontNameC  = $null
        fontSize   = $null
        boundWidth = $null
        boundHeight = $null
        error      = $null
    }
    try { $s.name = [string]$shape.Name } catch {}
    try {
        $r = $shape.TextFrame2.TextRange
        try { $s.fontName = [string]$r.Font.Name } catch {}
        try { $s.fontNameFE = [string]$r.Font.NameFarEast } catch {}
        try { $s.fontNameC = [string]$r.Font.NameComplexScript } catch {}
        try { $s.fontSize = [double]$r.Font.Size } catch {}
        try { $s.boundWidth = [double]$r.BoundWidth } catch {}
        try { $s.boundHeight = [double]$r.BoundHeight } catch {}
    }
    catch {
        $s.error = $_.Exception.Message
    }
    return $s
}

$decks = @()

foreach ($deck in $inputs.decks) {
    if ($Only -ne '' -and $deck.key -notlike $Only) { continue }
    $file = Join-Path $root $deck.file

    $record = [ordered]@{
        key      = $deck.key
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
                $entry = [ordered]@{ slide = $i; shape = $null; emf = $null; emfError = $null }
                foreach ($shape in $slide.Shapes) {
                    $entry.shape = Read-Shape $shape
                    break
                }
                $emfName = ("{0}-s{1}.emf" -f $deck.key, $i)
                $emfPath = Join-Path $emfDir $emfName
                if (Test-Path -LiteralPath $emfPath) { Remove-Item -LiteralPath $emfPath -Force }
                try {
                    $slide.Export($emfPath, 'EMF', 1920, 1080)
                    $entry.emf = "emf/" + $emfName
                }
                catch {
                    $entry.emfError = $_.Exception.Message
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
    Write-Host ("{0,-20} {1,-9} {2,4} slide(s)" -f $deck.key, $state, $record.slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ installedFamilies = ($installed | Sort-Object); decks = $decks } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'substitution-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("done - {0} installed families" -f $installed.Count)
