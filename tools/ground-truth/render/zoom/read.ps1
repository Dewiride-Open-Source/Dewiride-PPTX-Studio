# Experiment F2, step 2 - ask PowerPoint to draw the deck at six widths.
#
#   powershell -File tools/ground-truth/render/zoom/read.ps1 -Dir <work-dir>
#
# Reads `zoom-inputs.json`, opens the deck and exports every slide at every
# width, twice, as BMP - C4's format, so the analysis needs no decoder.
#
# ## Every export is made twice
#
# PowerPoint does not rasterise a slide identically twice (F1 measured 29 of
# 155), so each slide is exported to two files and the analysis decides on the
# decoded pixels whether a probe's two readings agree before it trusts either.
#
# ## A width PowerPoint refuses is a finding, not a crash
#
# `Slide.Export` is asked for 120 to 3840 pixels wide. An error at any width is
# recorded for that width and the run goes on, so a size ceiling is measured
# rather than guessed around.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason every other read.ps1
# gives: `Open` silently repairs, and a silently repaired deck is not the deck
# whose bytes were built.
#
# Read-only. Nothing is written back to the deck. Attaches to a running
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
$inputsPath = Join-Path $root 'zoom-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no zoom-inputs.json in $root - run tools/ground-truth/render/zoom/build-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$started = $false
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
}
catch {
    $app = New-Object -ComObject PowerPoint.Application
    $started = $true
}
$app.DisplayAlerts = $ppAlertsNone
$app.AutomationSecurity = $msoAutomationSecurityForceDisable

$version = $app.Version
$build = $app.Build

$deckPath = Join-Path $root $inputs.deck
$exports = @()
$refused = @()
$opened = $false
$repaired = $false

try {
    $pres = $null
    try {
        $pres = $app.Presentations.Open2007($deckPath, $msoTrue, $msoFalse, $msoFalse, $msoFalse)
        $opened = $true
        $pageWidth = [double]$pres.PageSetup.SlideWidth
        $pageHeight = [double]$pres.PageSetup.SlideHeight

        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            foreach ($w in $inputs.widths) {
                $width = [int]$w
                # Half up, as `Math.round` does on the other side.
                $height = [int][Math]::Round($width * $pageHeight / $pageWidth, [MidpointRounding]::AwayFromZero)
                foreach ($pass in @('a', 'b')) {
                    $name = ('slide{0}-{1}-{2}.bmp' -f $i, $width, $pass)
                    try {
                        $slide.Export((Join-Path $root $name), 'BMP', $width, $height)
                        $exports += [ordered]@{ slide = $i; width = $width; height = $height; pass = $pass; file = $name }
                    }
                    catch {
                        $refused += [ordered]@{ slide = $i; width = $width; pass = $pass; error = $_.Exception.Message }
                    }
                }
            }
        }
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match 'repair') { $repaired = $true }
        $refused += [ordered]@{ slide = 0; width = 0; pass = ''; error = $message }
    }
    finally {
        if ($null -ne $pres) { try { $pres.Close() } catch {} }
    }
}
finally {
    if ($started) { try { $app.Quit() } catch {} }
}

$out = [ordered]@{
    powerpoint = ('{0} build {1}' -f $version, $build)
    opened     = $opened
    repaired   = $repaired
    exports    = $exports
    refused    = $refused
}
$json = $out | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText((Join-Path $root 'zoom-export.json'), $json, [Text.UTF8Encoding]::new($false))
Write-Output ('exported {0} bitmap(s), {1} refusal(s), to {2}' -f $exports.Count, $refused.Count, $root)
