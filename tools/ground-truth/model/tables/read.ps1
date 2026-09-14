# Experiment C7, step 2 - ask PowerPoint which cell each grid position belongs to.
#
#   powershell -File tools/ground-truth/model/tables/read.ps1 -Dir <work-dir> [-Only <glob>]
#
# Opens each probe package (repair refused first, then allowed, so REPAIRED and REFUSED differ)
# and records the frame, the column widths, the row heights and every position's rectangle from
# `Table.Cell(r, c).Shape`; a copy goes under `resaved/` for what PowerPoint writes back.

param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Only = ''
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3
$ppSaveAsOpenXMLPresentation = 24

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'table-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no table-inputs.json in $root - run tools/ground-truth/model/tables/build-deck.ts first"
}
$resavedDir = Join-Path $root 'resaved'
New-Item -ItemType Directory -Force -Path $resavedDir | Out-Null
$pngDir = Join-Path $root 'png'
New-Item -ItemType Directory -Force -Path $pngDir | Out-Null

$text = Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

$created = $false
$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
$app.DisplayAlerts = $ppAlertsNone
$app.AutomationSecurity = $msoAutomationSecurityForceDisable

# Every property read is wrapped, so a position PowerPoint refuses to answer for is a
# recorded error rather than a crashed run.
function Read-Table($shape) {
    $t = $shape.Table
    $out = [ordered]@{
        name    = [string]$shape.Name
        left    = [double]$shape.Left
        top     = [double]$shape.Top
        width   = [double]$shape.Width
        height  = [double]$shape.Height
        rows    = [int]$t.Rows.Count
        cols    = [int]$t.Columns.Count
        colWidths  = @()
        rowHeights = @()
        cells   = @()
    }
    for ($c = 1; $c -le $out.cols; $c++) {
        $w = $null
        try { $w = [double]$t.Columns.Item($c).Width } catch {}
        $out.colWidths += $w
    }
    for ($r = 1; $r -le $out.rows; $r++) {
        $h = $null
        try { $h = [double]$t.Rows.Item($r).Height } catch {}
        $out.rowHeights += $h
        $line = @()
        for ($c = 1; $c -le $out.cols; $c++) {
            $cell = [ordered]@{ l = $null; t = $null; w = $null; h = $null; text = $null
                                margins = $null; vAnchor = $null; orientation = $null; error = $null }
            try {
                $sh = $t.Cell($r, $c).Shape
                $cell.l = [double]$sh.Left
                $cell.t = [double]$sh.Top
                $cell.w = [double]$sh.Width
                $cell.h = [double]$sh.Height
                try { $cell.text = [string]$sh.TextFrame.TextRange.Text } catch {}
                try {
                    $tf = $sh.TextFrame2
                    $cell.margins = @([double]$tf.MarginLeft, [double]$tf.MarginTop, [double]$tf.MarginRight, [double]$tf.MarginBottom)
                    $cell.vAnchor = [int]$tf.VerticalAnchor
                    $cell.orientation = [int]$tf.Orientation
                } catch {}
            }
            catch { $cell.error = $_.Exception.Message }
            $line += $cell
        }
        $out.cells += , $line
    }
    return $out
}

$decks = @()
foreach ($deck in $inputs.decks) {
    if ($Only -ne '' -and $deck.deck -notlike $Only) { continue }
    $file = Join-Path $root $deck.file
    $record = [ordered]@{
        deck = $deck.deck; file = $deck.file; hostile = [bool]$deck.hostile
        opened = $false; repaired = $null; error = $null
        table = $null; shapes = @(); resaved = $null; png = $null
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
        catch { $record.opened = $false; $record.repaired = $null }
    }

    if ($record.opened) {
        try {
            $slide = $pres.Slides.Item(1)
            foreach ($shape in $slide.Shapes) {
                $entry = [ordered]@{ name = [string]$shape.Name; hasTable = $false; type = $null }
                try { $entry.type = [int]$shape.Type } catch {}
                try { $entry.hasTable = [bool]$shape.HasTable } catch {}
                $record.shapes += $entry
                if ($entry.hasTable -and $null -eq $record.table) {
                    try { $record.table = Read-Table $shape }
                    catch { $record.table = [ordered]@{ error = $_.Exception.Message } }
                }
            }
            $png = Join-Path $pngDir ($deck.deck + '.png')
            try { $slide.Export($png, 'PNG', 1280, 720); $record.png = 'png/' + $deck.deck + '.png' } catch {}
            $copy = Join-Path $resavedDir ($deck.deck + '.pptx')
            if (Test-Path -LiteralPath $copy) { Remove-Item -LiteralPath $copy -Force }
            try { $pres.SaveCopyAs($copy, $ppSaveAsOpenXMLPresentation); $record.resaved = 'resaved/' + $deck.deck + '.pptx' }
            catch { $record.resaved = $null; $record.error = 'SaveCopyAs: ' + $_.Exception.Message }
        }
        finally {
            try { $pres.Saved = $msoTrue } catch {}
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    $cells = if ($null -ne $record.table -and $null -ne $record.table.rows) { "{0}x{1}" -f $record.table.rows, $record.table.cols } else { '-' }
    Write-Host ("{0,-24} {1,-9} {2}" -f $deck.deck, $state, $cells)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'table-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
