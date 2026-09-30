# Experiment C9, step 0 - ask PowerPoint to AUTHOR a table's borders, fills and flags, and save
# each, so the analysis can read which cells PowerPoint writes them into.
#
#   powershell -File tools/ground-truth/model/tables/cascade/author.ps1 -Dir <out-dir>
#
# Writes `author-<name>.pptx` per operation and `author-cascade-log.json`. Creates and saves into
# <out-dir>, opens nothing else, refuses to start while any deck is open, and never quits a
# PowerPoint it did not start.
param([Parameter(Mandatory = $true)][string]$Dir)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$ppSaveAsOpenXMLPresentation = 24
$ppLayoutBlank = 12
# ppBorderTop 1, Left 2, Bottom 3, Right 4.
$ppBorderTop = 1
$ppBorderLeft = 2
$ppBorderBottom = 3
$ppBorderRight = 4

$root = (Resolve-Path -LiteralPath $Dir).Path
$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
if ($app.Presentations.Count -gt 0) {
    throw "PowerPoint has $($app.Presentations.Count) presentation(s) open; close them first"
}
$app.DisplayAlerts = $ppAlertsNone

# A red 3-pt line, and a green fill: colours no default style draws. RGB is BGR-packed.
$red = 0x0000FF
$green = 0x00FF00

function Set-Border($cell, [int]$side) {
    $b = $cell.Borders.Item($side)
    $b.Visible = $msoTrue
    $b.Weight = 3
    $b.ForeColor.RGB = $red
}

# Each operation gets a fresh 3x3 table in the default style and is saved alone.
$operations = [ordered]@{
    'border-right'  = { param($t) Set-Border $t.Cell(2, 2) $ppBorderRight }
    'border-bottom' = { param($t) Set-Border $t.Cell(2, 2) $ppBorderBottom }
    'border-top'    = { param($t) Set-Border $t.Cell(2, 2) $ppBorderTop }
    'border-left'   = { param($t) Set-Border $t.Cell(2, 2) $ppBorderLeft }
    'border-hidden' = { param($t) $t.Cell(2, 2).Borders.Item($ppBorderBottom).Visible = $msoFalse }
    'fill'          = { param($t) $t.Cell(2, 2).Shape.Fill.ForeColor.RGB = $green }
    'merged-bottom' = { param($t) $t.Cell(1, 2).Merge($t.Cell(2, 2)); Set-Border $t.Cell(1, 2) $ppBorderBottom }
    'merged-right'  = { param($t) $t.Cell(1, 2).Merge($t.Cell(2, 2)); Set-Border $t.Cell(1, 2) $ppBorderRight }
    'merged-wide-bottom' = { param($t) $t.Cell(2, 2).Merge($t.Cell(2, 3)); Set-Border $t.Cell(2, 2) $ppBorderBottom }
    'merged-edge-bottom' = { param($t) $t.Cell(3, 1).Merge($t.Cell(3, 2)); Set-Border $t.Cell(3, 1) $ppBorderBottom }
    'rtl-left'      = { param($t) $t.TableDirection = 2; Set-Border $t.Cell(2, 2) $ppBorderLeft }
    'flags'         = { param($t) $t.FirstRow = $msoFalse; $t.LastRow = $msoTrue; $t.HorizBanding = $msoFalse; $t.FirstCol = $msoTrue }
}

$log = [ordered]@{ version = [string]$app.Version; build = [string]$app.Build; operations = @() }
foreach ($name in $operations.Keys) {
    $pres = $app.Presentations.Add($msoFalse)
    try {
        $slide = $pres.Slides.Add(1, $ppLayoutBlank)
        $shape = $slide.Shapes.AddTable(3, 3, 72, 72, 216, 108)
        & $operations[$name] $shape.Table
        $file = Join-Path $root "author-$name.pptx"
        if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force }
        $pres.SaveAs($file, $ppSaveAsOpenXMLPresentation)
        $log.operations += , [ordered]@{ name = $name; file = "author-$name.pptx"; style = [string]$shape.Table.Style.Id }
        Write-Host ('{0,-18} saved' -f $name)
    }
    finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }
}

if ($created) { try { $app.Quit() } catch {} }
$json = $log | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $root 'author-cascade-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
