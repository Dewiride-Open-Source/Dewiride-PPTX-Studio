# Experiment C7, step 0 - ask PowerPoint to AUTHOR merges, splits and row growth, then read
# what it wrote.
#
#   powershell -File tools/ground-truth/model/tables/author.ps1 -Dir <out-dir>
#
# Writes `pp-merges.pptx` and `pp-sizes.pptx` plus `author-tables-log.json`: per slide, the
# operation and the grid the object model reported after it. Creates and saves into <out-dir>,
# opens nothing else, and never quits a PowerPoint it did not start.
param([Parameter(Mandatory = $true)][string]$Dir)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1

$root = (Resolve-Path -LiteralPath $Dir).Path

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
$app.DisplayAlerts = $ppAlertsNone

function Read-Grid($shape) {
    $t = $shape.Table
    $out = [ordered]@{
        left = [double]$shape.Left; top = [double]$shape.Top
        width = [double]$shape.Width; height = [double]$shape.Height
        rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count
        colWidths = @(); rowHeights = @(); cells = @()
    }
    for ($c = 1; $c -le $out.cols; $c++) { $out.colWidths += [double]$t.Columns.Item($c).Width }
    for ($r = 1; $r -le $out.rows; $r++) {
        $out.rowHeights += [double]$t.Rows.Item($r).Height
        $line = @()
        for ($c = 1; $c -le $out.cols; $c++) {
            $sh = $t.Cell($r, $c).Shape
            $text = ''
            try { $text = [string]$sh.TextFrame.TextRange.Text } catch {}
            $line += [ordered]@{ l = [double]$sh.Left; t = [double]$sh.Top
                                 w = [double]$sh.Width; h = [double]$sh.Height; text = $text }
        }
        $out.cells += , $line
    }
    return $out
}

# A table of 108 pt columns and 36 pt rows at (72, 72), every cell carrying its marker.
function Add-Table($slide, $rows, $cols) {
    $shape = $slide.Shapes.AddTable($rows, $cols, 72, 72, 108 * $cols, 36 * $rows)
    $t = $shape.Table
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $t.Cell($r, $c).Shape.TextFrame.TextRange.Text = ('r' + $r + 'c' + $c)
            $t.Cell($r, $c).Shape.TextFrame.TextRange.Font.Size = 12
        }
    }
    return $shape
}

$log = @()

# ------------------------------------------------------------------ merges and splits ---
$pres = $app.Presentations.Add($msoTrue)
$blank = $pres.SlideMaster.CustomLayouts.Item(7)
$ops = @(
    @{ name = 'merge-h';         op = { param($t) $t.Cell(1, 1).Merge($t.Cell(1, 2)) } },
    @{ name = 'merge-v';         op = { param($t) $t.Cell(1, 1).Merge($t.Cell(2, 1)) } },
    @{ name = 'merge-block';     op = { param($t) $t.Cell(2, 2).Merge($t.Cell(3, 3)) } },
    @{ name = 'merge-block-23';  op = { param($t) $t.Cell(1, 2).Merge($t.Cell(2, 4)) } },
    @{ name = 'merge-corner';    op = { param($t) $t.Cell(3, 3).Merge($t.Cell(4, 4)) } },
    @{ name = 'merge-row';       op = { param($t) $t.Cell(2, 1).Merge($t.Cell(2, 4)) } },
    @{ name = 'merge-col';       op = { param($t) $t.Cell(1, 3).Merge($t.Cell(4, 3)) } },
    @{ name = 'merge-all';       op = { param($t) $t.Cell(1, 1).Merge($t.Cell(4, 4)) } },
    @{ name = 'merge-then-split'; op = { param($t) $t.Cell(1, 1).Merge($t.Cell(1, 2)); $t.Cell(1, 1).Split(1, 2) } },
    @{ name = 'split-cols';      op = { param($t) $t.Cell(2, 2).Split(1, 2) } },
    @{ name = 'split-rows';      op = { param($t) $t.Cell(2, 2).Split(2, 1) } },
    @{ name = 'split-both';      op = { param($t) $t.Cell(2, 2).Split(2, 2) } }
)
$i = 0
foreach ($entry in $ops) {
    $i += 1
    $slide = $pres.Slides.AddSlide($i, $blank)
    $slide.Name = $entry.name
    $shape = Add-Table $slide 4 4
    $shape.Name = $entry.name
    & $entry.op $shape.Table
    $log += [ordered]@{ deck = 'pp-merges.pptx'; slide = $i; name = $entry.name; grid = (Read-Grid $shape) }
    Write-Host ("{0,-18} {1}x{2}" -f $entry.name, $shape.Table.Rows.Count, $shape.Table.Columns.Count)
}
$pres.SaveAs((Join-Path $root 'pp-merges.pptx'), $ppSaveAsOpenXMLPresentation)
$pres.Close()

# ------------------------------------------------------------------ rows and columns ---
$pres = $app.Presentations.Add($msoTrue)
$blank = $pres.SlideMaster.CustomLayouts.Item(7)
$sizes = @(
    @{ name = 'grow'; op = { param($t, $s) $t.Cell(2, 2).Shape.TextFrame.TextRange.Text = "one`rtwo`rthree" } },
    @{ name = 'shrink'; op = { param($t, $s) $t.Rows.Item(2).Height = 10 } },
    @{ name = 'row-taller'; op = { param($t, $s) $t.Rows.Item(2).Height = 100 } },
    @{ name = 'col-wider'; op = { param($t, $s) $t.Columns.Item(1).Width = 200 } },
    @{ name = 'col-narrower'; op = { param($t, $s) $t.Columns.Item(2).Width = 1 } },
    @{ name = 'shape-wider'; op = { param($t, $s) $s.Width = 600 } },
    @{ name = 'shape-narrower'; op = { param($t, $s) $s.Width = 200 } },
    @{ name = 'shape-taller'; op = { param($t, $s) $s.Height = 300 } },
    @{ name = 'shape-shorter'; op = { param($t, $s) $s.Height = 50 } }
)
$i = 0
foreach ($entry in $sizes) {
    $i += 1
    $slide = $pres.Slides.AddSlide($i, $blank)
    $slide.Name = $entry.name
    $shape = Add-Table $slide 3 3
    $shape.Name = $entry.name
    $before = Read-Grid $shape
    & $entry.op $shape.Table $shape
    $log += [ordered]@{ deck = 'pp-sizes.pptx'; slide = $i; name = $entry.name; before = $before; grid = (Read-Grid $shape) }
    Write-Host ("{0,-18} {1}x{2} -> {3}x{4}" -f $entry.name, $before.width, $before.height, $shape.Width, $shape.Height)
}
$pres.SaveAs((Join-Path $root 'pp-sizes.pptx'), $ppSaveAsOpenXMLPresentation)
$pres.Close()

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ files = @('pp-merges.pptx', 'pp-sizes.pptx'); log = $log } | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'author-tables-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
