# Experiment C8, step 2 - ask PowerPoint what each probe's table draws, and what it writes back.
#
#   powershell -File tools/ground-truth/model/tables/styles/read.ps1 -Dir <work-dir>
#
# Opens each package (repair refused, then allowed), exports each slide as a BMP - twice unless a
# sweep - saves a copy and reopens it, re-applies every style of the reapply sweeps and saves them,
# then exports the controls again so a style leaked between packages shows.
param([Parameter(Mandatory = $true)][string]$Dir)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3
$ppSaveAsOpenXMLPresentation = 24

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'style-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no style-inputs.json in $root - run tools/ground-truth/model/tables/styles/build-deck.ts first"
}
$inputs = (Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8) | ConvertFrom-Json
foreach ($sub in @('bmp', 'resaved', 'resaved2', 'reapplied')) { New-Item -ItemType Directory -Force -Path (Join-Path $root $sub) | Out-Null }

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
if ($app.Presentations.Count -gt 0) {
    throw "PowerPoint has $($app.Presentations.Count) presentation(s) open; close them first"
}
$app.DisplayAlerts = $ppAlertsNone
$app.AutomationSecurity = $msoAutomationSecurityForceDisable

Add-Type -AssemblyName UIAutomationClient
$pid0 = [int][System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$app.HWND).Current.ProcessId
$session = [ordered]@{
    pid = $pid0; started = (Get-Process -Id $pid0).StartTime.ToString('o')
    version = [string]$app.Version; build = [string]$app.Build; created = $created
}

# FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair
function Open-Deck([string]$file, $record) {
    try {
        $record.opened = $true; $record.repaired = $false
        return $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoFalse)
    }
    catch {
        $record.error = $_.Exception.Message
        try {
            $record.repaired = $true
            return $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoTrue)
        }
        catch { $record.opened = $false; $record.repaired = $null; return $null }
    }
}

function Find-Table($slide) {
    foreach ($shape in $slide.Shapes) {
        $has = $false
        try { $has = [bool]$shape.HasTable } catch {}
        if ($has) { return $shape }
    }
    return $null
}

function Read-Fill($table, [int]$r, [int]$c) {
    $f = $table.Cell($r, $c).Shape.Fill
    return [ordered]@{ r = $r; c = $c; visible = [int]$f.Visible; rgb = [int]$f.ForeColor.RGB; transparency = [double]$f.Transparency }
}

function Export-Slide($slide, [string]$name) {
    $path = Join-Path (Join-Path $root 'bmp') $name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $slide.Export($path, 'BMP', [int]$inputs.export.width, [int]$inputs.export.height)
    return 'bmp/' + $name
}

function Read-Slide($slide, [string]$stem, [bool]$twice) {
    $reading = [ordered]@{ styleId = $null; styleName = $null; rect = $null; fills = @(); error = $null; a = $null; b = $null }
    $shape = Find-Table $slide
    if ($null -eq $shape) { $reading.error = 'no table' }
    else {
        try {
            $reading.styleId = [string]$shape.Table.Style.Id
            $reading.styleName = [string]$shape.Table.Style.Name
            $reading.rect = @([double]$shape.Left, [double]$shape.Top, [double]$shape.Width, [double]$shape.Height)
            $reading.fills = @((Read-Fill $shape.Table 1 1), (Read-Fill $shape.Table 2 2))
        }
        catch { $reading.error = $_.Exception.Message }
    }
    $reading.a = Export-Slide $slide "$stem.a.bmp"
    if ($twice) { $reading.b = Export-Slide $slide "$stem.b.bmp" }
    return $reading
}

$records = @()
foreach ($probe in $inputs.probes) {
    $file = Join-Path $root $probe.file
    $record = [ordered]@{
        id = $probe.id; opened = $false; repaired = $null; error = $null
        slides = @(); resaved = $null; reread = @(); resaved2 = $null
    }
    $pres = Open-Deck $file $record
    if ($record.opened) {
        try {
            for ($k = 1; $k -le $pres.Slides.Count; $k++) {
                $record.slides += , (Read-Slide $pres.Slides.Item($k) ('{0}-{1}' -f $probe.id, ($k - 1)) (-not $probe.sweep))
            }
            $copy = Join-Path $root "resaved\$($probe.id).pptx"
            if (Test-Path -LiteralPath $copy) { Remove-Item -LiteralPath $copy -Force }
            $pres.SaveCopyAs($copy, $ppSaveAsOpenXMLPresentation)
            $record.resaved = "resaved/$($probe.id).pptx"
        }
        finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }

        $again = [ordered]@{ opened = $false; repaired = $null; error = $null }
        $pres = Open-Deck $copy $again
        if ($again.opened) {
            try {
                if (-not $probe.sweep) {
                    for ($k = 1; $k -le $pres.Slides.Count; $k++) {
                        $one = Read-Slide $pres.Slides.Item($k) ('{0}-{1}.resaved' -f $probe.id, ($k - 1)) $false
                        $record.reread += , $one
                    }
                }
                $copy2 = Join-Path $root "resaved2\$($probe.id).pptx"
                if (Test-Path -LiteralPath $copy2) { Remove-Item -LiteralPath $copy2 -Force }
                $pres.SaveCopyAs($copy2, $ppSaveAsOpenXMLPresentation)
                $record.resaved2 = "resaved2/$($probe.id).pptx"
            }
            finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }
        }
        $record.resavedOpen = $again
    }
    $records += , $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ('{0,-28} {1,-9} {2}' -f $probe.id, $state, $record.slides.Count)
}

$reapplied = @()
foreach ($id in $inputs.reapply.probes) {
    $probe = $inputs.probes | Where-Object { $_.id -eq $id }
    $entry = [ordered]@{ id = $id; opened = $false; repaired = $null; error = $null; applied = 0; file = $null }
    $pres = Open-Deck (Join-Path $root $probe.file) $entry
    if ($entry.opened) {
        try {
            foreach ($slide in $pres.Slides) {
                $table = (Find-Table $slide).Table
                $own = [string]$table.Style.Id
                $table.ApplyStyle([string]$inputs.reapply.through, $false)
                $table.ApplyStyle($own, $false)
                $entry.applied += 1
            }
            $copy = Join-Path $root "reapplied\$id.pptx"
            if (Test-Path -LiteralPath $copy) { Remove-Item -LiteralPath $copy -Force }
            $pres.SaveCopyAs($copy, $ppSaveAsOpenXMLPresentation)
            $entry.file = "reapplied/$id.pptx"
        }
        finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }
    }
    $reapplied += , $entry
    Write-Host ('{0,-28} reapplied {1}' -f $id, $entry.applied)
}

$again = @()
foreach ($id in $inputs.reread) {
    $probe = $inputs.probes | Where-Object { $_.id -eq $id }
    $entry = [ordered]@{ id = $id; opened = $false; repaired = $null; error = $null; slide = $null }
    $pres = Open-Deck (Join-Path $root $probe.file) $entry
    if ($entry.opened) {
        try { $entry.slide = Read-Slide $pres.Slides.Item(1) "$id-0.again" $false }
        finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }
    }
    $again += , $entry
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = [ordered]@{ session = $session; probes = $records; again = $again; reapplied = $reapplied } | ConvertTo-Json -Depth 12
[System.IO.File]::WriteAllText((Join-Path $root 'style-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
