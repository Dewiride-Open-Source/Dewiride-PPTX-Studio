# Experiment C8, step 0 - ask PowerPoint which table styles it builds in, by driving its own
# Table Styles gallery, then have it write every one of them into decks.
#
#   powershell -File tools/ground-truth/model/tables/styles/author.ps1 -Dir <out-dir> -Discover
#   powershell -File tools/ground-truth/model/tables/styles/author.ps1 -Dir <out-dir>
#
# -Discover dumps the gallery's UI Automation subtree to uia-tree.json, expanding it once. The full
# run invokes every gallery item on a selected table and reads the style back over COM, then applies
# each GUID it found and saves the decks. Refuses to start while any presentation is open, touches
# only PowerPoint's own window, and never quits a PowerPoint it did not start.
param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [switch]$Discover
)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1
$ppWindowMaximized = 3

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

$root = (Resolve-Path -LiteralPath $Dir).Path

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
if ($app.Presentations.Count -gt 0) {
    throw "PowerPoint has $($app.Presentations.Count) presentation(s) open; close them first"
}
$uiLanguage = [int]$app.LanguageSettings.LanguageID(2)
if ($uiLanguage -ne 1033) { throw "PowerPoint's UI language is $uiLanguage; the gallery's names are English (1033)" }
$app.DisplayAlerts = $ppAlertsNone
$app.Visible = $msoTrue

# PowerPoint answers RPC_E_CALL_REJECTED while the ribbon is busy; only that is retried.
function Invoke-Com([scriptblock]$call) {
    for ($attempt = 0; ; $attempt++) {
        try { return & $call }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -ge 50) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

function Write-Json($value, [string]$name) {
    $json = $value | ConvertTo-Json -Depth 40
    [System.IO.File]::WriteAllText((Join-Path $root $name), $json, (New-Object System.Text.UTF8Encoding($false)))
}

$pid0 = [int]$AE::FromHandle([IntPtr]$app.HWND).Current.ProcessId
$session = [ordered]@{
    pid = $pid0; started = (Get-Process -Id $pid0).StartTime.ToString('o')
    version = [string]$app.Version; build = [string]$app.Build; uiLanguage = $uiLanguage; created = $created
}

function Get-Window { $AE::FromHandle([IntPtr]$app.HWND) }

function Find-First($scope, $property, $value) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($property, $value)
    return $scope.FindFirst($TS::Descendants, $cond)
}

function Describe($el) {
    $c = $el.Current
    return [ordered]@{
        name = [string]$c.Name; type = [string]$c.ControlType.ProgrammaticName
        className = [string]$c.ClassName; automationId = [string]$c.AutomationId
        patterns = @($el.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName })
    }
}

function Dump($el, [int]$depth) {
    $node = Describe $el
    if ($depth -gt 0) {
        $kids = @()
        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $child = $walker.GetFirstChild($el)
        while ($null -ne $child) { $kids += , (Dump $child ($depth - 1)); $child = $walker.GetNextSibling($child) }
        $node.children = $kids
    }
    return $node
}

function Open-TableDesign {
    $cond = New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::TabItem)),
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'Table Design')))
    $tab = (Get-Window).FindFirst($TS::Descendants, $cond)
    if ($null -eq $tab) { throw 'no "Table Design" tab: is the table selected?' }
    $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Milliseconds 300
}

function Find-Gallery {
    $g = Find-First (Get-Window) $AE::AutomationIdProperty 'TableStylesGallery'
    if ($null -eq $g) { Open-TableDesign; $g = Find-First (Get-Window) $AE::AutomationIdProperty 'TableStylesGallery' }
    if ($null -eq $g) { throw 'no TableStylesGallery on the ribbon' }
    return $g
}

# The in-ribbon row of the gallery: every style as a ListItem, in gallery order.
function Get-GalleryItems {
    $row = Find-First (Find-Gallery) $AE::ClassNameProperty 'NetUIGalleryButtonGroup'
    if ($null -eq $row) { throw 'the gallery has no NetUIGalleryButtonGroup' }
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::ListItem)
    return @($row.FindAll($TS::Children, $cond))
}

# The expanded gallery groups its items by category; read once, then closed.
function Get-Categories {
    $gallery = Find-Gallery
    $pattern = $gallery.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
    $pattern.Expand()
    Start-Sleep -Milliseconds 1000
    $groups = @($gallery.FindAll($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'NetUIGalleryCategoryContainer'))))
    $out = @()
    foreach ($group in $groups) {
        $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::ListItem)
        $names = @($group.FindAll($TS::Children, $cond) | ForEach-Object { [string]$_.Current.Name })
        $out += , [ordered]@{ category = [string]$group.Current.Name; items = $names }
    }
    $pattern.Collapse()
    Start-Sleep -Milliseconds 300
    return $out
}

$pres = $app.Presentations.Add($msoTrue)
$slide = $pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(7))
$shape = $slide.Shapes.AddTable(4, 4, 72, 72, 432, 144)
$pres.Windows.Item(1).Activate()
$app.WindowState = $ppWindowMaximized
$shape.Select()
$table = $shape.Table
$session.newTableStyle = [ordered]@{ id = [string]$table.Style.Id; name = [string]$table.Style.Name }
$session.master = [string]$pres.SlideMaster.Name
$session.window = [ordered]@{ width = [double]$app.Width; height = [double]$app.Height }

if ($Discover) {
    Open-TableDesign
    $gallery = Find-Gallery
    $tree = [ordered]@{ session = $session; gallery = (Dump $gallery 4) }
    $pattern = $gallery.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
    $pattern.Expand()
    Start-Sleep -Milliseconds 1000
    $tree.expanded = Dump $gallery 5
    $pattern.Collapse()
    Write-Json $tree 'uia-tree.json'
    $pres.Saved = $msoTrue
    $pres.Close()
    if ($created) { try { $app.Quit() } catch {} }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    Write-Host 'discovered'
    exit 0
}

# ------------------------------------------------------------------ the gallery, item by item ---
Open-TableDesign
$names = @(Get-GalleryItems | ForEach-Object { [string]$_.Current.Name })
Write-Host "the gallery lists $($names.Count) items"
$categories = Get-Categories
$default = [string]$table.Style.Id
$items = @()
for ($i = 0; $i -lt $names.Count; $i++) {
    $name = $names[$i]
    $sentinel = $default
    if ($name -eq $session.newTableStyle.name) { $sentinel = [string]$items[0].styleId }
    Invoke-Com { $table.ApplyStyle($sentinel, $false) } | Out-Null
    if ([string]$table.Style.Id -ne $sentinel) { throw "the sentinel $sentinel did not apply before item $i" }
    Invoke-Com { $shape.Select() } | Out-Null
    $item = @(Get-GalleryItems)[$i]
    if ([string]$item.Current.Name -ne $name) { throw "gallery item $i is now '$($item.Current.Name)', was '$name'" }
    $scroll = $null
    if ($item.TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern, [ref]$scroll)) { $scroll.ScrollIntoView() }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $item.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    $id = $sentinel
    while ($clock.ElapsedMilliseconds -lt 3000) {
        $id = [string](Invoke-Com { $table.Style.Id })
        if ($id -ne $sentinel) { break }
        Start-Sleep -Milliseconds 100
    }
    $items += , [ordered]@{
        seq = $i; uiaName = $name; method = 'invoke'; sentinel = $sentinel
        styleId = $id; comName = [string](Invoke-Com { $table.Style.Name }); ms = [int]$clock.ElapsedMilliseconds
    }
    Write-Host ('{0,2} {1,-36} {2}' -f $i, $name, $id)
}
$pres.Saved = $msoTrue
$pres.Close()

# ------------------------------------------------------------------ every style, by GUID ---
$ids = @($items | ForEach-Object { $_.styleId } | Select-Object -Unique)
$byName = @{}
foreach ($it in $items) { $byName[$it.uiaName] = $it.styleId }

function Set-Flags($t, [bool]$on) {
    $t.FirstRow = $on; $t.LastRow = $on; $t.FirstCol = $on; $t.LastCol = $on
    $t.HorizBanding = $on; $t.VertBanding = $on
}

function Read-Cell($t, [int]$r, [int]$c) {
    $f = $t.Cell($r, $c).Shape.Fill
    return [ordered]@{ r = $r; c = $c; visible = [int]$f.Visible; type = [int]$f.Type; rgb = [int]$f.ForeColor.RGB; transparency = [double]$f.Transparency }
}

function Read-Borders($t) {
    $out = @()
    foreach ($edge in 1..4) {
        $b = $t.Cell(2, 2).Borders.Item($edge)
        $out += , [ordered]@{ edge = $edge; visible = [int]$b.Visible; rgb = [int]$b.ForeColor.RGB; weight = [double]$b.Weight }
    }
    return $out
}

# One table per slide, each styled by GUID; flags all on, all off, or left as AddTable set them.
function Save-StyleDeck([string]$file, [string[]]$styleIds, [string]$flags, [string]$bmpDir) {
    $p = $app.Presentations.Add($msoFalse)
    $layout = $p.SlideMaster.CustomLayouts.Item(7)
    $readings = @()
    for ($k = 0; $k -lt $styleIds.Count; $k++) {
        $s = $p.Slides.AddSlide($k + 1, $layout)
        $t = $s.Shapes.AddTable(4, 4, 72, 72, 432, 144).Table
        $t.ApplyStyle($styleIds[$k], $false)
        if ($flags -eq 'on') { Set-Flags $t $true } elseif ($flags -eq 'off') { Set-Flags $t $false }
        $reading = [ordered]@{ slide = $k + 1; applied = $styleIds[$k]; id = [string]$t.Style.Id; name = [string]$t.Style.Name }
        if ($bmpDir -ne '') {
            $reading.cells = @((Read-Cell $t 1 1), (Read-Cell $t 2 2), (Read-Cell $t 3 2), (Read-Cell $t 4 4))
            $reading.borders = Read-Borders $t
            $image = Join-Path $bmpDir ('{0}-{1:D2}.bmp' -f [IO.Path]::GetFileNameWithoutExtension($file), ($k + 1))
            $s.Export($image, 'BMP', 960, 540)
        }
        $readings += , $reading
    }
    $p.SaveAs((Join-Path $root $file), $ppSaveAsOpenXMLPresentation)
    $p.Close()
    return $readings
}

$bmp = Join-Path $root 'author-bmp'
New-Item -ItemType Directory -Force $bmp | Out-Null
$decks = [ordered]@{}
$decks['pp-styles-off.pptx'] = Save-StyleDeck 'pp-styles-off.pptx' $ids 'off' $bmp
$decks['pp-styles-on.pptx'] = Save-StyleDeck 'pp-styles-on.pptx' $ids 'on' $bmp
for ($k = 0; $k -lt $ids.Count; $k++) {
    $file = 'pp-one-{0:D2}.pptx' -f ($k + 1)
    $decks[$file] = Save-StyleDeck $file @($ids[$k]) 'default' ''
}
Write-Host "saved $($decks.Count) decks"

# ------------------------------------------------------------------ what PowerPoint writes ---
$dark = $byName['Dark Style 1']
$light = $byName['Light Style 1']
$answers = [ordered]@{}

$p = $app.Presentations.Add($msoFalse)
[void]$p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(7)).Shapes.AddTable(4, 4, 72, 72, 432, 144)
$p.SaveAs((Join-Path $root 'pp-stock.pptx'), $ppSaveAsOpenXMLPresentation); $p.Close()
$answers.stock = 'a new table left as AddTable styled it'

$p = $app.Presentations.Add($msoFalse)
$t = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(7)).Shapes.AddTable(4, 4, 72, 72, 432, 144).Table
$t.ApplyStyle($dark, $false); $t.ApplyStyle($light, $false)
$p.SaveAs((Join-Path $root 'pp-restyle.pptx'), $ppSaveAsOpenXMLPresentation); $p.Close()
$answers.restyle = [ordered]@{ first = $dark; then = $light }

$p = $app.Presentations.Add($msoFalse)
$s1 = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(7))
$gone = $s1.Shapes.AddTable(4, 4, 72, 72, 432, 144)
$gone.Table.ApplyStyle($dark, $false)
$gone.Delete()
[void]$p.Slides.AddSlide(2, $p.SlideMaster.CustomLayouts.Item(7)).Shapes.AddTable(4, 4, 72, 72, 432, 144)
$p.SaveAs((Join-Path $root 'pp-deleted.pptx'), $ppSaveAsOpenXMLPresentation); $p.Close()
$answers.deleted = [ordered]@{ deleted = $dark }

$p = $app.Presentations.Add($msoFalse)
$s1 = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(7))
$s1.Shapes.AddTable(4, 4, 72, 72, 432, 144).Table.ApplyStyle($light, $false)
$s1.Shapes.AddTable(4, 4, 72, 300, 432, 144).Table.ApplyStyle($light, $false)
$p.SaveAs((Join-Path $root 'pp-two-same.pptx'), $ppSaveAsOpenXMLPresentation); $p.Close()
$answers.twoSame = [ordered]@{ style = $light }

# What the object model accepts is a fact about the API, not about the file format.
$api = @()
$p = $app.Presentations.Add($msoFalse)
$t = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(7)).Shapes.AddTable(4, 4, 72, 72, 432, 144).Table
foreach ($form in @($dark.ToLowerInvariant(), $dark.Trim('{', '}'), " $dark ", '{00000000-0000-0000-0000-00000000C700}')) {
    $t.ApplyStyle($light, $false)
    $entry = [ordered]@{ applied = $form; ok = $true; error = $null; id = $null }
    try { $t.ApplyStyle($form, $false) } catch { $entry.ok = $false; $entry.error = $_.Exception.Message }
    $entry.id = [string]$t.Style.Id
    $api += , $entry
}
$p.Saved = $msoTrue; $p.Close()

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

Write-Json ([ordered]@{
        session = $session; gallery = $items; categories = $categories
        decks = $decks; answers = $answers; api = $api
    }) 'author-styles-log.json'
Write-Host 'done'
