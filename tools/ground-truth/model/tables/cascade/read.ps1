# Experiment C9, step 2 - ask PowerPoint what every cell of every probe table reports, and export it.
#
#   powershell -File tools/ground-truth/model/tables/cascade/read.ps1 -Dir <work-dir>
#   powershell -File tools/ground-truth/model/tables/cascade/read.ps1 -Dir <work-dir> -Tag <tag> -Decks <id>,<id>
#
# One readings file per deck under readings/, so an interrupted run resumes where it stopped; each
# session reads the control deck first and again last, as control-<tag> in a session of named decks.
# Every slide is exported as PNG at 4 px per point, and each `twins` deck's first slide as BMP too;
# a deck marked `resave` is also saved by PowerPoint under resaved/, to read back what it keeps.
param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [string]$Tag = '',
    [string]$Decks = ''
)

$ErrorActionPreference = 'Stop'
$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

$root = (Resolve-Path -LiteralPath $Dir).Path
$inputsPath = Join-Path $root 'cascade-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no cascade-inputs.json in $root - run tools/ground-truth/model/tables/cascade/build-deck.ts first"
}
$inputs = (Get-Content -LiteralPath $inputsPath -Raw -Encoding UTF8) | ConvertFrom-Json
foreach ($sub in @('readings', 'png', 'twin', 'resaved')) { New-Item -ItemType Directory -Force -Path (Join-Path $root $sub) | Out-Null }

# COM costs about 1.5 ms a property from PowerShell; compiled late binding saves a quarter of that.
Add-Type -ReferencedAssemblies Microsoft.CSharp, System.Core -TypeDefinition @'
using System;
using System.Globalization;
using System.Text;
public static class CascadeReader {
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    // RGB is written raw when it is no colour (a mixed value), so the analysis can refuse it.
    static string Hex(int bgr) {
        if (bgr < 0 || bgr > 0xFFFFFF) return "raw" + bgr.ToString(Inv);
        return string.Format("{0:X2}{1:X2}{2:X2}", bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
    }
    static string Num(double v) { return v.ToString("R", Inv); }
    // A Borders object answers every Item(k) with the side it was first asked for: one per side.
    static dynamic Side(dynamic cell, int k) {
        dynamic borders = cell.Borders;
        return borders.Item(k);
    }
    static void Line(StringBuilder sb, string key, dynamic b) {
        sb.Append(key).Append('=');
        if ((int)b.Visible == 0) { sb.Append("none;"); return; }
        sb.Append(Num((double)(float)b.Weight)).Append('/').Append(Hex((int)b.ForeColor.RGB)).Append('/')
          .Append(Num((double)(float)b.Transparency)).Append('/').Append((int)b.Style).Append('/').Append((int)b.DashStyle).Append(';');
    }
    static void Fill(StringBuilder sb, string key, dynamic fill) {
        sb.Append(key).Append('=');
        if ((int)fill.Visible == 0) { sb.Append("none;"); return; }
        sb.Append(Hex((int)fill.ForeColor.RGB)).Append('/').Append(Num(Math.Round((double)(float)fill.Transparency, 3))).Append(';');
    }
    // ppBorderTop 1, Left 2, Bottom 3, Right 4, DiagonalDown 5, DiagonalUp 6.
    public static string ReadTable(object tableObject, bool full) {
        dynamic table = tableObject;
        int rows = (int)table.Rows.Count;
        int cols = (int)table.Columns.Count;
        var sb = new StringBuilder();
        sb.Append(rows).Append('x').Append(cols);
        if (full) { sb.Append('|'); Fill(sb, "background", table.Background.Fill); }
        sb.Append('\n');
        for (int r = 1; r <= rows; r++) {
            for (int c = 1; c <= cols; c++) {
                dynamic cell = table.Cell(r, c);
                dynamic shape = cell.Shape;
                sb.Append(r - 1).Append(',').Append(c - 1).Append('|');
                Fill(sb, "fill", shape.Fill);
                if (full) {
                    string[] names = { "top", "left", "bottom", "right", "down", "up" };
                    for (int k = 1; k <= 6; k++) Line(sb, names[k - 1], Side(cell, k));
                }
                dynamic range = shape.TextFrame2.TextRange;
                // A padded position has no text: there is no first character to ask about.
                if ((int)range.Length == 0) { sb.Append("text=empty\n"); continue; }
                dynamic first = range.Characters(1, 1);
                dynamic font = first.Font;
                sb.Append("text=").Append(Hex((int)font.Fill.ForeColor.RGB)).Append('/').Append(Tri((int)font.Bold, "b"))
                  .Append('/').Append((string)font.Name);
                if (full) {
                    sb.Append('/').Append(Tri((int)font.Italic, "i")).Append('/').Append(Num((double)(float)font.Size))
                      .Append('/').Append(Tri((int)first.ParagraphFormat.Bullet.Visible, "bullet"));
                    sb.Append("|rect=").Append(Num((double)(float)shape.Left)).Append(',').Append(Num((double)(float)shape.Top))
                      .Append(',').Append(Num((double)(float)shape.Width)).Append(',').Append(Num((double)(float)shape.Height));
                }
                sb.Append('\n');
            }
        }
        return sb.ToString();
    }
    // msoTrue -1, msoFalse 0; anything else (msoTriStateMixed -2) is written raw.
    static string Tri(int v, string on) { return v == -1 ? on : v == 0 ? "-" : "raw" + v.ToString(Inv); }
}
'@

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

$twins = @($inputs.twins)
$ids = @($Decks -split ',' | Where-Object { $_ -ne '' })
if (($ids.Count -gt 0) -ne ($Tag -ne '')) { throw '-Tag and -Decks go together' }
if ($Tag -ne '' -and $Tag -notmatch '^[a-z0-9]+$') { throw "-Tag $Tag is not a lower-case tag" }
$control = @($inputs.decks | Where-Object { $_.id -eq 'control' })[0]
function Get-ControlRead([bool]$again) {
    $read = $control.PSObject.Copy()
    $suffix = if ($Tag -ne '') { "-$Tag" } else { '' }
    $read | Add-Member -NotePropertyName name -NotePropertyValue ("control$suffix" + $(if ($again) { '-again' } else { '' })) -Force
    $read | Add-Member -NotePropertyName again -NotePropertyValue $again -Force
    $read
}
if ($ids.Count -gt 0) {
    $named = @($inputs.decks | Where-Object { $ids -contains $_.id })
    $missing = @($ids | Where-Object { @($named.id) -notcontains $_ })
    if ($missing.Count -gt 0) { throw "no such deck in cascade-inputs.json: $($missing -join ', ')" }
    $order = @(Get-ControlRead $false) + $named + @(Get-ControlRead $true)
}
else {
    $order = @(Get-ControlRead $false) + @($inputs.decks | Where-Object { $_.id -ne 'control' }) + @(Get-ControlRead $true)
}
$total = $order.Count
$done = 0
$clock = [Diagnostics.Stopwatch]::StartNew()
foreach ($deck in $order) {
    $done++
    $name = if ($null -ne $deck.name) { $deck.name } else { $deck.id }
    $out = Join-Path (Join-Path $root 'readings') "$name.json"
    if (Test-Path -LiteralPath $out) { continue }
    $record = [ordered]@{ id = $name; deck = $deck.id; opened = $false; repaired = $null; error = $null; slides = @() }
    $file = Join-Path $root $deck.file
    $pres = $null
    try { $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoFalse); $record.opened = $true; $record.repaired = $false }
    catch {
        $record.error = $_.Exception.Message
        try { $pres = $app.Presentations.Open2007($file, $msoTrue, $msoFalse, $msoFalse, $msoTrue); $record.opened = $true; $record.repaired = $true }
        catch { $pres = $null }
    }
    if ($null -ne $pres) {
        try {
            for ($k = 1; $k -le $pres.Slides.Count; $k++) {
                $slide = $pres.Slides.Item($k)
                $shape = $slide.Shapes.Item(1)
                $reading = [ordered]@{
                    style = [string]$shape.Table.Style.Id
                    frame = @([double]$shape.Left, [double]$shape.Top, [double]$shape.Width, [double]$shape.Height)
                    cells = [CascadeReader]::ReadTable($shape.Table, $deck.read -eq 'full')
                    png = $null; bmp = $null
                }
                $png = Join-Path (Join-Path $root 'png') ('{0}-{1:D2}.png' -f $name, ($k - 1))
                $slide.Export($png, 'PNG', [int]$inputs.export.width, [int]$inputs.export.height)
                $reading.png = 'png/' + (Split-Path -Leaf $png)
                if ($k -eq 1 -and $twins -contains $deck.id -and -not $deck.again) {
                    $bmp = Join-Path (Join-Path $root 'twin') "$name.bmp"
                    $slide.Export($bmp, 'BMP', [int]$inputs.export.width, [int]$inputs.export.height)
                    $reading.bmp = "twin/$name.bmp"
                }
                $record.slides += , $reading
            }
            if ($deck.resave -and -not $deck.again) { $pres.SaveCopyAs((Join-Path (Join-Path $root 'resaved') "$name.pptx")) }
        }
        finally { try { $pres.Saved = $msoTrue; $pres.Close() } catch {} }
    }
    $record.session = $session
    $json = $record | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($out, $json, (New-Object System.Text.UTF8Encoding($false)))
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ('{0,4}/{1} {2,-16} {3,-9} {4,3} slides  {5:N0} s' -f $done, $total, $name, $state, $record.slides.Count, $clock.Elapsed.TotalSeconds)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
Write-Host 'done'
