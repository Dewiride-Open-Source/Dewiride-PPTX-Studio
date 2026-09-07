# Experiment C6, step 0 - ask PowerPoint to AUTHOR picture fills, then read what
# it wrote.
#
#   powershell -File tools/ground-truth/paint/blips/author.ps1 -Dir <out-dir> -Image <png>
#
# Writes pp-blip-fill.pptx, pp-blip-effects.pptx and author-log.json. `analyse.ts`
# unzips ppt/slides/slide1.xml from each and reads the markup PowerPoint chose.
#
# Every question this settles is one about units, and units are exactly what a
# renderer written from the schema alone gets wrong silently: a:tile/@tx is a
# coordinate and the UI offers points, @sx is a percentage and the UI offers a
# fraction, and nothing in ECMA-376 says which way @algn anchors the grid. The
# application that defines the format can be made to answer in writing.
#
# Creates presentations and saves them into <out-dir>; opens nothing it did not
# create. Attaches to a running PowerPoint if there is one and never quits one it
# did not start.
param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [Parameter(Mandatory = $true)][string]$Image
)

$ErrorActionPreference = 'Stop'
$msoTrue = -1; $msoFalse = 0
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24
$ppAlertsNone = 1

$root = (Resolve-Path -LiteralPath $Dir).Path
$png = (Resolve-Path -LiteralPath $Image).Path

$created = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
$app.DisplayAlerts = $ppAlertsNone

$log = @()

function New-Deck {
    $p = $app.Presentations.Add($msoFalse)
    $p.PageSetup.SlideSize = 15   # ppSlideSizeOnScreen16x9
    return $p
}

# --------------------------------------------------------------- picture fills ---
# One shape per question. The name is the key `analyse.ts` matches the emitted
# a:blipFill against, so it has to survive the round trip: PowerPoint keeps
# Shape.Name in p:cNvPr/@name.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
function Add-Fill([string]$name, [scriptblock]$setup) {
    $script:n++
    $col = (($script:n - 1) % 6); $row = [Math]::Floor(($script:n - 1) / 6)
    try {
        $sh = $script:slide.Shapes.AddShape(1, 12 + $col * 110, 12 + $row * 90, 100, 80)
        $sh.Line.Visible = $msoFalse
        $sh.Name = $name
        & $setup $sh
        $script:log += [ordered]@{ deck = 'fill'; name = $name; ok = $true; error = $null }
    }
    catch {
        $script:log += [ordered]@{ deck = 'fill'; name = $name; ok = $false; error = $_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}

Add-Fill 'stretch' { param($sh) $sh.Fill.UserPicture($png) }
Add-Fill 'tile-default' { param($sh) $sh.Fill.UserTextured($png) }

# Offsets: the UI calls these points. If @tx comes back as 114300 for 9 pt then
# it is EMU, and the sign tells us which way the grid moves.
Add-Fill 'tile-offset-pos' {
    param($sh)
    $sh.Fill.UserTextured($png)
    $sh.Fill.TextureOffsetX = 9
    $sh.Fill.TextureOffsetY = 18
}
Add-Fill 'tile-offset-neg' {
    param($sh)
    $sh.Fill.UserTextured($png)
    $sh.Fill.TextureOffsetX = -9
    $sh.Fill.TextureOffsetY = -4.5
}

# Scale: the UI offers a fraction. 0.5 landing as sx="50000" makes @sx a
# hundred-thousandths percentage of the tile's natural size, which is the only
# reading under which the natural size itself has to be known.
Add-Fill 'tile-scale-half' {
    param($sh)
    $sh.Fill.UserTextured($png)
    $sh.Fill.TextureHorizontalScale = 0.5
    $sh.Fill.TextureVerticalScale = 0.25
}

# MsoTextureAlignment: 0 tl, 1 t, 2 tr, 3 l, 4 ctr, 5 r, 6 bl, 7 b, 8 br. Nine
# shapes because @algn is the one attribute whose meaning the schema states as a
# name and never as a rule.
$alignments = @('tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br')
for ($a = 0; $a -lt 9; $a++) {
    $value = $a
    $label = $alignments[$a]
    Add-Fill ("tile-algn-{0}" -f $label) {
        param($sh)
        $sh.Fill.UserTextured($png)
        $sh.Fill.TextureAlignment = $value
    }
}

Add-Fill 'fill-transparency-40' {
    param($sh)
    $sh.Fill.UserPicture($png)
    $sh.Fill.Transparency = 0.4
}
Add-Fill 'fill-rotate-with-shape-off' {
    param($sh)
    $sh.Fill.UserPicture($png)
    $sh.Fill.RotateWithObject = $msoFalse
    $sh.Rotation = 30
}
Add-Fill 'fill-rotate-with-shape-on' {
    param($sh)
    $sh.Fill.UserPicture($png)
    $sh.Fill.RotateWithObject = $msoTrue
    $sh.Rotation = 30
}

$fillPath = Join-Path $root 'pp-blip-fill.pptx'
$pres.SaveAs($fillPath, $ppSaveAsOpenXMLPresentation)
$pres.Close()
Write-Host ("picture fills: {0}" -f $n)

# ------------------------------------------------------------- blip effects ---
# The a:blip colour effects are reachable only through PictureFormat, which lives
# on a picture rather than on a fill. The markup is the same element either way.
$pres = New-Deck
$slide = $pres.Slides.Add(1, $ppLayoutBlank)
$n = 0
function Add-Picture([string]$name, [scriptblock]$setup) {
    $script:n++
    $col = (($script:n - 1) % 6); $row = [Math]::Floor(($script:n - 1) / 6)
    try {
        $sh = $script:slide.Shapes.AddPicture($png, $msoFalse, $msoTrue, 12 + $col * 110, 12 + $row * 90, 100, 80)
        $sh.Name = $name
        & $setup $sh
        $script:log += [ordered]@{ deck = 'effects'; name = $name; ok = $true; error = $null }
    }
    catch {
        $script:log += [ordered]@{ deck = 'effects'; name = $name; ok = $false; error = $_.Exception.Message }
        try { if ($sh) { $sh.Delete() } } catch {}
    }
}

Add-Picture 'plain' { param($sh) }
# msoPictureGrayscale 2, msoPictureBlackAndWhite 3, msoPictureWatermark 4.
Add-Picture 'colortype-grayscale' { param($sh) $sh.PictureFormat.ColorType = 2 }
Add-Picture 'colortype-blackwhite' { param($sh) $sh.PictureFormat.ColorType = 3 }
Add-Picture 'colortype-watermark' { param($sh) $sh.PictureFormat.ColorType = 4 }
# Brightness and contrast are 0..1 in the object model and signed percentages in
# the file, so 0.7 becoming bright="40000" is the whole conversion rule.
Add-Picture 'lum-bright-70' { param($sh) $sh.PictureFormat.Brightness = 0.7 }
Add-Picture 'lum-bright-30' { param($sh) $sh.PictureFormat.Brightness = 0.3 }
Add-Picture 'lum-contrast-80' { param($sh) $sh.PictureFormat.Contrast = 0.8 }
Add-Picture 'lum-contrast-20' { param($sh) $sh.PictureFormat.Contrast = 0.2 }
Add-Picture 'lum-both' {
    param($sh)
    $sh.PictureFormat.Brightness = 0.65
    $sh.PictureFormat.Contrast = 0.35
}
# Crop is in points and the file wants a percentage of the source, so this is the
# only place the image's own pixel size enters the conversion.
Add-Picture 'crop-lt' {
    param($sh)
    $sh.PictureFormat.CropLeft = 25
    $sh.PictureFormat.CropTop = 12.5
}
Add-Picture 'crop-rb' {
    param($sh)
    $sh.PictureFormat.CropRight = 20
    $sh.PictureFormat.CropBottom = 10
}
Add-Picture 'transparent-colour' {
    param($sh)
    $sh.PictureFormat.TransparencyColor = 2842709   # 0x2B60D5 BGR of D5602B
    $sh.PictureFormat.TransparentBackground = $msoTrue
}

$effectPath = Join-Path $root 'pp-blip-effects.pptx'
$pres.SaveAs($effectPath, $ppSaveAsOpenXMLPresentation)
$pres.Close()
Write-Host ("blip effects: {0}" -f $n)

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{
    files = @('pp-blip-fill.pptx', 'pp-blip-effects.pptx')
    image = $png
    log   = $log
} | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText((Join-Path $root 'author-log.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "done"
