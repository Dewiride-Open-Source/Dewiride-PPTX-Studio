# Experiment T1, step 2 - ask PowerPoint what it resolved.
#
#   powershell -File tools/ground-truth/read-text.ps1 -Dir <work-dir>
#
# Reads `text-inputs.json`, opens each probe deck, and for every paragraph of
# every shape on every slide records what the object model says the *resolved*
# text formatting is.
#
# ## Why the object model, and why per paragraph
#
# The whole subject of sub-phase 3.1 is a run whose `a:rPr` declares nothing.
# `Font.Size` on that run is the answer to the entire cascade - one number,
# reported to a quarter point, with no rendering in the way. Give each level of
# the cascade a size no other level has and the number names the winner. There
# is nothing here a bitmap could answer better and several things it could not
# answer at all: `LineRuleWithin` distinguishes a line-spacing multiple from a
# fixed point size, and no pixel does.
#
# Both text frames are read. `TextFrame2.TextRange.Font.Name` reports the
# resolved typeface; `TextFrame.TextRange.Font.Name` has been observed to report
# the theme indirection `+mn-lt` instead, and which of the two a given build
# does is itself a measurement rather than a thing to assume.
#
# `Open2007` with `OpenAndRepair:=msoFalse`, for the reason C2 found: `Open`
# repairs silently, and a silently repaired deck reports success on exactly the
# files being asked about. A second attempt with repair allowed distinguishes
# REFUSED from REPAIRED.
#
# Read-only. Nothing is written back to any deck. Attaches to a running
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
$inputsPath = Join-Path $root 'text-inputs.json'
if (-not (Test-Path -LiteralPath $inputsPath)) {
    throw "no text-inputs.json in $root - run build-text-deck.ts first"
}

$text = Get-Content -LiteralPath $inputsPath -Raw
if ($text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
$inputs = $text | ConvertFrom-Json

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

# Every read is wrapped, because a property that throws on one shape kind is a
# missing field rather than a failed experiment - and a null in the JSON is a
# fact the analysis can score, while a crashed script is not.
function Read-Paragraph($para) {
    $out = [ordered]@{
        text             = $null
        size             = $null
        font             = $null
        fontEastAsian    = $null
        fontComplex      = $null
        bold             = $null
        italic           = $null
        underline        = $null
        strike           = $null
        caps             = $null
        spacing          = $null
        kerning          = $null
        rgb              = $null
        alignment        = $null
        indentLevel      = $null
        leftIndent       = $null
        firstLineIndent  = $null
        spaceBefore      = $null
        spaceAfter       = $null
        spaceWithin      = $null
        lineRuleBefore   = $null
        lineRuleAfter    = $null
        lineRuleWithin   = $null
        bulletVisible    = $null
        bulletCharacter  = $null
        bulletFont       = $null
        bulletRelSize    = $null
        bulletType       = $null
    }
    try { $out.text = [string]$para.Text } catch {}

    $font = $null
    try { $font = $para.Font } catch {}
    if ($null -ne $font) {
        try { $out.size = [double]$font.Size } catch {}
        try { $out.font = [string]$font.Name } catch {}
        try { $out.fontEastAsian = [string]$font.NameFarEast } catch {}
        try { $out.fontComplex = [string]$font.NameComplexScript } catch {}
        try { $out.bold = [int]$font.Bold } catch {}
        try { $out.italic = [int]$font.Italic } catch {}
        try { $out.underline = [int]$font.UnderlineStyle } catch {}
        try { $out.strike = [int]$font.Strike } catch {}
        try { $out.caps = [int]$font.Caps } catch {}
        try { $out.spacing = [double]$font.Spacing } catch {}
        try { $out.kerning = [double]$font.Kerning } catch {}
        try { $out.rgb = [int]$font.Fill.ForeColor.RGB } catch {}
    }

    $pf = $null
    try { $pf = $para.ParagraphFormat } catch {}
    if ($null -ne $pf) {
        try { $out.alignment = [int]$pf.Alignment } catch {}
        try { $out.indentLevel = [int]$pf.IndentLevel } catch {}
        try { $out.leftIndent = [double]$pf.LeftIndent } catch {}
        try { $out.firstLineIndent = [double]$pf.FirstLineIndent } catch {}
        try { $out.spaceBefore = [double]$pf.SpaceBefore } catch {}
        try { $out.spaceAfter = [double]$pf.SpaceAfter } catch {}
        try { $out.spaceWithin = [double]$pf.SpaceWithin } catch {}
        try { $out.lineRuleBefore = [int]$pf.LineRuleBefore } catch {}
        try { $out.lineRuleAfter = [int]$pf.LineRuleAfter } catch {}
        try { $out.lineRuleWithin = [int]$pf.LineRuleWithin } catch {}
        try { $out.bulletVisible = [int]$pf.Bullet.Visible } catch {}
        try { $out.bulletCharacter = [int]$pf.Bullet.Character } catch {}
        try { $out.bulletFont = [string]$pf.Bullet.Font.Name } catch {}
        try { $out.bulletRelSize = [double]$pf.Bullet.RelativeSize } catch {}
        try { $out.bulletType = [int]$pf.Bullet.Type } catch {}
    }
    return $out
}

function Read-Shape($shape, $prefix) {
    $s = [ordered]@{
        name        = $prefix + [string]$shape.Name
        placeholder = $null
        contained   = $null
        left        = $null
        top         = $null
        hasText     = $null
        paragraphs  = @()
        legacyFonts = @()
    }
    try { $s.placeholder = [int]$shape.PlaceholderFormat.Type } catch {}
    try { $s.contained = [int]$shape.PlaceholderFormat.ContainedType } catch {}
    try { $s.left = [double]$shape.Left } catch {}
    try { $s.top = [double]$shape.Top } catch {}

    $frame = $null
    try { $frame = $shape.TextFrame2 } catch {}
    if ($null -ne $frame) {
        try { $s.hasText = [int]$frame.HasText } catch {}
        $count = 0
        try { $count = [int]$frame.TextRange.Paragraphs().Count } catch {}
        for ($p = 1; $p -le $count; $p++) {
            $para = $null
            try { $para = $frame.TextRange.Paragraphs($p, 1) } catch {}
            if ($null -ne $para) { $s.paragraphs += Read-Paragraph $para }
        }

        # The v1 text frame, for the one question it answers differently:
        # whether `Font.Name` reports the resolved face or the theme indirection
        # that produced it.
        for ($p = 1; $p -le $count; $p++) {
            $legacy = $null
            try { $legacy = [string]$shape.TextFrame.TextRange.Paragraphs($p, 1).Font.Name } catch {}
            $s.legacyFonts += $legacy
        }
    }
    return $s
}

$decks = @()

foreach ($deck in $inputs.decks) {
    $file = Join-Path $root $deck.file

    $record = [ordered]@{
        deck     = $deck.deck
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
                $entry = [ordered]@{
                    index      = $i
                    name       = [string]$slide.Name
                    layoutName = $null
                    masterName = $null
                    shapes     = @()
                }
                try { $entry.layoutName = [string]$slide.CustomLayout.Name } catch {}
                try { $entry.masterName = [string]$slide.Master.Name } catch {}

                foreach ($shape in $slide.Shapes) {
                    $entry.shapes += Read-Shape $shape ''
                }

                # A shape on the layout or the master is drawn on the slide
                # without being on it, and `p:otherStyle` has to be read by
                # something. These two are the last candidates, so they are read
                # here rather than guessed at.
                try {
                    foreach ($shape in $slide.CustomLayout.Shapes) {
                        $entry.shapes += Read-Shape $shape 'layout:'
                    }
                }
                catch {}
                try {
                    foreach ($shape in $slide.Master.Shapes) {
                        $entry.shapes += Read-Shape $shape 'master:'
                    }
                }
                catch {}

                $record.slides += $entry
            }
        }
        finally {
            try { $pres.Close() } catch {}
        }
    }

    $decks += $record
    $state = if ($record.opened) { if ($record.repaired) { 'REPAIRED' } else { 'ok' } } else { 'REFUSED' }
    Write-Host ("{0,-22} {1,-9} {2} slide(s)" -f $deck.deck, $state, $record.slides.Count)
}

if ($created) { try { $app.Quit() } catch {} }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}

$json = @{ decks = $decks } | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $root 'text-readings.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Host 'done'
