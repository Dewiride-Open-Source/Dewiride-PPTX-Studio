# Sub-phase 1.1, experiments E6 and E8.
#
# Produces two decks with PowerPoint 365 so that `analyse-conventions.ts` can
# measure what Microsoft's own serializer writes. Nothing here is committed:
# the decks go to a directory the caller names, and what lands in the
# repository is the measurement, the same posture as `eot-headers.json`.
#
#   powershell -File tools/corpus/tiers/b-authored/probe-conventions.ps1 -Out <dir>
#   node tools/corpus/tiers/b-authored/analyse-conventions.ts <dir>
#
# E6 asks what a modern build writes for an embedded OLE object. Microsoft
# stopped emitting the VML fallback at build 2205 while continuing to read it,
# so whether our synthetic OLE deck needs a `vmlDrawing` part is a measurement,
# not a reading of the specification.
#
# E8 asks for the byte conventions - BOM, declaration, line endings, escaping,
# self-closing form, ZIP flags. That is the specification our own writer has to
# satisfy, and it cannot be derived from the schema.
#
# Reads nothing outside the output directory. The OLE object is a *new* empty
# workbook created in place, so no file on this machine is opened to make it.

param(
    [Parameter(Mandatory = $true)][string]$Out
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Out)) { New-Item -ItemType Directory -Force $Out | Out-Null }
$Out = (Resolve-Path $Out).Path

$olePath = Join-Path $Out 'ole.pptx'
$plainPath = Join-Path $Out 'plain.pptx'
foreach ($p in @($olePath, $plainPath)) { if (Test-Path $p) { Remove-Item $p -Force } }

$app = New-Object -ComObject PowerPoint.Application
# ppAlertsNone. A repair prompt is a modal dialog: it blocks the COM call
# rather than returning an error, so this is not optional.
$app.DisplayAlerts = 1

$result = @{}

try {
    # --- E8: text chosen to force every escaping decision -------------------
    $pres = $app.Presentations.Add(-1)
    $slide = $pres.Slides.Add(1, 2)                      # ppLayoutText
    $slide.Shapes.Item(1).TextFrame.TextRange.Text =
        'Ampersand & less < greater > quote " apostrophe '' end'
    # Leading and trailing spaces, to settle whether PowerPoint writes
    # xml:space="preserve" on an a:t the way Word does on a w:t.
    $slide.Shapes.Item(2).TextFrame.TextRange.Text = '  leading and trailing spaces  '
    $pres.SaveAs($plainPath, 24, 0)                      # ppSaveAsOpenXMLPresentation
    $pres.Close()

    # --- E6: an embedded OLE object ----------------------------------------
    $pres2 = $app.Presentations.Add(-1)
    $slide2 = $pres2.Slides.Add(1, 12)                   # ppLayoutBlank
    $shape = $slide2.Shapes.AddOLEObject(100, 100, 400, 300, 'Excel.Sheet.12')
    $result['oleProgId'] = $shape.OLEFormat.ProgID
    $pres2.SaveAs($olePath, 24, 0)
    $pres2.Close()
}
finally {
    try { $app.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

$result['plain'] = $plainPath
$result['ole'] = $olePath
$result | ConvertTo-Json -Depth 3
