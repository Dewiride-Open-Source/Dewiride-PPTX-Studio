# Open every .pptx in a directory with PowerPoint and report what it found.
#
#   powershell -File tools/corpus/tiers/a-generated/open-in-powerpoint.ps1 -Dir <dir>
#
# The gate a corpus deck has to pass is "opens in real PowerPoint with no repair
# prompt", and no amount of schema conformance substitutes for it: PowerPoint
# rejects some schema-legal markup for reasons it does not enumerate anywhere,
# and emits no diagnostic log at all. This is the only feedback loop there is.
#
# Two failure modes, and they look different:
#
#   - A refusal throws out of `Presentations.Open`, and the error text is the
#     only thing PowerPoint will ever say about it.
#   - A *silent repair* opens successfully with content quietly dropped, so the
#     script reports the shape count of every slide. Comparing those against
#     what the generator wrote is what catches it. PowerPoint counts a group as
#     one shape and counts placeholders, so the numbers are top-level shapes.
#
# Reads only the directory it is given, and writes nothing anywhere.

param(
    [Parameter(Mandatory = $true)][string]$Dir
)

$ErrorActionPreference = 'Stop'
$Dir = (Resolve-Path $Dir).Path

$app = New-Object -ComObject PowerPoint.Application
# ppAlertsNone. A repair prompt is a modal dialog: it blocks the COM call rather
# than returning an error, so this is not optional.
$app.DisplayAlerts = 1

$report = @()

try {
    # `.pptm` as well as `.pptx`: `a32-macros` is macro-enabled, and PowerPoint
    # refuses a package whose extension disagrees with the content type of its
    # main part - so the one deck that most needs opening is the one a `*.pptx`
    # filter would silently skip.
    # `-Include` needs a wildcard path, and there is deliberately no `-Recurse`:
    # this reads the one directory it was given and no other.
    foreach ($file in Get-ChildItem -Path (Join-Path $Dir '*') -Include *.pptx, *.pptm -File | Sort-Object Name) {
        $row = [ordered]@{ name = $file.Name; opened = $false; error = $null; slides = @() }
        try {
            # ReadOnly, not Untitled, no window: nothing is written back.
            $pres = $app.Presentations.Open($file.FullName, $true, $false, $false)
            $row.opened = $true
            $counts = @()
            foreach ($slide in $pres.Slides) {
                $counts += $slide.Shapes.Count
            }
            $row.slides = $counts
            $pres.Close()
        }
        catch {
            $row.error = $_.Exception.Message
        }
        $report += [pscustomobject]$row
    }
}
finally {
    try { $app.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

$report | ConvertTo-Json -Depth 4 -Compress
