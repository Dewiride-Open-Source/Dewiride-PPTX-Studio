# Ask the real PowerPoint whether it will open one file, and say so in an exit code.
#
#   powershell -File packages/cli/scripts/powerpoint-oracle.ps1 -File <path>
#
#   exit 0  opened, with no repair
#   exit 1  PowerPoint would not open it as it stands
#   exit 2  the harness itself failed - PowerPoint missing, COM refused, a crash
#
# This is the oracle behind `pptx-studio bisect --oracle powerpoint`. The third
# exit code is not decoration: a bisector that read "the harness is broken" as
# "the file is broken" would blame whichever change happened to be applied when
# PowerPoint fell over, and do it with a straight face.
#
# ## Why Open2007 and not Open
#
# `Presentations.Open` has no repair parameter. `Presentations.Open2007` does -
# `OpenAndRepair`, documented as **defaulting to msoTrue** - and there is no
# documented way to make `Open` behave differently, so a file PowerPoint would
# repair opens through `Open` *successfully, already repaired*, and reports
# `opened = $true`.
#
# That is the whole difficulty of this project stated in one API. The prompt a
# user sees - "PowerPoint found a problem with content" - is a repair, and under
# automation the repair is silent: `DisplayAlerts = ppAlertsNone` is documented
# to choose the message box's default answer and carry on, and the default
# answer to that box is Repair. So the obvious script reports success on exactly
# the files this project exists to avoid producing.
#
# Passing `OpenAndRepair:=msoFalse` is what turns the silent repair back into a
# catchable error. `-AllowRepair` flips it, which is how the two are told apart:
# a file that fails without repair and opens with it is one PowerPoint *repairs*;
# a file that fails both ways is one it *refuses*.
#
# ## What it will not do to your machine
#
#   - Read-only, no window, and nothing is ever written back.
#   - `AutomationSecurity = msoAutomationSecurityForceDisable`, so a macro in a
#     `.pptm` cannot run. The property defaults to *Low*, which enables all
#     macros, and this script's whole job is opening files that are wrong on
#     purpose.
#   - If PowerPoint is already running it attaches to that instance and never
#     quits it, because it is yours. It only quits one it started itself, and
#     only when asked with -Quit.
#
# Reads the one file it is given and nothing else.

param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Test')][string]$File,
    [Parameter(ParameterSetName = 'Test')][switch]$AllowRepair,
    [Parameter(ParameterSetName = 'Test')][switch]$Quit,
    [Parameter(Mandatory = $true, ParameterSetName = 'Shutdown')][switch]$QuitOnly
)

$ErrorActionPreference = 'Stop'

# msoFalse is 0 and msoTrue is -1. Spelled out because -1 as a literal in an
# argument list is the kind of thing that gets "tidied" into $true, and $true
# marshals to the same value only by luck of the VARIANT_BOOL representation.
$msoTrue = -1
$msoFalse = 0
$ppAlertsNone = 1
$msoAutomationSecurityForceDisable = 3

function Get-PowerPoint {
    # Attach to a running instance if there is one. Starting a second PowerPoint
    # per oracle run costs about three seconds each, and a bisection is dozens
    # of runs; more to the point, quitting one we did not start would close
    # whatever the person at this machine had open.
    $created = $false
    $app = $null
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
    catch {
        $app = New-Object -ComObject PowerPoint.Application
        $created = $true
    }
    return @{ app = $app; created = $created }
}

if ($PSCmdlet.ParameterSetName -eq 'Shutdown') {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        $app.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
    catch {
        # Nothing running is the outcome we wanted anyway.
    }
    exit 0
}

if (-not (Test-Path -LiteralPath $File -PathType Leaf)) {
    [pscustomobject]@{ ok = $false; harness = $true; error = "no such file: $File" } |
        ConvertTo-Json -Compress
    exit 2
}
$path = (Resolve-Path -LiteralPath $File).Path

$handle = $null
try {
    $handle = Get-PowerPoint
}
catch {
    [pscustomobject]@{ ok = $false; harness = $true; error = $_.Exception.Message } |
        ConvertTo-Json -Compress
    exit 2
}

$app = $handle.app
$result = [ordered]@{
    file    = $path
    ok      = $false
    harness = $false
    repair  = [bool]$AllowRepair
    slides  = 0
    shapes  = @()
    error   = $null
}
$code = 1

try {
    $app.DisplayAlerts = $ppAlertsNone
    $app.AutomationSecurity = $msoAutomationSecurityForceDisable

    $repair = if ($AllowRepair) { $msoTrue } else { $msoFalse }
    $pres = $null
    try {
        # FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair
        $pres = $app.Presentations.Open2007($path, $msoTrue, $msoFalse, $msoFalse, $repair)
        $result.ok = $true
        $result.slides = $pres.Slides.Count
        # Shape counts per slide, because a repair is not the only way to lose
        # content: PowerPoint will also open a file happily and drop a shape it
        # did not like. Comparing these against what the writer put in is the
        # only way to see that.
        $counts = @()
        foreach ($slide in $pres.Slides) { $counts += $slide.Shapes.Count }
        $result.shapes = $counts
        $code = 0
    }
    catch {
        # The one sentence PowerPoint will ever say about it.
        $result.error = $_.Exception.Message
        $code = 1
    }
    finally {
        if ($null -ne $pres) { try { $pres.Close() } catch {} }
    }
}
catch {
    $result.harness = $true
    $result.error = $_.Exception.Message
    $code = 2
}
finally {
    if ($handle.created -and $Quit) { try { $app.Quit() } catch {} }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
}

[pscustomobject]$result | ConvertTo-Json -Depth 4 -Compress
exit $code
