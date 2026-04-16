[CmdletBinding()]
param(
    [string]$ProcessName = "StereoDrive",
    [string]$WindowTitleRegex = ".*",
    [int]$MaxDepth = 5,
    [int]$MaxChildrenPerNode = 200,
    [int]$DelaySeconds = 0,
    [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Get-PatternNames {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $patternNames = @()
    foreach ($pattern in $Element.GetSupportedPatterns()) {
        if ($null -ne $pattern) {
            $patternNames += $pattern.ProgrammaticName
        }
    }
    return $patternNames | Sort-Object -Unique
}

function Get-ElementRecord {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element,
        [Parameter(Mandatory = $true)]
        [int]$Depth,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $current = $Element.Current
    $rect = $current.BoundingRectangle

    return [ordered]@{
        path = $Path
        depth = $Depth
        name = $current.Name
        automation_id = $current.AutomationId
        class_name = $current.ClassName
        control_type = $current.ControlType.ProgrammaticName
        framework_id = $current.FrameworkId
        native_window_handle = $current.NativeWindowHandle
        process_id = $current.ProcessId
        is_enabled = $current.IsEnabled
        is_offscreen = $current.IsOffscreen
        has_keyboard_focus = $current.HasKeyboardFocus
        bounding_rectangle = [ordered]@{
            left = $rect.Left
            top = $rect.Top
            width = $rect.Width
            height = $rect.Height
        }
        patterns = @(Get-PatternNames -Element $Element)
    }
}

function Add-ElementTree {
     param(
         [Parameter(Mandatory = $true)]
         [System.Windows.Automation.AutomationElement]$Root,
         [AllowEmptyCollection()]
         [Parameter(Mandatory = $true)]
         [System.Collections.Generic.List[object]]$Rows,
         [Parameter(Mandatory = $true)]
         [int]$MaxDepth,
        [Parameter(Mandatory = $true)]
        [int]$MaxChildrenPerNode,
        [int]$Depth = 0,
        [string]$Path = "0"
    )

    $Rows.Add((Get-ElementRecord -Element $Root -Depth $Depth -Path $Path)) | Out-Null

    if ($Depth -ge $MaxDepth) {
        return
    }

    $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
    $child = $walker.GetFirstChild($Root)
    $index = 0
    while ($null -ne $child -and $index -lt $MaxChildrenPerNode) {
        Add-ElementTree -Root $child -Rows $Rows -MaxDepth $MaxDepth -MaxChildrenPerNode $MaxChildrenPerNode -Depth ($Depth + 1) -Path "$Path.$index"
        $child = $walker.GetNextSibling($child)
        $index++
    }
}

function Get-TargetWindows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,
        [Parameter(Mandatory = $true)]
        [string]$WindowTitleRegex
    )

    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) {
        throw "No running process named '$ProcessName' was found."
    }

    $windows = @()
    foreach ($proc in $procs) {
        if ($proc.MainWindowHandle -eq 0) {
            continue
        }

        try {
            $element = [System.Windows.Automation.AutomationElement]::FromHandle($proc.MainWindowHandle)
            if ($null -eq $element) {
                continue
            }

            $title = $element.Current.Name
            if ($title -match $WindowTitleRegex) {
                $windows += $element
            }
        } catch {
            continue
        }
    }

    if (-not $windows) {
        throw "A '$ProcessName' process is running, but no main window matched title regex '$WindowTitleRegex'."
    }

    return $windows
}

$rows = New-Object 'System.Collections.Generic.List[object]'

if ($DelaySeconds -gt 0) {
    Start-Sleep -Seconds $DelaySeconds
}

$targets = Get-TargetWindows -ProcessName $ProcessName -WindowTitleRegex $WindowTitleRegex

foreach ($target in $targets) {
    Add-ElementTree -Root $target -Rows $rows -MaxDepth $MaxDepth -MaxChildrenPerNode $MaxChildrenPerNode
}

$json = $rows | ConvertTo-Json -Depth 6

if ($OutFile) {
    $dir = Split-Path -Parent $OutFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
    [System.IO.File]::WriteAllText($OutFile, $json)
    Write-Host "Wrote UI tree to $OutFile"
} else {
    $json
}
