[CmdletBinding()]
param(
    [ValidateSet(
        "show-map",
        "read-coords",
        "stop",
        "goto-home",
        "goto-work",
        "goto",
        "store-coords",
        "use-target",
        "set-mode-drill",
        "set-mode-syringe",
        "jog-ap-ant",
        "jog-ap-post",
        "jog-ml-left",
        "jog-ml-right",
        "jog-dv-sup",
        "jog-dv-inf",
        "set-step-ap",
        "set-step-ml",
        "set-step-dv"
    )]
    [string]$Action,
    [string]$Value,
    [string]$ProcessName = "StereoDrive"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$signature = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class StereoDriveWin32
{
    public const uint WM_GETTEXT = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const uint WM_SETTEXT = 0x000C;
    public const uint BM_CLICK = 0x00F5;
    public const uint CB_GETCOUNT = 0x0146;
    public const uint CB_GETLBTEXTLEN = 0x0149;
    public const uint CB_GETLBTEXT = 0x0148;
    public const uint CB_SETCURSEL = 0x014E;

    public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetDlgCtrlID(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);
}
"@

Add-Type -TypeDefinition $signature

function Get-MainWindowHandle {
    param([string]$ProcessName)

    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1

    if (-not $proc) {
        throw "No running process named '$ProcessName' with a main window was found."
    }

    return $proc.MainWindowHandle
}

function Get-ControlText {
    param([IntPtr]$Handle)

    $length = [StereoDriveWin32]::SendMessage($Handle, [StereoDriveWin32]::WM_GETTEXTLENGTH, [IntPtr]::Zero, [IntPtr]::Zero).ToInt32()
    $buffer = New-Object System.Text.StringBuilder ($length + 1)
    [void][StereoDriveWin32]::SendMessage($Handle, [StereoDriveWin32]::WM_GETTEXT, [IntPtr]$buffer.Capacity, $buffer)
    return $buffer.ToString()
}

function Get-ClassName {
    param([IntPtr]$Handle)

    $buffer = New-Object System.Text.StringBuilder 256
    [void][StereoDriveWin32]::GetClassName($Handle, $buffer, $buffer.Capacity)
    return $buffer.ToString()
}

function Get-WindowCaption {
    param([IntPtr]$Handle)

    $buffer = New-Object System.Text.StringBuilder 256
    [void][StereoDriveWin32]::GetWindowText($Handle, $buffer, $buffer.Capacity)
    return $buffer.ToString()
}

function Get-ChildControls {
    param([IntPtr]$ParentHandle)

    $rows = New-Object System.Collections.Generic.List[object]
    $callback = [StereoDriveWin32+EnumChildProc]{
        param([IntPtr]$hwnd, [IntPtr]$lParam)

        $rows.Add([pscustomobject]@{
            Handle = $hwnd
            ControlId = [StereoDriveWin32]::GetDlgCtrlID($hwnd)
            ClassName = Get-ClassName -Handle $hwnd
            Caption = Get-WindowCaption -Handle $hwnd
            Text = Get-ControlText -Handle $hwnd
        }) | Out-Null
        return $true
    }

    [void][StereoDriveWin32]::EnumChildWindows($ParentHandle, $callback, [IntPtr]::Zero)
    return $rows
}

function Get-ControlMap {
    param([IntPtr]$MainWindowHandle)

    $controls = Get-ChildControls -ParentHandle $MainWindowHandle
    $map = @{}
    foreach ($control in $controls) {
        if ($control.ControlId -gt 0) {
            $map[[string]$control.ControlId] = $control
        }
    }
    return $map
}

function Invoke-ButtonClick {
    param(
        [hashtable]$ControlMap,
        [int]$ControlId
    )

    $control = $ControlMap[[string]$ControlId]
    if (-not $control) {
        throw "Control ID $ControlId was not found."
    }

    [void][StereoDriveWin32]::SendMessage($control.Handle, [StereoDriveWin32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 150
}

function Set-EditText {
    param(
        [hashtable]$ControlMap,
        [int]$ControlId,
        [string]$Text
    )

    $control = $ControlMap[[string]$ControlId]
    if (-not $control) {
        throw "Control ID $ControlId was not found."
    }

    [void][StereoDriveWin32]::SendMessage($control.Handle, [StereoDriveWin32]::WM_SETTEXT, [IntPtr]::Zero, $Text)
    Start-Sleep -Milliseconds 100
}

function Get-ComboItems {
    param([IntPtr]$Handle)

    $count = [StereoDriveWin32]::SendMessage($Handle, [StereoDriveWin32]::CB_GETCOUNT, [IntPtr]::Zero, [IntPtr]::Zero).ToInt32()
    $items = @()
    for ($i = 0; $i -lt $count; $i++) {
        $len = [StereoDriveWin32]::SendMessage($Handle, [StereoDriveWin32]::CB_GETLBTEXTLEN, [IntPtr]$i, [IntPtr]::Zero).ToInt32()
        $buffer = New-Object System.Text.StringBuilder ($len + 1)
        [void][StereoDriveWin32]::SendMessage($Handle, [StereoDriveWin32]::CB_GETLBTEXT, [IntPtr]$i, $buffer)
        $items += $buffer.ToString()
    }
    return $items
}

function Set-ComboByText {
    param(
        [hashtable]$ControlMap,
        [int]$ControlId,
        [string]$WantedText
    )

    $control = $ControlMap[[string]$ControlId]
    if (-not $control) {
        throw "Control ID $ControlId was not found."
    }

    $items = Get-ComboItems -Handle $control.Handle
    $index = [Array]::IndexOf($items, $WantedText)
    if ($index -lt 0) {
        throw "Combo item '$WantedText' was not found for control ID $ControlId. Available: $($items -join ', ')"
    }

    [void][StereoDriveWin32]::SendMessage($control.Handle, [StereoDriveWin32]::CB_SETCURSEL, [IntPtr]$index, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 100
}

function Get-Coords {
    param([hashtable]$ControlMap)

    return [pscustomobject]@{
        AP_Current = (Get-ControlText -Handle $ControlMap["1138"].Handle)
        AP_Target = (Get-ControlText -Handle $ControlMap["1141"].Handle)
        ML_Current = (Get-ControlText -Handle $ControlMap["1139"].Handle)
        ML_Target = (Get-ControlText -Handle $ControlMap["1142"].Handle)
        DV_Current = (Get-ControlText -Handle $ControlMap["1140"].Handle)
        DV_Target = (Get-ControlText -Handle $ControlMap["1143"].Handle)
    }
}

$mainHandle = Get-MainWindowHandle -ProcessName $ProcessName
$controlMap = Get-ControlMap -MainWindowHandle $mainHandle

$buttonIds = @{
    "stop" = 1018
    "goto-home" = 1540
    "goto-work" = 1541
    "goto" = 1014
    "store-coords" = 1019
    "use-target" = 1152
    "set-mode-drill" = 1043
    "set-mode-syringe" = 1101
    "jog-ap-post" = 1103
    "jog-ap-ant" = 1102
    "jog-ml-left" = 1104
    "jog-ml-right" = 1105
    "jog-dv-sup" = 1106
    "jog-dv-inf" = 1107
}

$comboIds = @{
    "set-step-ap" = 1132
    "set-step-ml" = 1133
    "set-step-dv" = 1134
}

if ($Action -eq "show-map") {
    $controlMap.GetEnumerator() |
        Sort-Object { [int]$_.Key } |
        ForEach-Object {
            [pscustomobject]@{
                ControlId = [int]$_.Key
                ClassName = $_.Value.ClassName
                Caption = $_.Value.Caption
                Text = $_.Value.Text
            }
        }
    return
}

if ($Action -eq "read-coords") {
    Get-Coords -ControlMap $controlMap
    return
}

if ($comboIds.ContainsKey($Action)) {
    if (-not $Value) {
        throw "Action '$Action' requires -Value, for example -Value '0.1 mm'."
    }
    Set-ComboByText -ControlMap $controlMap -ControlId $comboIds[$Action] -WantedText $Value
    Get-Coords -ControlMap $controlMap
    return
}

if ($buttonIds.ContainsKey($Action)) {
    Invoke-ButtonClick -ControlMap $controlMap -ControlId $buttonIds[$Action]
    Get-Coords -ControlMap $controlMap
    return
}

throw "Unsupported action '$Action'."
