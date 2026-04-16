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
        "get-reference-status",
        "open-reference-panel",
        "dump-tools-menu",
        "set-reference-other",
        "set-reference-bregma",
        "set-reference-lambda",
        "set-drill-to-bregma",
        "set-syringe-to-bregma",
        "close-reference-panel",
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
    [ValidateSet("both", "bmclick", "wmcommand")]
    [string]$ClickMode = "both",
    [int]$DelaySeconds = 0,
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
    public const uint MF_BYPOSITION = 0x00000400;
    public const uint MN_GETHMENU = 0x01E1;
    public const uint WM_COMMAND = 0x0111;
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
    public static extern bool EnumWindows(EnumChildProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetMenu(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetMenuItemCount(IntPtr hMenu);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern int GetMenuString(IntPtr hMenu, uint uIDItem, StringBuilder lpString, int cchMax, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetMenuItemID(IntPtr hMenu, int nPos);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

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

function Get-TopLevelWindows {
    $rows = New-Object System.Collections.Generic.List[object]
    $callback = [StereoDriveWin32+EnumChildProc]{
        param([IntPtr]$hwnd, [IntPtr]$lParam)

        $windowPid = [uint32]0
        [void][StereoDriveWin32]::GetWindowThreadProcessId($hwnd, [ref]$windowPid)
        $rows.Add([pscustomobject]@{
            Handle = $hwnd
            ProcessId = [int]$windowPid
            ClassName = Get-ClassName -Handle $hwnd
            Caption = Get-WindowCaption -Handle $hwnd
        }) | Out-Null
        return $true
    }

    [void][StereoDriveWin32]::EnumWindows($callback, [IntPtr]::Zero)
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
        [int]$ControlId,
        [IntPtr]$MainWindowHandle,
        [string]$ClickMode
    )

    $control = $ControlMap[[string]$ControlId]
    if (-not $control) {
        throw "Control ID $ControlId was not found."
    }

    if ($ClickMode -eq "bmclick" -or $ClickMode -eq "both") {
        [void][StereoDriveWin32]::SendMessage($control.Handle, [StereoDriveWin32]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
    }

    if ($ClickMode -eq "wmcommand" -or $ClickMode -eq "both") {
        $wParam = [IntPtr]$ControlId
        [void][StereoDriveWin32]::SendMessage($MainWindowHandle, [StereoDriveWin32]::WM_COMMAND, $wParam, $control.Handle)
    }

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

function Normalize-MenuLabel {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    $normalized = $Text.ToLowerInvariant()
    $normalized = $normalized -replace '&', ''
    $normalized = $normalized -replace '\.\.\.', ''
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Find-MenuItem {
    param(
        [IntPtr]$MenuHandle,
        [string]$WantedLabel
    )

    $count = [StereoDriveWin32]::GetMenuItemCount($MenuHandle)
    if ($count -lt 0) {
        return $null
    }

    $wanted = Normalize-MenuLabel -Text $WantedLabel
    for ($i = 0; $i -lt $count; $i++) {
        $buffer = New-Object System.Text.StringBuilder 256
        [void][StereoDriveWin32]::GetMenuString($MenuHandle, [uint32]$i, $buffer, $buffer.Capacity, [StereoDriveWin32]::MF_BYPOSITION)
        $label = $buffer.ToString()
        if ((Normalize-MenuLabel -Text $label) -eq $wanted) {
            $rawId = [uint32][StereoDriveWin32]::GetMenuItemID($MenuHandle, $i)
            return [pscustomobject]@{
                Position = $i
                Label = $label
                Id = if ($rawId -eq [uint32]::MaxValue) { $null } else { $rawId }
                RawId = ('0x{0:X8}' -f $rawId)
                SubMenu = [StereoDriveWin32]::GetSubMenu($MenuHandle, $i)
            }
        }
    }

    return $null
}

function Get-MenuItems {
    param([IntPtr]$MenuHandle)

    $count = [StereoDriveWin32]::GetMenuItemCount($MenuHandle)
    if ($count -lt 0) {
        return @()
    }

    $items = @()
    for ($i = 0; $i -lt $count; $i++) {
        $buffer = New-Object System.Text.StringBuilder 256
        [void][StereoDriveWin32]::GetMenuString($MenuHandle, [uint32]$i, $buffer, $buffer.Capacity, [StereoDriveWin32]::MF_BYPOSITION)
        $label = $buffer.ToString()
        $rawId = [uint32][StereoDriveWin32]::GetMenuItemID($MenuHandle, $i)
        $items += [pscustomobject]@{
            Position = $i
            Label = $label
            NormalizedLabel = Normalize-MenuLabel -Text $label
            Id = if ($rawId -eq [uint32]::MaxValue) { $null } else { $rawId }
            RawId = ('0x{0:X8}' -f $rawId)
            HasSubMenu = ([StereoDriveWin32]::GetSubMenu($MenuHandle, $i) -ne [IntPtr]::Zero)
        }
    }

    return $items
}

function Get-PopupMenuWindow {
    param(
        [int]$ProcessId
    )

    return Get-TopLevelWindows |
        Where-Object { $_.ProcessId -eq $ProcessId -and $_.ClassName -eq "#32768" } |
        Select-Object -First 1
}

function Get-PopupMenuHandle {
    param([IntPtr]$PopupWindowHandle)

    $menu = [StereoDriveWin32]::SendMessage($PopupWindowHandle, [StereoDriveWin32]::MN_GETHMENU, [IntPtr]::Zero, [IntPtr]::Zero)
    if ($menu -eq [IntPtr]::Zero) {
        throw "Could not retrieve HMENU from popup menu window."
    }

    return $menu
}

function Invoke-PopupMenuItem {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId,
        [string]$Label
    )

    Invoke-ButtonClick -ControlMap $controlMap -ControlId 1010 -MainWindowHandle $MainWindowHandle -ClickMode $ClickMode
    Start-Sleep -Milliseconds 250

    $popup = Get-PopupMenuWindow -ProcessId $ProcessId
    if (-not $popup) {
        throw "Tools popup window was not found."
    }

    $menu = Get-PopupMenuHandle -PopupWindowHandle $popup.Handle
    $match = Find-MenuItem -MenuHandle $menu -WantedLabel $Label
    if (-not $match -or $null -eq $match.Id) {
        throw "Popup menu item '$Label' was not found."
    }

    [void][StereoDriveWin32]::SendMessage($MainWindowHandle, [StereoDriveWin32]::WM_COMMAND, [IntPtr]::new([int64]$match.Id), [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
}

function Open-ReferencePanel {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId
    )

    [void][StereoDriveWin32]::SendMessage($MainWindowHandle, [StereoDriveWin32]::WM_COMMAND, [IntPtr]::new(32809), [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
}

function Ensure-ReferencePanel {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId,
        [hashtable]$ControlMap
    )

    if (-not $ControlMap.ContainsKey("1095")) {
        Open-ReferencePanel -MainWindowHandle $MainWindowHandle -ProcessId $ProcessId
        return (Get-ControlMap -MainWindowHandle $MainWindowHandle)
    }

    return $ControlMap
}

function Get-ReferenceStatus {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId,
        [hashtable]$ControlMap
    )

    $ControlMap = Ensure-ReferencePanel -MainWindowHandle $MainWindowHandle -ProcessId $ProcessId -ControlMap $ControlMap

    $status = ""
    if ($ControlMap.ContainsKey("1097")) {
        $status = Get-ControlText -Handle $ControlMap["1097"].Handle
    }

    return [pscustomobject]@{
        ReferenceSelector = if ($ControlMap.ContainsKey("1387")) { Get-ControlText -Handle $ControlMap["1387"].Handle } else { "" }
        ReferenceStatus = $status
    }
}

$proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Select-Object -First 1

if (-not $proc) {
    throw "No running process named '$ProcessName' with a main window was found."
}

$mainHandle = $proc.MainWindowHandle
$mainProcessId = $proc.Id
$controlMap = Get-ControlMap -MainWindowHandle $mainHandle

$buttonIds = @{
    "stop" = 1018
    "goto-home" = 1540
    "goto-work" = 1541
    "goto" = 1014
    "store-coords" = 1019
    "use-target" = 1152
    "set-reference-other" = 1094
    "set-reference-bregma" = 1095
    "set-reference-lambda" = 1096
    "set-drill-to-bregma" = 1071
    "set-syringe-to-bregma" = 1072
    "close-reference-panel" = 1042
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

if ($Action -eq "open-reference-panel") {
    Open-ReferencePanel -MainWindowHandle $mainHandle -ProcessId $mainProcessId
    Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap (Get-ControlMap -MainWindowHandle $mainHandle)
    return
}

if ($Action -eq "dump-tools-menu") {
    if ($DelaySeconds -gt 0) {
        Start-Sleep -Seconds $DelaySeconds
    } else {
        Invoke-ButtonClick -ControlMap $controlMap -ControlId 1010 -MainWindowHandle $mainHandle -ClickMode $ClickMode
        Start-Sleep -Milliseconds 250
    }
    $popup = Get-PopupMenuWindow -ProcessId $mainProcessId
    if (-not $popup) {
        throw "Tools popup window was not found."
    }
    $menu = Get-PopupMenuHandle -PopupWindowHandle $popup.Handle
    Get-MenuItems -MenuHandle $menu
    return
}

if ($Action -eq "get-reference-status") {
    Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap $controlMap
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
    if ($Action -like "set-reference-*" -or $Action -like "set-*-to-bregma" -or $Action -eq "close-reference-panel") {
        $controlMap = Ensure-ReferencePanel -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap $controlMap
    }
    Invoke-ButtonClick -ControlMap $controlMap -ControlId $buttonIds[$Action] -MainWindowHandle $mainHandle -ClickMode $ClickMode
    [pscustomobject]@{
        Coords = Get-Coords -ControlMap $controlMap
        Reference = Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap (Get-ControlMap -MainWindowHandle $mainHandle)
    }
    return
}

throw "Unsupported action '$Action'."
