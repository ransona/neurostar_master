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
        "show-injectomate",
        "hide-injectomate",
        "injectomate-map",
        "scan-injectomate-value",
        "scan-injectomate-region",
        "open-injectomate-calibrate",
        "probe-injectomate-calibrate",
        "probe-scale-control",
        "probe-plunger-gauge-control",
        "read-scale-popup-api",
        "scan-process-memory-value",
        "refine-process-memory-value",
        "clear-process-memory-candidates",
        "scan-open-windows",
        "set-injection-volume",
        "set-syringe-type",
        "inject",
        "fill",
        "close-injectomate",
        "dump-tools-menu",
        "test-hidden-reference-bregma",
        "test-hidden-drill-to-bregma",
        "probe-command",
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
    [int]$CommandId = 0,
    [string]$ProcessName = "StereoDrive"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$MemoryCandidatesPath = Join-Path $PSScriptRoot "stereodrive_memory_candidates.json"

$signature = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class StereoDriveWin32
{
    public const uint MF_BYPOSITION = 0x00000400;
    public const uint MN_GETHMENU = 0x01E1;
    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_GETTEXT = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const uint WM_SETTEXT = 0x000C;
    public const uint BM_CLICK = 0x00F5;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    public const uint CB_GETCOUNT = 0x0146;
    public const uint CB_GETLBTEXTLEN = 0x0149;
    public const uint CB_GETLBTEXT = 0x0148;
    public const uint CB_SETCURSEL = 0x014E;
    public const uint PROCESS_QUERY_INFORMATION = 0x0400;
    public const uint PROCESS_VM_READ = 0x0010;
    public const uint MEM_COMMIT = 0x1000;
    public const uint PAGE_NOACCESS = 0x01;
    public const uint PAGE_GUARD = 0x100;

    public delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MEMORY_BASIC_INFORMATION
    {
        public IntPtr BaseAddress;
        public IntPtr AllocationBase;
        public uint AllocationProtect;
        public UIntPtr RegionSize;
        public uint State;
        public uint Protect;
        public uint Type;
    }

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

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool IsWindowEnabled(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

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

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, int dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, UIntPtr nSize, out UIntPtr lpNumberOfBytesRead);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern UIntPtr VirtualQueryEx(IntPtr hProcess, IntPtr lpAddress, out MEMORY_BASIC_INFORMATION lpBuffer, UIntPtr dwLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);
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

    $buffer = New-Object System.Text.StringBuilder 1024
    $result = [IntPtr]::Zero
    $sent = [StereoDriveWin32]::SendMessageTimeout(
        $Handle,
        [StereoDriveWin32]::WM_GETTEXT,
        [IntPtr]$buffer.Capacity,
        $buffer,
        [StereoDriveWin32]::SMTO_ABORTIFHUNG,
        100,
        [ref]$result
    )
    if ($sent -eq [IntPtr]::Zero) {
        return ""
    }
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

function Get-WindowRectObject {
    param([IntPtr]$Handle)

    try {
        $rect = New-Object StereoDriveWin32+RECT
        [void][StereoDriveWin32]::GetWindowRect($Handle, [ref]$rect)
        return [pscustomobject]@{
            Left = $rect.Left
            Top = $rect.Top
            Right = $rect.Right
            Bottom = $rect.Bottom
            Width = $rect.Right - $rect.Left
            Height = $rect.Bottom - $rect.Top
        }
    } catch {
        return $null
    }
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
            Rect = Get-WindowRectObject -Handle $hwnd
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
            Text = ""
            Rect = Get-WindowRectObject -Handle $hwnd
            Visible = [StereoDriveWin32]::IsWindowVisible($hwnd)
        }) | Out-Null
        return $true
    }

    [void][StereoDriveWin32]::EnumWindows($callback, [IntPtr]::Zero)
    return $rows
}

function Get-ProcessWindows {
    param([int]$ProcessId)

    return Get-TopLevelWindows | Where-Object { $_.ProcessId -eq $ProcessId }
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

function Invoke-ControlMouseClick {
    param(
        [hashtable]$ControlMap,
        [int]$ControlId,
        [IntPtr]$MainWindowHandle
    )

    $control = $ControlMap[[string]$ControlId]
    if (-not $control) {
        throw "Control ID $ControlId was not found."
    }
    if (-not $control.Rect) {
        throw "Control ID $ControlId has no rectangle."
    }

    $x = [int](($control.Rect.Left + $control.Rect.Right) / 2)
    $y = [int](($control.Rect.Top + $control.Rect.Bottom) / 2)
    [void][StereoDriveWin32]::SetForegroundWindow($MainWindowHandle)
    Start-Sleep -Milliseconds 150
    [void][StereoDriveWin32]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 80
    [StereoDriveWin32]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 40
    [StereoDriveWin32]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 350
}

function Invoke-DirectCommand {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$CommandId
    )

    [void][StereoDriveWin32]::SendMessage($MainWindowHandle, [StereoDriveWin32]::WM_COMMAND, [IntPtr]::new([int64]$CommandId), [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
}

function Get-WindowSnapshot {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId
    )

    $processWindows = Get-ProcessWindows -ProcessId $ProcessId |
        Select-Object Handle, ProcessId, ClassName, Caption

    $controlMap = Get-ControlMap -MainWindowHandle $MainWindowHandle
    $interesting = $controlMap.GetEnumerator() |
        Sort-Object { [int]$_.Key } |
        ForEach-Object {
            [pscustomobject]@{
                ControlId = [int]$_.Key
                ClassName = $_.Value.ClassName
                Caption = $_.Value.Caption
                Text = $_.Value.Text
            }
        }

    return [pscustomobject]@{
        Windows = $processWindows
        Controls = $interesting
    }
}

function Get-NumericCandidatesFromControls {
    param([object[]]$Rows)

    $numericPattern = "-?\d+(?:\.\d+)?"
    $candidates = @()
    foreach ($row in @($Rows)) {
        foreach ($fieldName in @("Text", "Caption")) {
            $property = $row.PSObject.Properties[$fieldName]
            if (-not $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                continue
            }
            foreach ($match in [regex]::Matches([string]$property.Value, $numericPattern)) {
                $candidates += [pscustomobject]@{
                    Value = [double]$match.Value
                    SourceField = $fieldName
                    Handle = $row.Handle
                    ProcessId = if ($row.PSObject.Properties["ProcessId"]) { $row.ProcessId } else { $null }
                    ControlId = if ($row.PSObject.Properties["ControlId"]) { $row.ControlId } else { $null }
                    ClassName = $row.ClassName
                    Caption = if ($row.PSObject.Properties["Caption"]) { $row.Caption } else { "" }
                    Text = if ($row.PSObject.Properties["Text"]) { $row.Text } else { "" }
                    Rect = if ($row.PSObject.Properties["Rect"]) { $row.Rect } else { $null }
                }
            }
        }
    }
    return $candidates
}

function Get-WindowTreeSnapshot {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId,
        [bool]$IncludeNearbyWindows = $false
    )

    $allWindows = @(
        Get-TopLevelWindows |
            Where-Object { $_.Visible -and $_.Rect -and $_.Rect.Width -gt 0 -and $_.Rect.Height -gt 0 }
    )
    $processWindows = @($allWindows | Where-Object { $_.ProcessId -eq $ProcessId })
    $mainWindow = $allWindows | Where-Object { $_.Handle -eq $MainWindowHandle } | Select-Object -First 1
    $candidateWindows = @($processWindows)
    if ($IncludeNearbyWindows -and $mainWindow) {
        $left = $mainWindow.Rect.Left - 80
        $top = $mainWindow.Rect.Top - 80
        $right = $mainWindow.Rect.Right + 160
        $bottom = $mainWindow.Rect.Bottom + 160
        $nearbyWindows = @(
            $allWindows |
                Where-Object {
                    $_.ProcessId -ne $ProcessId -and
                    $_.Rect.Right -ge $left -and
                    $_.Rect.Left -le $right -and
                    $_.Rect.Bottom -ge $top -and
                    $_.Rect.Top -le $bottom
                }
        )
        $candidateWindows = @($candidateWindows + $nearbyWindows | Sort-Object Handle -Unique)
    }

    $windowTrees = @()
    $allRows = @()
    foreach ($window in $candidateWindows) {
        $children = @(Get-ChildControls -ParentHandle $window.Handle)
        $allRows += $window
        $allRows += $children
        $windowTrees += [pscustomobject]@{
            Handle = $window.Handle
            ProcessId = $window.ProcessId
            ClassName = $window.ClassName
            Caption = $window.Caption
            Text = $window.Text
            Rect = $window.Rect
            Children = $children
        }
    }

    $mainChildren = @()
    if (-not ($candidateWindows | Where-Object { $_.Handle -eq $MainWindowHandle })) {
        $mainChildren = @(Get-ChildControls -ParentHandle $MainWindowHandle)
        $allRows += $mainChildren
    }

    return [pscustomobject]@{
        MainWindow = $mainWindow
        ProcessWindows = $processWindows
        CandidateWindows = $windowTrees
        MainChildren = $mainChildren
        NumericCandidates = @(Get-NumericCandidatesFromControls -Rows $allRows)
    }
}

function Write-WindowTreeReport {
    param([pscustomobject]$Snapshot)

    Write-Output "StereoDrive open-window scan"
    if ($Snapshot.MainWindow) {
        Write-Output ("MainWindow handle={0} class={1} caption='{2}' rect=({3},{4},{5},{6})" -f `
            $Snapshot.MainWindow.Handle,
            $Snapshot.MainWindow.ClassName,
            $Snapshot.MainWindow.Caption,
            $Snapshot.MainWindow.Rect.Left,
            $Snapshot.MainWindow.Rect.Top,
            $Snapshot.MainWindow.Rect.Right,
            $Snapshot.MainWindow.Rect.Bottom)
    }

    Write-Output ""
    Write-Output ("Scanned windows: {0}" -f @($Snapshot.CandidateWindows).Count)
    foreach ($window in @($Snapshot.CandidateWindows)) {
        Write-Output ("  handle={0} class={1} caption='{2}' rect=({3},{4},{5},{6})" -f `
            $window.Handle,
            $window.ClassName,
            $window.Caption,
            $window.Rect.Left,
            $window.Rect.Top,
            $window.Rect.Right,
            $window.Rect.Bottom)
    }

    Write-Output ""
    Write-Output ("Numeric candidates: {0}" -f @($Snapshot.NumericCandidates).Count)
    foreach ($candidate in @($Snapshot.NumericCandidates)) {
        $rectText = ""
        if ($candidate.Rect) {
            $rectText = " rect=({0},{1},{2},{3})" -f $candidate.Rect.Left, $candidate.Rect.Top, $candidate.Rect.Right, $candidate.Rect.Bottom
        }
        Write-Output ("  value={0} source={1} handle={2} control_id={3} class={4}{5} text='{6}' caption='{7}'" -f `
            $candidate.Value,
            $candidate.SourceField,
            $candidate.Handle,
            $candidate.ControlId,
            $candidate.ClassName,
            $rectText,
            $candidate.Text,
            $candidate.Caption)
    }

    Write-Output ""
    Write-Output "Tip: if the precise plunger value is visible in the calibrate popup but missing above, the popup is probably not a Win32 child/control. In that case we need OCR on the popup rectangle."
}

function Get-WindowDetails {
    param([IntPtr]$Handle)

    return [pscustomobject]@{
        Handle = $Handle
        ProcessId = $(try {
            $pid = [uint32]0
            [void][StereoDriveWin32]::GetWindowThreadProcessId($Handle, [ref]$pid)
            [int]$pid
        } catch { $null })
        ControlId = [StereoDriveWin32]::GetDlgCtrlID($Handle)
        ClassName = Get-ClassName -Handle $Handle
        Caption = Get-WindowCaption -Handle $Handle
        Text = Get-ControlText -Handle $Handle
        Rect = Get-WindowRectObject -Handle $Handle
        Visible = [StereoDriveWin32]::IsWindowVisible($Handle)
        Enabled = [StereoDriveWin32]::IsWindowEnabled($Handle)
        Parent = [StereoDriveWin32]::GetParent($Handle)
        Style = ("0x{0:X}" -f [StereoDriveWin32]::GetWindowLongPtr($Handle, -16).ToInt64())
        ExStyle = ("0x{0:X}" -f [StereoDriveWin32]::GetWindowLongPtr($Handle, -20).ToInt64())
    }
}

function Find-ScaleControl {
    param(
        [int]$ProcessId,
        [IntPtr]$MainWindowHandle = [IntPtr]::Zero,
        [bool]$SearchAllVisibleWindows = $true
    )

    $windows = if ($SearchAllVisibleWindows) {
        @(Get-TopLevelWindows | Where-Object { $_.Visible -and $_.Rect -and $_.Rect.Width -gt 0 -and $_.Rect.Height -gt 0 })
    } else {
        @(Get-ProcessWindows -ProcessId $ProcessId)
    }

    $mainWindow = $null
    if ($MainWindowHandle -ne [IntPtr]::Zero) {
        $mainWindow = $windows | Where-Object { $_.Handle -eq $MainWindowHandle } | Select-Object -First 1
    }

    foreach ($window in $windows) {
        if (
            $window.ProcessId -ne $ProcessId -and
            $window.Caption -notmatch "Microdrive Calibrate Scale|Injectomate" -and
            $mainWindow -and
            (
                $window.Rect.Right -lt ($mainWindow.Rect.Left - 120) -or
                $window.Rect.Left -gt ($mainWindow.Rect.Right + 240) -or
                $window.Rect.Bottom -lt ($mainWindow.Rect.Top - 120) -or
                $window.Rect.Top -gt ($mainWindow.Rect.Bottom + 240)
            )
        ) {
            continue
        }
        $children = @(Get-ChildControls -ParentHandle $window.Handle)
        $match = $children | Where-Object { $_.ControlId -eq 3242 } | Select-Object -First 1
        if ($match) {
            return [pscustomobject]@{
                Window = $window
                Control = $match
                Children = $children
            }
        }
    }

    return $null
}

function Get-UiAutomationDetails {
    param([IntPtr]$Handle)

    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        $element = [System.Windows.Automation.AutomationElement]::FromHandle($Handle)
        if (-not $element) {
            return [pscustomobject]@{ Available = $false; Error = "AutomationElement.FromHandle returned null." }
        }

        $patterns = @()
        $value = $null
        $rangeValue = $null
        $text = $null

        try {
            $pattern = $element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
            if ($pattern) {
                $patterns += "ValuePattern"
                $value = $pattern.Current.Value
            }
        } catch {}

        try {
            $pattern = $element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern)
            if ($pattern) {
                $patterns += "RangeValuePattern"
                $rangeValue = $pattern.Current.Value
            }
        } catch {}

        try {
            $pattern = $element.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
            if ($pattern) {
                $patterns += "TextPattern"
                $text = $pattern.DocumentRange.GetText(1024)
            }
        } catch {}

        return [pscustomobject]@{
            Available = $true
            Name = $element.Current.Name
            AutomationId = $element.Current.AutomationId
            ClassName = $element.Current.ClassName
            ControlType = $element.Current.ControlType.ProgrammaticName
            LocalizedControlType = $element.Current.LocalizedControlType
            IsEnabled = $element.Current.IsEnabled
            IsKeyboardFocusable = $element.Current.IsKeyboardFocusable
            HasKeyboardFocus = $element.Current.HasKeyboardFocus
            Patterns = $patterns -join ", "
            ValuePatternValue = $value
            RangeValuePatternValue = $rangeValue
            TextPatternText = $text
        }
    } catch {
        return [pscustomobject]@{ Available = $false; Error = $_.Exception.Message }
    }
}

function Write-ScaleControlProbe {
    param(
        [int]$ProcessId,
        [IntPtr]$MainWindowHandle = [IntPtr]::Zero
    )

    $probe = Find-ScaleControl -ProcessId $ProcessId -MainWindowHandle $MainWindowHandle -SearchAllVisibleWindows $true
    if (-not $probe) {
        Write-Output "Scale control 3242 was not found. Open the Injectomate calibrate popup first."
        Write-Output ""
        Write-Output "Current StereoDrive process windows:"
        Get-ProcessWindows -ProcessId $ProcessId | ForEach-Object {
            Write-Output ("  handle={0} class={1} caption='{2}' rect=({3},{4},{5},{6})" -f $_.Handle, $_.ClassName, $_.Caption, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom)
        }
        return
    }

    Write-Output "Scale popup:"
    Get-WindowDetails -Handle $probe.Window.Handle | Format-List | Out-String -Width 240 | Write-Output
    Write-Output "Scale value control:"
    Get-WindowDetails -Handle $probe.Control.Handle | Format-List | Out-String -Width 240 | Write-Output
    Write-Output "UI Automation details:"
    Get-UiAutomationDetails -Handle $probe.Control.Handle | Format-List | Out-String -Width 240 | Write-Output
    Write-Output "Popup children:"
    $probe.Children |
        Sort-Object { $_.Rect.Top }, { $_.Rect.Left } |
        ForEach-Object {
            Write-Output ("  handle={0} control_id={1} class={2} rect=({3},{4},{5},{6}) text='{7}' caption='{8}'" -f $_.Handle, $_.ControlId, $_.ClassName, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom, $_.Text, $_.Caption)
        }
}

function Wait-ScaleControl {
    param(
        [int]$ProcessId,
        [IntPtr]$MainWindowHandle = [IntPtr]::Zero,
        [int]$TimeoutMilliseconds = 5000
    )

    $seen = @()
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    do {
        $probe = Find-ScaleControl -ProcessId $ProcessId -MainWindowHandle $MainWindowHandle -SearchAllVisibleWindows $true
        if ($probe) {
            return $probe
        }
        $seen = @(
            Get-TopLevelWindows |
                Where-Object { $_.Visible -and $_.Rect -and $_.Rect.Width -gt 0 -and $_.Rect.Height -gt 0 } |
                Select-Object -First 20
        )
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)

    return [pscustomobject]@{ Found = $false; SeenWindows = $seen }
}

function Read-ScaleViaPopupApi {
    param(
        [IntPtr]$MainWindowHandle,
        [int]$ProcessId,
        [string]$ClickMode
    )

    Open-InjectomateCalibrate -MainWindowHandle $MainWindowHandle -ClickMode $ClickMode
    $probe = Wait-ScaleControl -ProcessId $ProcessId -MainWindowHandle $MainWindowHandle -TimeoutMilliseconds 15000
    if ($probe.PSObject.Properties["Found"] -and $probe.Found -eq $false) {
        $windowSummary = @($probe.SeenWindows | ForEach-Object {
            "handle=$($_.Handle) pid=$($_.ProcessId) class=$($_.ClassName) caption='$($_.Caption)' rect=($($_.Rect.Left),$($_.Rect.Top),$($_.Rect.Right),$($_.Rect.Bottom))"
        }) -join "; "
        throw "Scale popup/control 3242 was not found after opening calibrate. Visible windows: $windowSummary"
    }

    $deadline = [DateTime]::UtcNow.AddMilliseconds(5000)
    $lastText = ""
    do {
        $lastText = Get-ControlText -Handle $probe.Control.Handle
        if ($lastText -match "^-?\d+(?:\.\d+)?$") {
            [void][StereoDriveWin32]::SendMessage($probe.Window.Handle, [StereoDriveWin32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
            return [pscustomobject]@{
                Value = [double]$lastText
                Text = $lastText
                PopupHandle = $probe.Window.Handle
                ControlHandle = $probe.Control.Handle
                ControlId = $probe.Control.ControlId
                ClassName = $probe.Control.ClassName
                Rect = $probe.Control.Rect
            }
        }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)

    [void][StereoDriveWin32]::SendMessage($probe.Window.Handle, [StereoDriveWin32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
    throw "Scale control was found but did not expose a numeric value. Last text='$lastText'."
}

function Find-BytePattern {
    param(
        [byte[]]$Buffer,
        [int]$Length,
        [byte[]]$Pattern
    )

    $matches = @()
    if ($Pattern.Count -eq 0 -or $Length -lt $Pattern.Count) {
        return $matches
    }
    $last = $Length - $Pattern.Count
    for ($i = 0; $i -le $last; $i++) {
        if ($Buffer[$i] -ne $Pattern[0]) {
            continue
        }
        $ok = $true
        for ($j = 1; $j -lt $Pattern.Count; $j++) {
            if ($Buffer[$i + $j] -ne $Pattern[$j]) {
                $ok = $false
                break
            }
        }
        if ($ok) {
            $matches += $i
        }
    }
    return $matches
}

function Get-MemorySearchPatterns {
    param([double]$NumericValue, [string]$OriginalText)

    $patterns = @()
    $patterns += [pscustomobject]@{ Name = "double"; Bytes = [BitConverter]::GetBytes([double]$NumericValue) }
    $patterns += [pscustomobject]@{ Name = "float"; Bytes = [BitConverter]::GetBytes([single]$NumericValue) }
    foreach ($scale in @(1, 10, 100, 1000, 10000)) {
        $scaled = [int64][Math]::Round($NumericValue * $scale)
        if ($scaled -ge [int32]::MinValue -and $scaled -le [int32]::MaxValue) {
            $patterns += [pscustomobject]@{ Name = "int32_x$scale"; Bytes = [BitConverter]::GetBytes([int32]$scaled) }
        }
        $patterns += [pscustomobject]@{ Name = "int64_x$scale"; Bytes = [BitConverter]::GetBytes([int64]$scaled) }
    }
    if (-not [string]::IsNullOrWhiteSpace($OriginalText)) {
        $patterns += [pscustomobject]@{ Name = "ascii"; Bytes = [System.Text.Encoding]::ASCII.GetBytes($OriginalText) }
        $patterns += [pscustomobject]@{ Name = "utf16le"; Bytes = [System.Text.Encoding]::Unicode.GetBytes($OriginalText) }
    }
    return $patterns
}

function Scan-ProcessMemoryForValue {
    param(
        [int]$ProcessId,
        [string]$Needle,
        [int]$MaxMatches = 80
    )

    if ([string]::IsNullOrWhiteSpace($Needle)) {
        throw "scan-process-memory-value requires -Value, for example -Value 3071.057."
    }
    $numericValue = [double]$Needle
    $patterns = @(Get-MemorySearchPatterns -NumericValue $numericValue -OriginalText $Needle)
    $access = [StereoDriveWin32]::PROCESS_QUERY_INFORMATION -bor [StereoDriveWin32]::PROCESS_VM_READ
    $processHandle = [StereoDriveWin32]::OpenProcess($access, $false, $ProcessId)
    if ($processHandle -eq [IntPtr]::Zero) {
        throw "OpenProcess failed for PID $ProcessId. Try running PowerShell as administrator."
    }

    $matches = @()
    $regionCount = 0
    $readableRegionCount = 0
    $bytesScanned = [uint64]0
    $address = [uint64]0
    $maxAddress = if ([IntPtr]::Size -eq 8) { [uint64]0x7FFFFFFF0000 } else { [uint64]0x7FFF0000 }
    $chunkSize = 1048576
    $startTime = [DateTime]::UtcNow
    $lastProgress = [DateTime]::UtcNow.AddSeconds(-10)
    Write-Host ("Starting memory scan for {0}. Patterns: {1}" -f $Needle, (($patterns | ForEach-Object { $_.Name }) -join ", "))
    try {
        while ($address -lt $maxAddress -and $matches.Count -lt $MaxMatches) {
            $mbi = New-Object StereoDriveWin32+MEMORY_BASIC_INFORMATION
            $result = [StereoDriveWin32]::VirtualQueryEx(
                $processHandle,
                [IntPtr]::new([int64]$address),
                [ref]$mbi,
                [UIntPtr][uint64][System.Runtime.InteropServices.Marshal]::SizeOf($mbi)
            )
            if ($result -eq [UIntPtr]::Zero) {
                break
            }

            $regionCount++
            $base = [uint64]$mbi.BaseAddress.ToInt64()
            $size = [uint64]$mbi.RegionSize.ToUInt64()
            $next = $base + [Math]::Max($size, [uint64]4096)
            $isReadable = (
                $mbi.State -eq [StereoDriveWin32]::MEM_COMMIT -and
                (($mbi.Protect -band [StereoDriveWin32]::PAGE_NOACCESS) -eq 0) -and
                (($mbi.Protect -band [StereoDriveWin32]::PAGE_GUARD) -eq 0)
            )
            if ($isReadable) {
                $readableRegionCount++
                $offset = [uint64]0
                while ($offset -lt $size -and $matches.Count -lt $MaxMatches) {
                    $toRead = [int][Math]::Min([uint64]$chunkSize, $size - $offset)
                    $buffer = New-Object byte[] $toRead
                    $bytesRead = [UIntPtr]::Zero
                    $ok = [StereoDriveWin32]::ReadProcessMemory(
                        $processHandle,
                        [IntPtr]::new([int64]($base + $offset)),
                        $buffer,
                        [UIntPtr][uint64]$toRead,
                        [ref]$bytesRead
                    )
                    $actual = [int]$bytesRead.ToUInt64()
                    if ($ok -and $actual -gt 0) {
                        $bytesScanned += [uint64]$actual
                        if ((([DateTime]::UtcNow - $lastProgress).TotalSeconds -ge 1.0)) {
                            $elapsed = [Math]::Max(0.1, ([DateTime]::UtcNow - $startTime).TotalSeconds)
                            $mb = [Math]::Round($bytesScanned / 1MB, 1)
                            $rate = [Math]::Round(($bytesScanned / 1MB) / $elapsed, 1)
                            $status = "{0} MB scanned, {1} regions, {2} matches, {3} MB/s" -f $mb, $readableRegionCount, $matches.Count, $rate
                            Write-Progress -Activity "Scanning StereoDrive process memory" -Status $status -PercentComplete -1
                            Write-Host ("progress: {0}" -f $status)
                            $lastProgress = [DateTime]::UtcNow
                        }
                        foreach ($pattern in $patterns) {
                            foreach ($hit in @(Find-BytePattern -Buffer $buffer -Length $actual -Pattern $pattern.Bytes)) {
                                $matches += [pscustomobject]@{
                                    Address = ("0x{0:X}" -f ($base + $offset + [uint64]$hit))
                                    AddressValue = ($base + $offset + [uint64]$hit)
                                    Pattern = $pattern.Name
                                    RegionBase = ("0x{0:X}" -f $base)
                                    RegionSize = $size
                                    Protect = ("0x{0:X}" -f $mbi.Protect)
                                    Type = ("0x{0:X}" -f $mbi.Type)
                                }
                                if ($matches.Count -ge $MaxMatches) {
                                    break
                                }
                            }
                            if ($matches.Count -ge $MaxMatches) {
                                break
                            }
                        }
                    }
                    $offset += [uint64]$toRead
                }
            }
            if ($next -le $address) {
                break
            }
            $address = $next
        }
    } finally {
        Write-Progress -Activity "Scanning StereoDrive process memory" -Completed
        [void][StereoDriveWin32]::CloseHandle($processHandle)
    }

    return [pscustomobject]@{
        ProcessId = $ProcessId
        Needle = $Needle
        Patterns = ($patterns | ForEach-Object { $_.Name }) -join ", "
        RegionsVisited = $regionCount
        ReadableRegions = $readableRegionCount
        BytesScanned = $bytesScanned
        MatchCount = $matches.Count
        Matches = $matches
    }
}

function Read-ProcessBytes {
    param(
        [IntPtr]$ProcessHandle,
        [uint64]$Address,
        [int]$Length
    )

    $buffer = New-Object byte[] $Length
    $bytesRead = [UIntPtr]::Zero
    $ok = [StereoDriveWin32]::ReadProcessMemory(
        $ProcessHandle,
        [IntPtr]::new([int64]$Address),
        $buffer,
        [UIntPtr][uint64]$Length,
        [ref]$bytesRead
    )
    if (-not $ok -or $bytesRead.ToUInt64() -lt [uint64]$Length) {
        return $null
    }
    return $buffer
}

function Test-ByteArrayEqual {
    param([byte[]]$Left, [byte[]]$Right)

    if ($Left.Count -ne $Right.Count) {
        return $false
    }
    for ($i = 0; $i -lt $Left.Count; $i++) {
        if ($Left[$i] -ne $Right[$i]) {
            return $false
        }
    }
    return $true
}

function Convert-HexAddressToUInt64 {
    param([string]$Address)

    $text = $Address.Trim()
    if ($text.StartsWith("0x", [System.StringComparison]::OrdinalIgnoreCase)) {
        $text = $text.Substring(2)
    }
    return [Convert]::ToUInt64($text, 16)
}

function Refine-ProcessMemoryCandidates {
    param(
        [int]$ProcessId,
        [string]$Needle,
        [string]$CandidatePath
    )

    if ([string]::IsNullOrWhiteSpace($Needle)) {
        throw "refine-process-memory-value requires -Value, for example -Value 3072.123."
    }
    if (-not (Test-Path -LiteralPath $CandidatePath)) {
        throw "No saved memory candidate file exists at $CandidatePath. Run -Action scan-process-memory-value first."
    }

    $previous = Get-Content -LiteralPath $CandidatePath -Raw | ConvertFrom-Json
    $previousMatches = @($previous.Matches)
    if ($previousMatches.Count -eq 0) {
        throw "Saved candidate file has no matches to refine."
    }

    $patterns = @(Get-MemorySearchPatterns -NumericValue ([double]$Needle) -OriginalText $Needle)
    $patternByName = @{}
    foreach ($pattern in $patterns) {
        $patternByName[$pattern.Name] = $pattern
    }

    $access = [StereoDriveWin32]::PROCESS_QUERY_INFORMATION -bor [StereoDriveWin32]::PROCESS_VM_READ
    $processHandle = [StereoDriveWin32]::OpenProcess($access, $false, $ProcessId)
    if ($processHandle -eq [IntPtr]::Zero) {
        throw "OpenProcess failed for PID $ProcessId. Try running PowerShell as administrator."
    }

    $kept = @()
    $checked = 0
    try {
        foreach ($candidate in $previousMatches) {
            $checked++
            $address = if ($candidate.PSObject.Properties["AddressValue"] -and $null -ne $candidate.AddressValue) {
                [uint64]$candidate.AddressValue
            } else {
                Convert-HexAddressToUInt64 -Address ([string]$candidate.Address)
            }

            $patternsToTry = @()
            if ($patternByName.ContainsKey([string]$candidate.Pattern)) {
                $patternsToTry = @($patternByName[[string]$candidate.Pattern])
            } else {
                $patternsToTry = $patterns
            }

            foreach ($pattern in $patternsToTry) {
                $bytes = Read-ProcessBytes -ProcessHandle $processHandle -Address $address -Length $pattern.Bytes.Count
                if ($bytes -and (Test-ByteArrayEqual -Left $bytes -Right $pattern.Bytes)) {
                    $kept += [pscustomobject]@{
                        Address = ("0x{0:X}" -f $address)
                        AddressValue = $address
                        Pattern = $pattern.Name
                        PreviousPattern = $candidate.Pattern
                        MatchedPattern = $pattern.Name
                        RegionBase = if ($candidate.PSObject.Properties["RegionBase"]) { $candidate.RegionBase } elseif ($candidate.PSObject.Properties["PreviousRegionBase"]) { $candidate.PreviousRegionBase } else { $null }
                        Protect = if ($candidate.PSObject.Properties["Protect"]) { $candidate.Protect } elseif ($candidate.PSObject.Properties["PreviousProtect"]) { $candidate.PreviousProtect } else { $null }
                        Type = if ($candidate.PSObject.Properties["Type"]) { $candidate.Type } elseif ($candidate.PSObject.Properties["PreviousType"]) { $candidate.PreviousType } else { $null }
                    }
                    break
                }
            }
        }
    } finally {
        [void][StereoDriveWin32]::CloseHandle($processHandle)
    }

    return [pscustomobject]@{
        ProcessId = $ProcessId
        PreviousNeedle = $previous.Needle
        Needle = $Needle
        CandidatePath = $CandidatePath
        PreviousMatchCount = $previousMatches.Count
        CheckedCount = $checked
        MatchCount = $kept.Count
        Matches = $kept
    }
}

function Find-PlungerGaugeControl {
    param([IntPtr]$MainWindowHandle)

    $controls = @(Get-ChildControls -ParentHandle $MainWindowHandle)
    $exact = $controls |
        Where-Object { $_.ClassName -eq "AfxWnd140" -and ($_.Text -eq "MMCDepth" -or $_.Caption -eq "MMCDepth") } |
        Sort-Object { $_.Rect.Width * $_.Rect.Height } -Descending |
        Select-Object -First 1
    if ($exact) {
        return [pscustomobject]@{ Control = $exact; AllControls = $controls }
    }

    $fallback = $controls |
        Where-Object { $_.ClassName -eq "AfxWnd140" -and $_.Rect -and $_.Rect.Height -gt 200 -and $_.Rect.Width -gt 20 } |
        Sort-Object { $_.Rect.Left } -Descending |
        Select-Object -First 1
    if ($fallback) {
        return [pscustomobject]@{ Control = $fallback; AllControls = $controls }
    }

    return [pscustomobject]@{ Control = $null; AllControls = $controls }
}

function Get-ParentChain {
    param([IntPtr]$Handle)

    $chain = @()
    $current = $Handle
    for ($i = 0; $i -lt 8; $i++) {
        $parent = [StereoDriveWin32]::GetParent($current)
        if ($parent -eq [IntPtr]::Zero) {
            break
        }
        $chain += Get-WindowDetails -Handle $parent
        $current = $parent
    }
    return $chain
}

function Write-PlungerGaugeProbe {
    param([IntPtr]$MainWindowHandle)

    $probe = Find-PlungerGaugeControl -MainWindowHandle $MainWindowHandle
    if (-not $probe.Control) {
        Write-Output "MMCDepth plunger gauge control was not found."
        Write-Output ""
        Write-Output "AfxWnd140 candidates:"
        $probe.AllControls |
            Where-Object { $_.ClassName -eq "AfxWnd140" } |
            ForEach-Object {
                Write-Output ("  handle={0} control_id={1} text='{2}' caption='{3}' rect=({4},{5},{6},{7})" -f $_.Handle, $_.ControlId, $_.Text, $_.Caption, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom)
            }
        return
    }

    $control = $probe.Control
    Write-Output "Plunger gauge control:"
    Get-WindowDetails -Handle $control.Handle | Format-List | Out-String -Width 240 | Write-Output
    Write-Output "UI Automation details:"
    Get-UiAutomationDetails -Handle $control.Handle | Format-List | Out-String -Width 240 | Write-Output

    Write-Output "Parent chain:"
    Get-ParentChain -Handle $control.Handle | ForEach-Object {
        Write-Output ("  handle={0} control_id={1} class={2} text='{3}' caption='{4}' rect=({5},{6},{7},{8})" -f $_.Handle, $_.ControlId, $_.ClassName, $_.Text, $_.Caption, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom)
    }

    Write-Output "Gauge child windows:"
    $children = @(Get-ChildControls -ParentHandle $control.Handle)
    if ($children.Count -eq 0) {
        Write-Output "  none"
    } else {
        $children | ForEach-Object {
            Write-Output ("  handle={0} control_id={1} class={2} text='{3}' caption='{4}' rect=({5},{6},{7},{8})" -f $_.Handle, $_.ControlId, $_.ClassName, $_.Text, $_.Caption, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom)
        }
    }

    Write-Output "Nearby sibling controls:"
    $left = $control.Rect.Left - 40
    $right = $control.Rect.Right + 40
    $top = $control.Rect.Top - 40
    $bottom = $control.Rect.Bottom + 40
    $probe.AllControls |
        Where-Object {
            $_.Handle -ne $control.Handle -and
            $_.Rect -and
            $_.Rect.Right -ge $left -and
            $_.Rect.Left -le $right -and
            $_.Rect.Bottom -ge $top -and
            $_.Rect.Top -le $bottom
        } |
        Sort-Object { $_.Rect.Top }, { $_.Rect.Left } |
        ForEach-Object {
            Write-Output ("  handle={0} control_id={1} class={2} text='{3}' caption='{4}' rect=({5},{6},{7},{8})" -f $_.Handle, $_.ControlId, $_.ClassName, $_.Text, $_.Caption, $_.Rect.Left, $_.Rect.Top, $_.Rect.Right, $_.Rect.Bottom)
        }
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

function Show-Injectomate {
    param([IntPtr]$MainWindowHandle)

    $map = Get-ControlMap -MainWindowHandle $MainWindowHandle
    if ($map.ContainsKey("10030") -or $map.ContainsKey("10001")) {
        return
    }
    [void][StereoDriveWin32]::SendMessage($MainWindowHandle, [StereoDriveWin32]::WM_COMMAND, [IntPtr]::new(32815), [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
}

function Open-InjectomateCalibrate {
    param(
        [IntPtr]$MainWindowHandle,
        [string]$ClickMode
    )

    Show-Injectomate -MainWindowHandle $MainWindowHandle
    $map = Get-ControlMap -MainWindowHandle $MainWindowHandle
    Invoke-ButtonClick -ControlMap $map -ControlId 10030 -MainWindowHandle $MainWindowHandle -ClickMode $ClickMode
    Start-Sleep -Milliseconds 500
    $probe = Wait-ScaleControl -ProcessId $mainProcessId -MainWindowHandle $MainWindowHandle -TimeoutMilliseconds 1000
    if (-not ($probe.PSObject.Properties["Found"] -and $probe.Found -eq $false)) {
        return
    }
    Invoke-ControlMouseClick -ControlMap $map -ControlId 10030 -MainWindowHandle $MainWindowHandle
}

function Get-InjectomateMap {
    param([IntPtr]$MainWindowHandle)

    $controls = Get-ChildControls -ParentHandle $MainWindowHandle
    return $controls |
        Where-Object {
            $_.ClassName -eq "Button" -or
            $_.ClassName -eq "ComboBox" -or
            $_.ClassName -eq "Edit" -or
            $_.Caption -match "Inject|Syringe|nl|ul|Start|Stop|Volume|Rate|Speed" -or
            $_.Text -match "Inject|Syringe|nl|ul|Start|Stop|Volume|Rate|Speed"
        } |
        Sort-Object ControlId |
        ForEach-Object {
            $items = @()
            if ($_.ClassName -eq "ComboBox") {
                try {
                    $items = @(Get-ComboItems -Handle $_.Handle)
                } catch {
                    $items = @()
                }
            }
            [pscustomobject]@{
                ControlId = $_.ControlId
                ClassName = $_.ClassName
                Caption = $_.Caption
                Text = $_.Text
                ComboItems = if ($items.Count -gt 0) { $items -join ", " } else { "" }
            }
        }
}

function Find-InjectomateValueControls {
    param(
        [IntPtr]$MainWindowHandle,
        [string]$WantedValue
    )

    $controls = Get-ChildControls -ParentHandle $MainWindowHandle
    $injectomateControls = @($controls | Where-Object { $_.ControlId -ge 10000 })
    if ($injectomateControls.Count -eq 0) {
        return @()
    }

    $numericPattern = "-?\d+(\.\d+)?"
    $numericControls = @(
        $controls |
            Where-Object {
                $_.Text -match $numericPattern -or $_.Caption -match $numericPattern
            }
    )

    $exactControls = @()
    if ($WantedValue) {
        $wantedPattern = [regex]::Escape($WantedValue)
        $exactControls = @(
            $numericControls |
                Where-Object {
                    $_.Text -match $wantedPattern -or $_.Caption -match $wantedPattern
                }
        )
    }

    $outputControls = if ($exactControls.Count -gt 0) { $exactControls } else { $numericControls }
    return $outputControls |
        Sort-Object ControlId, ClassName |
        ForEach-Object {
            $textMatches = if ($WantedValue) { $_.Text -match ([regex]::Escape($WantedValue)) } else { $false }
            $captionMatches = if ($WantedValue) { $_.Caption -match ([regex]::Escape($WantedValue)) } else { $false }
            [pscustomobject]@{
                Handle = $_.Handle
                ControlId = $_.ControlId
                ClassName = $_.ClassName
                Caption = $_.Caption
                Text = $_.Text
                ExactMatch = ($textMatches -or $captionMatches)
            }
        }
}

function Get-InjectomateRegionControls {
    param([IntPtr]$MainWindowHandle)

    $controls = Get-ChildControls -ParentHandle $MainWindowHandle
    $injectomateControls = @($controls | Where-Object { $_.ControlId -ge 10000 })
    if ($injectomateControls.Count -eq 0) {
        return @()
    }

    $left = ($injectomateControls | ForEach-Object { $_.Rect.Left } | Measure-Object -Minimum).Minimum
    $right = ($injectomateControls | ForEach-Object { $_.Rect.Right } | Measure-Object -Maximum).Maximum
    $top = ($injectomateControls | ForEach-Object { $_.Rect.Top } | Measure-Object -Minimum).Minimum
    $bottom = ($injectomateControls | ForEach-Object { $_.Rect.Bottom } | Measure-Object -Maximum).Maximum

    return $controls |
        Where-Object {
            $_.Rect -and
            $_.Rect.Right -ge ($left - 30) -and
            $_.Rect.Left -le ($right + 30) -and
            $_.Rect.Bottom -ge ($top - 30) -and
            $_.Rect.Top -le ($bottom + 30)
        } |
        Sort-Object { $_.Rect.Top }, { $_.Rect.Left } |
        ForEach-Object {
            [pscustomobject]@{
                Handle = $_.Handle
                ControlId = $_.ControlId
                ClassName = $_.ClassName
                Caption = $_.Caption
                Text = $_.Text
                Left = $_.Rect.Left
                Top = $_.Rect.Top
                Width = $_.Rect.Width
                Height = $_.Rect.Height
            }
        }
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
        [hashtable]$ControlMap,
        [bool]$EnsureVisible = $true
    )

    if ($EnsureVisible) {
        $ControlMap = Ensure-ReferencePanel -MainWindowHandle $MainWindowHandle -ProcessId $ProcessId -ControlMap $ControlMap
    }

    $status = ""
    if ($ControlMap.ContainsKey("1097")) {
        $status = Get-ControlText -Handle $ControlMap["1097"].Handle
    }

    return [pscustomobject]@{
        ReferenceSelector = if ($ControlMap.ContainsKey("1387")) { Get-ControlText -Handle $ControlMap["1387"].Handle } else { "" }
        ReferenceStatus = $status
        ReferencePanelVisible = $ControlMap.ContainsKey("1095")
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
    "inject" = 10018
    "fill" = 10032
    "open-injectomate-calibrate" = 10030
    "close-injectomate" = 10031
}

$comboIds = @{
    "set-step-ap" = 1132
    "set-step-ml" = 1133
    "set-step-dv" = 1134
    "set-injection-volume" = 10001
    "set-syringe-type" = 10006
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

if ($Action -eq "show-injectomate") {
    Show-Injectomate -MainWindowHandle $mainHandle
    Get-InjectomateMap -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "hide-injectomate") {
    Show-Injectomate -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "injectomate-map") {
    Get-InjectomateMap -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "scan-injectomate-value") {
    Show-Injectomate -MainWindowHandle $mainHandle
    Find-InjectomateValueControls -MainWindowHandle $mainHandle -WantedValue $Value
    return
}

if ($Action -eq "scan-injectomate-region") {
    Show-Injectomate -MainWindowHandle $mainHandle
    Get-InjectomateRegionControls -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "scan-open-windows") {
    if ($DelaySeconds -gt 0) {
        Start-Sleep -Seconds $DelaySeconds
    }
    $includeNearby = ($Value -eq "nearby" -or $Value -eq "popups")
    $snapshot = Get-WindowTreeSnapshot -MainWindowHandle $mainHandle -ProcessId $mainProcessId -IncludeNearbyWindows $includeNearby
    if ($Value -eq "json") {
        $snapshot
    } else {
        Write-WindowTreeReport -Snapshot $snapshot
    }
    return
}

if ($Action -eq "probe-injectomate-calibrate") {
    $before = Get-WindowTreeSnapshot -MainWindowHandle $mainHandle -ProcessId $mainProcessId
    Open-InjectomateCalibrate -MainWindowHandle $mainHandle -ClickMode $ClickMode
    $after = Get-WindowTreeSnapshot -MainWindowHandle $mainHandle -ProcessId $mainProcessId
    [pscustomobject]@{
        ClickMode = $ClickMode
        Before = $before
        After = $after
    }
    return
}

if ($Action -eq "probe-scale-control") {
    Write-ScaleControlProbe -ProcessId $mainProcessId -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "probe-plunger-gauge-control") {
    Write-PlungerGaugeProbe -MainWindowHandle $mainHandle
    return
}

if ($Action -eq "read-scale-popup-api") {
    Read-ScaleViaPopupApi -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ClickMode $ClickMode
    return
}

if ($Action -eq "scan-process-memory-value") {
    $result = Scan-ProcessMemoryForValue -ProcessId $mainProcessId -Needle $Value
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $MemoryCandidatesPath -Encoding UTF8
    Write-Host "Saved memory candidates to $MemoryCandidatesPath"
    $result
    return
}

if ($Action -eq "refine-process-memory-value") {
    $result = Refine-ProcessMemoryCandidates -ProcessId $mainProcessId -Needle $Value -CandidatePath $MemoryCandidatesPath
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $MemoryCandidatesPath -Encoding UTF8
    Write-Host "Updated memory candidates at $MemoryCandidatesPath"
    $result
    return
}

if ($Action -eq "clear-process-memory-candidates") {
    if (Test-Path -LiteralPath $MemoryCandidatesPath) {
        Remove-Item -LiteralPath $MemoryCandidatesPath -Force
        Write-Output "Deleted $MemoryCandidatesPath"
    } else {
        Write-Output "No memory candidate file exists at $MemoryCandidatesPath"
    }
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

if ($Action -eq "probe-command") {
    if ($CommandId -le 0) {
        throw "Action 'probe-command' requires -CommandId."
    }

    $before = Get-WindowSnapshot -MainWindowHandle $mainHandle -ProcessId $mainProcessId
    Invoke-DirectCommand -MainWindowHandle $mainHandle -CommandId $CommandId
    $after = Get-WindowSnapshot -MainWindowHandle $mainHandle -ProcessId $mainProcessId

    [pscustomobject]@{
        CommandId = $CommandId
        Before = $before
        After = $after
    }
    return
}

if ($Action -eq "test-hidden-reference-bregma") {
    Invoke-DirectCommand -MainWindowHandle $mainHandle -CommandId 1095
    $postMap = Get-ControlMap -MainWindowHandle $mainHandle
    [pscustomobject]@{
        Coords = Get-Coords -ControlMap $postMap
        Reference = Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap $postMap -EnsureVisible $false
    }
    return
}

if ($Action -eq "test-hidden-drill-to-bregma") {
    Invoke-DirectCommand -MainWindowHandle $mainHandle -CommandId 1071
    $postMap = Get-ControlMap -MainWindowHandle $mainHandle
    [pscustomobject]@{
        Coords = Get-Coords -ControlMap $postMap
        Reference = Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap $postMap -EnsureVisible $false
    }
    return
}

if ($comboIds.ContainsKey($Action)) {
    if (-not $Value) {
        throw "Action '$Action' requires -Value, for example -Value '0.1 mm'."
    }
    if ($Action -like "set-injection-*" -or $Action -eq "set-syringe-type") {
        Show-Injectomate -MainWindowHandle $mainHandle
        $controlMap = Get-ControlMap -MainWindowHandle $mainHandle
    }
    Set-ComboByText -ControlMap $controlMap -ControlId $comboIds[$Action] -WantedText $Value
    if ($Action -like "set-injection-*" -or $Action -eq "set-syringe-type") {
        Get-InjectomateMap -MainWindowHandle $mainHandle
    } else {
        Get-Coords -ControlMap $controlMap
    }
    return
}

if ($buttonIds.ContainsKey($Action)) {
    if ($Action -like "set-reference-*" -or $Action -like "set-*-to-bregma" -or $Action -eq "close-reference-panel") {
        $controlMap = Ensure-ReferencePanel -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap $controlMap
    }
    if ($Action -eq "inject" -or $Action -eq "fill" -or $Action -eq "close-injectomate") {
        Show-Injectomate -MainWindowHandle $mainHandle
        $controlMap = Get-ControlMap -MainWindowHandle $mainHandle
    }
    Invoke-ButtonClick -ControlMap $controlMap -ControlId $buttonIds[$Action] -MainWindowHandle $mainHandle -ClickMode $ClickMode
    if ($Action -eq "inject" -or $Action -eq "fill" -or $Action -eq "close-injectomate") {
        Get-InjectomateMap -MainWindowHandle $mainHandle
        return
    }
    [pscustomobject]@{
        Coords = Get-Coords -ControlMap $controlMap
        Reference = Get-ReferenceStatus -MainWindowHandle $mainHandle -ProcessId $mainProcessId -ControlMap (Get-ControlMap -MainWindowHandle $mainHandle)
    }
    return
}

throw "Unsupported action '$Action'."
