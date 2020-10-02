[CmdletBinding()] Param([parameter(mandatory=$false)] [string]$data_filename)

$showDebugMessages = $true

cls

if ($data_filename -eq '')
{
    $data_filename = "$PSScriptRoot\files.json"
}
if (!(test-path -Path $data_filename -PathType leaf))
{
    write-host "Input file not found: $data_filename"
    exit
}

. "$PSScriptRoot\VirtualDesktop.ps1"

Set-StrictMode -Version Latest

$winApi = @"

public enum SWP_FLAGS : uint {
    NOSIZE = 0x0001,
    NOMOVE = 0x0002,
    NOZORDER = 0x0004,
    NOREDRAW = 0x0008,
    NOACTIVATE = 0x0010,
    FRAMECHANGED = 0x0020,  /* The frame changed: send WM_NCCALCSIZE */
    SHOWWINDOW = 0x0040,
    HIDEWINDOW = 0x0080,
    NOCOPYBITS = 0x0100,
    NOOWNERZORDER = 0x0200,  /* Don't do owner Z ordering */
    NOSENDCHANGING = 0x0400,  /* Don't send WM_WINDOWPOSCHANGING */
}

[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, SWP_FLAGS uFlags);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); 
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
[DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

public struct WINDOWPLACEMENT
{
    public int length;
    public int flags;
    public int showCmd;
    public POINT ptMinPosition;
    public POINT ptMaxPosition;
    public RECT rcNormalPosition;
}

public struct POINT
{
    public int X;
    public int Y;
}
public struct RECT
{
    public int left;
    public int top;
    public int right;
    public int bottom;
}

public static void GetWindowPos(IntPtr hWnd, out int x, out int y)
{
    RECT rect;
    GetWindowRect(hWnd, out rect);
    x = rect.left;
    y = rect.top;
}
public static void SetWindowPos(IntPtr hWnd, int x, int y, int width, int height)
{
    SWP_FLAGS flags = SWP_FLAGS.NOZORDER | SWP_FLAGS.NOACTIVATE;
    if (x == -1 || y == -1)
        flags |= SWP_FLAGS.NOMOVE;
    if (width == -1 || height == -1)
        flags |= SWP_FLAGS.NOSIZE;
    SetWindowPos(hWnd, new IntPtr(0), x, y, width, height, flags);
}
public static void MaximiseWindow(IntPtr hWnd)
{
    ShowWindow(hWnd, /*SW_MAXIMIZE*/ 3);
}
public static void MinimiseWindow(IntPtr hWnd)
{
    ShowWindow(hWnd, /*SW_MINIMIZE*/ 6);
}
public static void RestoreWindow(IntPtr hWnd)
{
    ShowWindow(hWnd, /*SW_RESTORE*/ 9);
}
public static void SetWindowOnTop(IntPtr hWnd)
{
    SetWindowPos(hWnd, new IntPtr( /*HWND_TOPMOST*/ -1), 0, 0, 0, 0, SWP_FLAGS.NOMOVE | SWP_FLAGS.NOSIZE);
}
public static void SetWindowNotOnTop(IntPtr hWnd)
{
    SetWindowPos(hWnd, new IntPtr( /*HWND_NOTTOPMOST*/ -2), 0, 0, 0, 0, SWP_FLAGS.NOMOVE | SWP_FLAGS.NOSIZE);
}
public static bool IsWindowShowing(IntPtr hWnd)
{
    return IsWindowVisible(hWnd);
}
public static int GetWindowState(IntPtr hWnd)
{
    WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
    placement.length = Marshal.SizeOf(placement);
    GetWindowPlacement(hWnd, ref placement);
    return placement.showCmd;
}

public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

public static System.Collections.Generic.IEnumerable<IntPtr> FindWindows(EnumWindowsProc filter)
{
  IntPtr found = IntPtr.Zero;
  System.Collections.Generic.List<IntPtr> windows = new System.Collections.Generic.List<IntPtr>();

  EnumWindows(delegate(IntPtr wnd, IntPtr param)
  {
      if (filter(wnd, param))
      {
          // only add the windows that pass the filter
          windows.Add(wnd);
      }

      // but return true here so that we iterate all windows
      return true;
  }, IntPtr.Zero);

  return windows;
}

public static System.Collections.Generic.IEnumerable<IntPtr> FindWindowsWithClassname(string name)
{
    return FindWindows(delegate(IntPtr wnd, IntPtr param)
    {
        var builder = new System.Text.StringBuilder(200);
        GetClassName(wnd, builder, builder.Capacity);
        return builder.ToString() == name;
    });
}

"@

Add-Type -MemberDefinition $winApi -Namespace Win -Name API
Add-Type -AssemblyName System.Windows.Forms

Function ReadData($file)
{
    $objects = New-Object System.Collections.ArrayList

    $data = $(get-content $data_filename | ConvertFrom-Json)

    $maxDesktop = -1

    foreach ($item in $data)
    {
    	$object = [PSCustomObject] @{ process = ''; arguments = @(); desktop = -1; monitor = -1; x = -1; y = -1; width = -1; height = -1; state = ''; hasSplash = $false; showOnAllDesktops = $false; skipIfAlreadyRunning = $false; hwnd = 0; done = $false; proc = 0; index = -1; processName = ''; launchOnDesktop = $false; waitForClass = ''; waitForProcessToClose = $false; retry=0; windowClassList = 0 }

        $item.PSOBject.Properties | ForEach-Object {
            $value = $_.Value
            switch ($_.Name)
            {
                'process' { $object.process = $value }
                'processName' { $object.processName = $value }
                'arguments' { $object.arguments += $value }
                'desktop' { $object.desktop = [int]$value }
                'monitor' { $object.monitor = [int]$value }
                'x' { $object.x = [int]$value }
                'y' { $object.y = [int]$value }
                'width' { $object.width = [int]$value }
                'height' { $object.height = [int]$value }
                'state' { $object.state = $value }
                'hassplash' { $object.hassplash = [System.Convert]::ToBoolean($value) }
                'showOnAllDesktops' { $object.showOnAllDesktops = [System.Convert]::ToBoolean($value) }
                'skipIfAlreadyRunning' { $object.skipIfAlreadyRunning = [System.Convert]::ToBoolean($value) }
                'launchOnDesktop' { $object.launchOnDesktop = [System.Convert]::ToBoolean($value) }
                'waitForClass' { $object.waitForClass = $value }
                'waitForProcessToClose' { $object.waitForProcessToClose = [System.Convert]::ToBoolean($value) }
                default { write-host "Unknown setting: $_ = $value" }
            }
        }
        
        if ($object.process -ne '')
        {
            if ($object.desktop -ne -1)
            {
                if ($object.desktop -eq 0)
                {
                    Write-host "Desktop numbers start from 1 not 0!"
                    $object.desktop = -1
                }
                else
                {
                    if ($object.desktop -gt $maxDesktop)
                    {
                        $maxDesktop = $object.desktop
                    }
                }
            }
            if ($object.monitor -eq 0)
            {
                Write-host "Monitor numbers start from 1 not 0!"
                $object.monitor = -1
            }

            $object.index = $objects.Add($object) + 1
        }
        else
        {
            write-host 'Ignoring item with no process name!'
        }
    }

    if ($maxDesktop -ne -1)
    {
        $NumDesktops = $(Get-DesktopCount)
        while ($NumDesktops -lt $maxDesktop)
        {
            New-Desktop | Out-Null
            $NumDesktops++;
        }
    }

    return $objects
}

function DebugMessage($msg)
{
    if ($showDebugMessages)
    {
        write-host -ForegroundColor DarkGray "<< $msg >>"
    }
}

function OutputObjectMessage($object, $msg)
{
    write-host -NoNewline -ForegroundColor Red "$($object.index): $($object.processName) "
    if ($object.hwnd)
    {
        write-host "$msg ($($object.hwnd))"
    }
    else
    {
        write-host $msg
    }
}

function WaitForWindowToShow($hwnd)
{
    $trys = 0
    $amount = 100
    while (![Win.API]::IsWindowShowing($hwnd))
    {
        start-sleep -Milliseconds $amount
        $trys++
        $amount += 100
        if ($trys -eq 20)
        {
            write-host "$hwnd not yet showing (timed out)"
            return
        }
    }

    DebugMessage "$hwnd Showing"
}

function WaitForWindowToStopShowing($hwnd)
{
    $trys = 0
    while ([Win.API]::IsWindowShowing($hwnd))
    {
        start-sleep -Milliseconds 100
        $trys++
        if ($trys -eq 100)
        {
            write-host "$hwnd not closing (timed out)"
            return
        }
    }

    DebugMessage "$hwnd Hidden"
}

function FindNewWindow($object, [ref]$retHwnd)
{
    $newWidowClassList = [Win.API]::FindWindowsWithClassname($object.waitForClass)

    $newHwnd = $newWidowClassList | Where-Object { $_ -notin $object.windowClassList }

    if ($newHwnd)
    {
        DebugMessage $newWidowClassList

        if ($newHwnd -is [Array])
        {
            write-host "Taking first!"
            $newHwnd = $newHwnd[0]
        }
        DebugMessage "Found window $newHwnd"
        WaitForWindowToShow $newHwnd
    }

    $retHwnd.Value = $newHwnd
}

function StartProcess($object)
{
    try
    {
        $process = [Environment]::ExpandEnvironmentVariables($object.process)
        if ($object.processName -eq '')
        {
            $object.processName = $process | select-string '(.*[\\/])?([^.]+)(\..*)*' | foreach {$_.matches.groups[2].value} # Just the process naem bit
        }

        if ($object.skipIfAlreadyRunning)
        {   # See if the process is already running
            $running = get-process $object.processName -ErrorAction SilentlyContinue
            if ($running)
            {
                OutputObjectMessage $object "Skipped - already running"
                $object.done = $true
                return
            }
        }

        if ($object.waitForClass)
        {
            $object.windowClassList = [Win.API]::FindWindowsWithClassname($object.waitForClass)
            DebugMessage $object.windowClassList
        }

        if ($object.launchOnDesktop)
        {
            if ($object.desktop -eq -1)
            {
                OutputObjectMessage $object "No desktop specified - needed for launchOnDesktop flag"
                $object.launchOnDesktop = $false
            }
            else
            {
                Get-Desktop $($object.desktop - 1) | Switch-Desktop
                $object.desktop = -1
            }
        }

        if ($object.arguments.count -gt 0)
        {
            $object.proc = Start-Process $process -ArgumentList $object.arguments -PassThru
        }
        else
        {
            $object.proc = Start-Process $process -PassThru
        }

        if ($object.waitForClass)
        {
            OutputObjectMessage $object "Waiting for new window class window"
            $sleepAmount = 0
            do
            {
                $newHwnd = $null
                FindNewWindow $object ([ref]$newHwnd)

                start-sleep -Milliseconds $sleepAmount
                $sleepAmount += 100
            } while (!$newHwnd)

            $object.hwnd = $newHwnd

            if ($object.hassplash)
            {
                WaitForWindowToStopShowing $newHwnd
            }
        }
        else
        {
            OutputObjectMessage $object "Started"
        }

        if ($object.waitForProcessToClose)
        {
            OutputObjectMessage $object "Waiting for process to close"
            do
            {
                start-sleep -Milliseconds 100
                $object.proc.Refresh()
            } while (!$object.proc.HasExited)
        }
    }
    catch
    {
        OutputObjectMessage $object "Failed to start process $process! [$_]"
        $object.done = $true
        return
    }

    if ($object.hasSplash)
    {
        $waitAmount = 0
        while ($object.proc.HasExited -ne $true -and $object.proc.MainWindowHandle -eq 0)
        {
            start-sleep -Milliseconds $waitAmount
            $waitAmount += 10
            $object.proc.Refresh()
        }
        $object.hwnd = $object.proc.MainWindowHandle
    }
}

function TryToMove($object)
{
    if (!$object.hwnd -or $object.hasSplash)
    {
        $object.proc.Refresh()

        if ($object.proc.HasExited)
        {
            OutputObjectMessage $object "Process has exited!"
            $object.done = $true
            return
        }

        if ($object.hwnd -eq $object.proc.MainWindowHandle)
        {
            return
        }
        $object.hasSplash = $false
        $object.hwnd = $object.proc.MainWindowHandle
        WaitForWindowToShow $object.hwnd
    }

    OutputObjectMessage $object "Moving window"

    # Window has been created - move it
    if ($object.desktop -ne -1)
    {
        if ($object.retry -gt 0)
        {
            start-sleep -Milliseconds $($object.retry * 200)

            if ($object.WaitForClass)
            {
                $newHwnd = $null
                FindNewWindow $object ([ref]$newHwnd)
                if ($newHwnd -and ($newHwnd -ne $object.hwnd))
                {
                    DebugMessage "Taking new hwnd: $newHwnd"
                    $object.hwnd = $newHwnd
                    $object.windowClassList.Add($object.hwnd)
                }
            }
        }

        $destDesktop = Get-Desktop $($object.desktop - 1)
        DebugMessage "Moving $($object.hwnd) to desktop $($object.desktop) - $destDesktop"
        try
        {
            $object.hwnd | Move-Window $destDesktop | Out-Null
        }
        catch
        {
            $object.retry++
            if ($object.retry -eq 10)
            {
                OutputObjectMessage $object "Failed to move window ($_)"
                $object.desktop = -1
            }
            return
        }
    }

    if ($object.showOnAllDesktops)
    {
        $object.hwnd | Pin-Window
    }

    $curState = [Win.API]::GetWindowState($object.hwnd)
    $SW_NORMAL = 1; $SW_MINIMIZE = 6; $SW_MAXIMIZE = 3

    if ($object.monitor -ne -1)
    {
        $screens = [System.Windows.Forms.Screen]::AllScreens
        $screen = $($screens | ?{$_.DeviceName -like "*DISPLAY*$($object.monitor)"})
        if (!$screen)
        {
            write-host "Monitor $($object.monitor) not present!"
        }
        else
        {
            if ($curState -ne $SW_NORMAL)
            {
                if ($object.state -eq '')
                {
                    if ($curState -eq $SW_MAXIMIZE)
                    {
                        $object.state = "Max"
                    }
                    if ($curState -eq $SW_MINIMIZE)
                    {
                        $object.state = "Min"
                    }
                }
                $curState = $SW_NORMAL
                [Win.API]::RestoreWindow($object.hwnd);
            }

            if ($object.x -eq -1 -and $object.y -eq -1)
            {    # Get current x,y and move to same x,y on specified monitor
                $winX = 0
                $winY = 0
                [Win.API]::GetWindowPos($object.hwnd, [ref]$winX, [ref]$winY)
                $screens | % {
                    if (($winX -gt $_.Bounds.X -and $winX -lt $_.Bounds.X + $_.Bounds.Width) -and
                        ($winY -gt $_.Bounds.Y -and $winY -lt $_.Bounds.Y + $_.Bounds.height))
                    {
                        $winX -= $_.Bounds.X
                        $winY -= $_.Bounds.Y
                    }
                }
                $object.x = $winX
                $object.y = $winY
            }

            # Calculate origin of monitor and adjust x and y for this monitor
            $object.x = $object.x + $screen.Bounds.X
            $object.y = $object.y + $screen.Bounds.Y
        }
    }

    if ($object.state -eq 'Normal' -and $curState -ne $SW_NORMAL)
    {
        [Win.API]::RestoreWindow($object.hwnd);
    }

    if (($object.x -ne -1 -and $object.y -ne -1) -or ($object.width -ne -1 -and $object.height -ne -1))
    {
        [Win.API]::SetWindowPos($object.hwnd, $object.x, $object.y, $object.width, $object.height)
    }

    if ($object.state -eq 'Max')
    {
        [Win.API]::MaximiseWindow($object.hwnd);
    }
    if ($object.state -eq 'Min')
    {
        [Win.API]::MinimiseWindow($object.hwnd);
    }

    $object.done = $true
}

#Main logic

#Make running window on top and locked to all desktops (so output can be seen)
$currentWindow = Get-Process | Where ID -eq $PID | % { $_.MainWindowHandle }
$currentWindow | Pin-Window
$currentDesktop = Get-Desktop
[Win.API]::SetWindowOnTop($currentWindow)

#Read data into array of objects
#################
$objects = ReadData "$data_filename"
write-host "File read: $data_filename"

#Start the requried processes
#################
foreach ($object in $objects)
{
    StartProcess $object
    if (!$object.done -and
        # Check if anything needs to move
        (    
        ($object.x -ne -1 -and $object.y -ne -1) -or
        ($object.width -ne -1 -and $object.height -ne -1) -or
        ($object.desktop -ne -1) -or
        ($object.monitor -ne -1) -or
        ($object.showOnAllDesktops) -or
        ($object.state -ne '')
        ))
    {
        #Move the window (once showing) to the required desktop/monitor/position/state
        $waitAmount = 0
        $waitCount = 0
        while (!$object.done)
        {
            if ($waitCount -gt 2000)
            {
                $waitCount = 0
                if ($object.hasSplash)
                {
                    OutputObjectMessage $object "Waiting for splash to close"
                }
                else
                {
                    OutputObjectMessage $object "Waiting for window to show"
                }
            }
            TryToMove($object)

            start-sleep -Milliseconds $waitAmount
            $waitCount += $waitAmount
            if ($waitAmount -lt 250)
            {
                $waitAmount += 10
            }
        }
    }
}

#Restore window to normal state
$currentWindow | UnPin-Window
$currentWindow | Move-Window $currentDesktop | Out-Null
$currentDesktop | Switch-Desktop
[Win.API]::SetWindowNotOnTop($currentWindow)
 
$childConsole = Get-WmiObject win32_process | where {($_.ParentProcessID -eq $pid) -and ($_.Name -eq "conhost.exe")}
if ($childConsole)
{
    Stop-Process $childConsole.ProcessID
}

write-host "fin"

# SIG # Begin signature block
# MIIFdgYJKoZIhvcNAQcCoIIFZzCCBWMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUc9kY1eFmBFafiimIF6Jm2ir
# PhqgggMOMIIDCjCCAfKgAwIBAgIQJ7QhgRtobKJFyrX3zvZHpjANBgkqhkiG9w0B
# AQUFADAdMRswGQYDVQQDDBJMb2NhbCBDb2RlIFNpZ25pbmcwHhcNMjAwOTAyMTUw
# NjQ0WhcNMjEwOTAyMTUyNjQ0WjAdMRswGQYDVQQDDBJMb2NhbCBDb2RlIFNpZ25p
# bmcwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDZgckf5EM3ekbD9UHM
# uPAb3n+ruyIkD5y/uRfKbjQ9R6hpjzz5MK8qD9EOPn1rX9QUy1QJyuGQENCbOAot
# k9UE6NUtdv6cOT+4EHtcyfpW03l1LWaBpplS+4CBddfig03rpyaqQuVJe93WQ4Ar
# eUAvWozM9/Smt39MxwsyneTAHWzPFJnBXBry3OfGqg7sfJdFh02+KDOfgBfDUrGT
# dFiXzwPyZt1MfuHbQUvd+o1ErneWSqKXFjnxP/8va7mSLFjXEDZqLhm+p96jlGEs
# 0nNSLf7AfG/ZMTKfMY/X77fd08U1Dy//xc/e4qkraQfiN5TJCz8im/E6WFzrF1eQ
# XRJFAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcD
# AzAdBgNVHQ4EFgQU9nmxQrbevGH/Lmr0GxeZa9LIwJ0wDQYJKoZIhvcNAQEFBQAD
# ggEBAC4XGOmt8kosp5Wdu5FbVmLUanHNUcA0S46UyEyNK8yBDnGiNJf0hVG5D/Jh
# w4mk3BuKZjFcsV5bA1ZKcmE6hrUFxmGGfgnFmEGmdI+Tt9vR0bv3re7SG4jlaMnu
# VgOEjn+AIpon8djn10Warb6q45Q28O+3lWP0D832Sx69sI06QNbpzB88QQZJNgBD
# 29AmDvESdzCyCRH1aJAN93AgjwmXkTzgfnWL3N6myKYRMmSfShdJD2s/GxoB0hZy
# 1OleXmsYOEZP9KUkqo5p7eW0Koxq0EdzTmuvJSUeRr7MT+PpVrp/vaKYRg/DTGxW
# jcTmxfUgz2AjWi+8UcjXZAjw95kxggHSMIIBzgIBATAxMB0xGzAZBgNVBAMMEkxv
# Y2FsIENvZGUgU2lnbmluZwIQJ7QhgRtobKJFyrX3zvZHpjAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUEVvgsPwQL/2pkQ0KlVTZpVggSI8wDQYJKoZIhvcNAQEBBQAEggEAzH8fGKg3
# O3GIaU5qIMsDaqlU31KMo8zYPBgz5Auz2V15htbYhNUXY1S5zGmAhLr9aTZvUrjG
# hNYwLf8+AzxLYHj/HyQROrNHmwgfugcVT/WmBuzo8d98JShrijui0C6nOPTu9Nzi
# hFkm7Q91OZtUnsjjRMYgDoh5n0MldrUAXqz9CF9GSsPhdwa0jXPWBgG3pBvwT2cN
# S4lFnYWEXcMK2wEjky+5akqw/npvlZ8sYyUUXRQMvWx28RJvy4xbFZe81eCSyGMQ
# 6tqe7ee5wLXlsucXbEKrDtTlK679lUmb8gi92k6UGvw8vPRN8DfiDP1Bd/pCO3sQ
# UxOm8v0ZD3dOXg==
# SIG # End signature block
