[CmdletBinding()] Param( [parameter(mandatory=$false)] [string]$data_filename )

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

Function ReadData($file)
{
    $objects = New-Object System.Collections.ArrayList

    $data = $(get-content $data_filename | ConvertFrom-Json)

    $maxDesktop = -1

    foreach ($item in $data)
    {
    	$object = [PSCustomObject] @{ process = ''; arguments = @(); desktop = -1; monitor = -1; x = -1; y = -1; width = -1; height = -1; state = ''; hasSplash = $false; showOnAllDesktops = $false; skipIfAlreadyRunning = $false; hwnd = 0; done = $false; proc = 0; index = -1; processName = '' }

        $item.PSOBject.Properties | ForEach-Object {
            $value = $_.Value
            switch ($_.Name)
            {
                'process' { $object.process = $value }
                'arguments' { $object.arguments += $value }
                'desktop' { $object.desktop = $value }
                'monitor' { $object.monitor = $value }
                'x' { $object.x = $value }
                'y' { $object.y = $value }
                'width' { $object.width = $value }
                'height' { $object.height = $value }
                'state' { $object.state = $value }
                'hassplash' { $object.hassplash = $value }
                'showOnAllDesktops' { $object.showOnAllDesktops = $value }
                'skipIfAlreadyRunning' { $object.skipIfAlreadyRunning = $value }
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

            $object.index = $objects.Add($object)
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

function OutputObjectMessage($object, $msg)
{
    write-host -NoNewline -ForegroundColor Red "$($object.index): $($object.processName) "
    write-host $msg
}

function StartProcess($object)
{
    try
    {
        $process = [Environment]::ExpandEnvironmentVariables($object.process)
        $object.processName = $process | select-string '(.*[\\/])?([^.]+)(\..*)*' | foreach {$_.matches.groups[2].value} # Just the process naem bit

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

        if ($object.arguments.count -gt 0)
        {
            $object.proc = Start-Process $process -ArgumentList $object.arguments -PassThru;
        }
        else
        {
            $object.proc = Start-Process $process -PassThru;
        }
        OutputObjectMessage $object "Started"
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
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); 

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
public static void SetWindowOnTop(IntPtr hWnd)
{
    SetWindowPos(hWnd, new IntPtr( /*HWND_TOPMOST*/ -1), 0, 0, 0, 0, SWP_FLAGS.NOMOVE | SWP_FLAGS.NOSIZE);
}
public static void SetWindowNotOnTop(IntPtr hWnd)
{
    SetWindowPos(hWnd, new IntPtr( /*HWND_NOTTOPMOST*/ -2), 0, 0, 0, 0, SWP_FLAGS.NOMOVE | SWP_FLAGS.NOSIZE);
}
"@

Add-Type -MemberDefinition $winApi -Namespace Win -Name API
Add-Type -AssemblyName System.Windows.Forms

function TryToMove($object)
{
    $object.proc.Refresh()

    if ($object.hwnd -ne $object.proc.MainWindowHandle)
    {    # Window has been created - move it
        $object.hwnd = $object.proc.MainWindowHandle

        if ($object.desktop -ne -1)
        {
            $object.hwnd | Move-Window (Get-Desktop $($object.desktop - 1)) | Out-Null
        }
        if ($object.showOnAllDesktops)
        {
            $object.hwnd | Pin-Window
        }

        if ($object.monitor -gt 1 -and $object.x -ne -1 -and $object.y -ne -1)
        {    # Calculate origin of monitor and adjust x and y for this monitor
            $screens = [System.Windows.Forms.Screen]::AllScreens
            $screen = $($screens | ?{$_.DeviceName -like "*DISPLAY$($object.monitor)"})
            if (!$screen)
            {
                write-host "Monitor $($object.monitor) not present!"
            }
            else
            {
                $object.x += $screen.Bounds.X
                $object.y += $screen.Bounds.Y
            }
        }

        [Win.API]::SetWindowPos($object.hwnd, $object.x, $object.y, $object.width, $object.height)

        if ($object.state -eq 'Max')
        {
            [Win.API]::MaximiseWindow($object.hwnd);
        }
        if ($object.state -eq 'Min')
        {
            [Win.API]::MinimiseWindow($object.hwnd);
        }

        OutputObjectMessage $object "Process moved"
        $object.done = $true
    }

    if ($object.proc.HasExited)
    {
        OutputObjectMessage $object "Process has exited!"
        $object.done = $true
        return
    }
}

#Main logic

#Make running window on top and locked to all desktops (so output can be seen)
$currentWindow = Get-Process | Where ID -eq $PID | % { $_.MainWindowHandle }
$currentWindow | Pin-Window
[Win.API]::SetWindowOnTop($currentWindow)

#Read data into array of objects
#################
$objects = ReadData "$data_filename"

#Start the requried processes
#################
$leftToProcess = 0
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
        ($object.state -eq 'min' -or $object.state -eq 'max')
        ))
    {
        $leftToProcess++
    }
}

#Move the windows (once created) to the required desktop/monitor/position/state
#################
$waitAmount = 0
$waitCount = 0
$showDetails = $false
while ($leftToProcess -gt 0)
{
    start-sleep -Milliseconds $waitAmount
    $waitCount += $waitAmount
    if ($waitCount -gt 2000)
    {
        write-host "Reaminig items ($leftToProcess)"
        $showDetails = $true
    }
    foreach ($object in $objects)
    {
        if (!$object.done)
        {
            if ($showDetails)
            {
                OutputObjectMessage $object "Waiting"
            }
            TryToMove($object)
            if ($object.done)
            {
                $leftToProcess -= 1
            }
        }
    }
    if ($waitAmount -lt 250)
    {
        $waitAmount += 10
    }
    if ($showDetails)
    {
        $waitCount = 0
        $showDetails = $false
    }
}

#Restore window to normal state
$currentWindow | UnPin-Window
[Win.API]::SetWindowNotOnTop($currentWindow)

write-host "fin"
