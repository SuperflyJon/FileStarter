# FileStarter - Start processes on multiple desktops

This script allows you to start a list of processes on specific desktops, monitors, window size, and window state on windows 10.  You need to create a Json file with the details of what to start.  See files.json in the repository for an example.

## Usage

powershell \<path to FileStarter.ps1\> [\<path to json file\>]

The script will either use the local files.json file or another file if you pass the path as an argument to the script

## Settings

Simple processes (e.g. notepad) just create a new window and will work easily.  Other appications have splash screens, create multiple processes or windows and in general start in custom ways.  The flags below help to cover most of these cases:

- 'process' | **The only required setting**.  This is the string which will run the process, one way to check if this is correct is to test it in the windows run dialog (shortcut win+R)
- 'arguments' | A list of arguments to pass to the process
- 'processName' | A pretty name for the output, also an override for the process name checked in 'skipIfAlreadyRunning'
- 'desktop' | The number of the desktop to put the main window on, if it doesn't exist it will be created
- 'desktopname' | Will rename desktop in windows task view (requires version 2004 of windows 10)
- 'monitor' | The monitor number to put the window on
- 'x', 'y' | The x and y co-ord of where to put the window.  This is from (0, 0) on the specific monitor
- 'width', 'height' | The width and height to set the window size
- 'state' | Set the window state to 'normal', 'min' or 'max'
- 'hassplash' | Will wait for a splash screen to close before moving the following window
- 'showOnAllDesktops' | Sets the flag to pin the window to all desktops
- 'skipIfAlreadyRunning' | If the process is already running don't start it again
- 'waitForClass' | This is a good option for non-standard processes.  You can use the Spy++ tool to find the window class.  See the example file for a couple of useful examples
- 'waitForProcessToClose' | This may help for windows that get created in child processes
- 'launchOnDesktop' | If all eslse fails, this flag will switch to a desktop before launching it

## Links

The multiple desktop logic is in the VirtualDesktop.ps1 file which is a direct copy from [PSVirtualDesktop](https://github.com/MScholtes/TechNet-Gallery/tree/master/VirtualDesktop)
