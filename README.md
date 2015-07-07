# PinnedApplications
PowerShell module to support pinning items to the Taskbar and Start Menu

Running **PowerShell V5**? Run this to install the module:
````PowerShell
Install-Module -Name PinnedItem
````

###Examples

````PowerShell
Get-PinnedItem
````

````PowerShell
 New-PinnedItem -TargetPath "C:\Program Files (x86)\Internet Explorer\iexplore.exe" -Type TaskBar
````

````PowerShell
$TargetPath = 'PowerShell.exe'
$ShortCutPath = 'WinDbg.lnk'
$Argument = "-ExecutionPolicy Bypass -NoProfile -NoLogo -Command `"& 'C:\users\proxb\desktop\Windbg.exe'`""
$Icon = 'C:\users\proxb\desktop\Windbg.exe'
New-PinnedItem -TargetPath $TargetPath -ShortCutPath $ShortcutPath -Argument $Argument -Type TaskBar -IconLocation $Icon
````

````PowerShell
Get-PinnedItem -Type StartMenu | Where {$_.Name -eq 'Snipping Tool'} | Remove-PinnedItem
````
