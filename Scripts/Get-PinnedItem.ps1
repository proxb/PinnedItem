Function Get-PinnedItem {
    <#
        .SYNOPSIS
            Gets pinned items on StartMenu and/or Taskbar

        .DESCRIPTION
            Gets pinned items on StartMenu and/or Taskbar

        .PARAMETER Type
            Determine what types of pinned items will be returned.

            Acceptable values:
                StartMenu
                Taskbar

            Default is that all items will be returned.

        .NOTES
            Name: Get-PinnedItem
            Author: Boe Prox
            Version History
                1.0 //Boe Prox - 03 June 2015
                    - Initial Build

        .EXAMPLE
            Get-PinnedItem

            Description
            -----------
            Returns all pinned items (StartMenu and Taskbar).

        .EXAMPLE
            Get-PinnedItem -Type StartMenu

            Description
            -----------
            Returns all pinned StartMenu items.
    #>

    [cmdletbinding()]
    Param(
        [parameter()]
        [PinnedType[]]$Type
    )

    $WShell = New-Object -ComObject WScript.Shell
    $TypeList = New-Object System.Collections.ArrayList
    If (-NOT $PSBoundParameters.ContainsKey('Type')) {
        $Type = 'TaskBar','StartMenu'
    }
    If ($Type -contains "TaskBar") {
        Write-Verbose "Pulling pinned Taskbar items"
        $Taskbar = "$($Env:AppData)\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        Try {
            Get-ChildItem $TaskBar -Include '*.lnk','*.url' -Recurse -ErrorAction Stop | ForEach {
                $Object = New-Object System.File.PSItem.PinnedItem 
                $Object.Name = $_.fullname -replace '.*\\(.*)\..*','$1'
                $Object.FullName = $_.Fullname
                $Object.Destination = $wshell.CreateShortcut($_.Fullname).TargetPath
                $Object.Type = 'TaskBar'
                $Object
            }
        } Catch {
            Write-Warning $_
        }
    }
    If ($Type -contains "StartMenu") {
        Write-Verbose "Pulling pinned StartMenu items"
        $StartMenu =  "$($Env:AppData)\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"
        Try {
            Get-ChildItem $StartMenu -Include '*.lnk','*.url' -Recurse -ErrorAction Stop | ForEach {
                $Object = New-Object System.File.PSItem.PinnedItem 
                $Object.Name = $_.fullname -replace '.*\\(.*)\..*','$1'
                $Object.FullName = $_.Fullname
                $Object.Destination = $wshell.CreateShortcut($_.Fullname).TargetPath
                $Object.Type = 'StartMenu'
                $Object
            }
        } Catch {
            Write-Warning $_
        }
    }

    Write-Verbose 'Cleanup ComObject'
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$WShell)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable WShell -ErrorAction SilentlyContinue
}