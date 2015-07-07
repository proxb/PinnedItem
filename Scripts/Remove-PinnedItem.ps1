Function Remove-PinnedItem {
    <#
        .SYNOPSIS
            Remove pinned items from StartMenu/Taskbar

        .DESCRIPTION
            Remove pinned items from StartMenu/Taskbar

        .PARAMETER InputObject
            Full path of pinned items that will be removed.

            Acceptable values:
                StartMenu
                Taskbar

        .NOTES
            Name: Remove-PinnedItem
            Author: Boe Prox
            Version History
                1.0 //Boe Prox - 03 June 2015
                    - Initial Build

        .EXAMPLE
            Get-PinnedItem -Type TaskBar | Remove-PinnedItem

            Description
            -----------
            Removes all pinned items from the TaskBar

        .EXAMPLE
            Get-PinnedItem -Type StartMenu | Where {$_.Name -eq 'Snipping Tool'} | Remove-PinnedItem

            Description
            -----------
            Removes Snipping Tool from the StartMenu as a pinned item.
    #>
    [cmdletbinding(
        SupportsShouldProcess = $True
    )]
    Param (
        [parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
        [System.File.PSItem.PinnedItem[]]$InputObject
    )
    Begin {
        Write-Verbose "Creating Shell ComObject"
        $Shell = New-Object -ComObject Shell.Application
    }
    Process {
        ForEach ($Item in $InputObject) {
            Write-Verbose "Item: $($Item.fullname)"
            Write-Verbose "Type: $($Item.Type)"
            If ($Item.fullname -match '^(?<Path>.*\\)(?<File>.*)$') {
                $Path = $Matches.Path
                $File = $Matches.File
                $NameSpace = $Shell.NameSpace($Path)
                $NameSpaceFile = $NameSpace.ParseName($File)
                Switch ($Item.Type) {
                    'Taskbar' {
                        $_Verb = ConvertToVerb -Action UnpinfromTaskbar
                        $Verb = $NameSpaceFile.Verbs() | Where {
                            $_.Name -eq $_Verb
                        }                    
                    }
                    'StartMenu' {
                        $_Verb = ConvertToVerb -Action UnpinfromStartMenu
                        $Verb = $NameSpaceFile.Verbs() | Where {
                            $_.Name -eq $_Verb
                        }                     
                    }
                    Default {
                        Write-Warning "No Type found!"
                        Continue
                    }
                }
                If ($Verb) {
                    If ($PSCmdlet.ShouldProcess($Item.Fullname, "Unpin From $($Type)")) {
                        $Verb.DoIt()
                    }
                } Else {
                    Try {
                        Remove-Item -Path $Item.Fullname -Erroraction Stop
                    } Catch {
                        Write-Warning "Unable to delete file: $($Item.Fullname)"
                    }
                }
            } Else {
                Write-Warning "Unable to parse File and Path from provided Fullname: $($Item.Fullname)!"
            }
        }
    }
    End {
        Write-Verbose 'Cleanup ComObject'
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Shell)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable Shell -ErrorAction SilentlyContinue -WhatIf:$False
    }
}