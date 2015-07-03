$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#region Define Custom Object
Add-Type -TypeDefinition @"
using System;

public enum PinnedType
{
    StartMenu,
    TaskBar
}

namespace System.File.PSItem
{
    public class PinnedItem
    {
        public string Name;
        public string FullName;
        public string Destination;
        public PinnedType Type;
    }
}
"@ -Language CSharpVersion3
#endregion Define Custom Object

#region Load Functions
Try {
    Get-ChildItem "$ScriptPath\Scripts" -Filter *.ps1 | Select -Expand FullName | ForEach {
        $Function = Split-Path $_ -Leaf
        . $_
    }
} Catch {
    Write-Warning ("{0}: {1}" -f $Function,$_.Exception.Message)
    Continue
}
#endregion Load Functions

#region Aliases
New-Alias -Name gpi -Value Get-PinnedItem
New-Alias -Name rpi -Value Remove-PinnedItem
New-Alias -Name npi -Value New-PinnedItem
#endregion Aliases

Export-ModuleMember -Alias * -Function *pinned*