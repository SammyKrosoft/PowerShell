Function Check-PSVersion {
    <#
    .SYNOPSIS    
    Just printing Powershell version and returning the Powershell
    Major version number.

    .DESCRIPTION
    Just printing Powershell version and returning the Powershell
    Major version number -> use it to adapt part of your scripts
    and execute one or the other Powershell cmdlet whether you're
    on PS Version 2 or a higher one...

    .OUTPUTS
    Output Powershell major version
    
    .EXAMPLE
    Test-PSVersion
        
    #>

    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        $msgInfoPSVersion2orless = "Powershell version is 2 or less... consider upgrading !"
        Write-host $msgInfoPSVersion2orless -BackgroundColor red
    } Else {
        $msgInfoPSVersion3plus = "Powershell version: v$PowerShellMajorVersion => Powershell is v3 or later -> good, keep it up !"
        Write-Host $msgInfoPSVersion3plus -BackgroundColor green -ForegroundColor black
    }
    Return $PowerShellMajorVersion
}
    



