Function IsPSV3 {
    <#
    .DESCRIPTION
    Just printing Powershell version and returning "true" if powershell version
    is Powershell v3 or more recent, and "false" if it's version 2.
    .OUTPUTS
    Returns $true or $false
    .EXAMPLE
    IsPSVersionV3
    #>
    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        Return $false
    } Else {
        Return $true
        }
}
    



