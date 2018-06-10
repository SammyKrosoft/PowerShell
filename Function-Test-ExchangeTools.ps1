Function Test-ExchTools(){
    <#
.SYNOPSIS
This small function will just check if you have Exchange tools installed or available on the
current PowerShell session 

.DESCRIPTION

Just a dummy script that prints your first name and last name.
Takes any strings for first name and last name.

.PARAMETER FirstName
Specifies the First Name. "Merlin" is the default.

.PARAMETER LastName
Specifies the last name. "the Wizard" is the default.

.INPUTS

None. You cannot pipe objects to that script.

.OUTPUTS

System.String. The script (Full-Name.ps1 or whatever you name it) returns a string with the full
name.

.EXAMPLE
Test-ExchTools
=> will exit the script/program si Exchange tools are not installed

}

#>
    Try
    {
        Get-command Get-mailbox -ErrorAction Stop
        $ExchInstalledStatus = $true
        $Message = "Exchange tools are present !"
        Write-Host $Message -ForegroundColor Blue -BackgroundColor Red
    }
    Catch [System.SystemException]
    {
        $ExchInstalledStatus = $false
        $Message = "Exchange Tools are not present ! This script/tool need these. Exiting..."
        Write-Host $Message -ForegroundColor red -BackgroundColor Blue
        Exit
    }
    Return $ExchInstalledStatus
}
