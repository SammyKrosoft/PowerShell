Function Test-ExchTools(){
<#
.SYNOPSIS
This small function will just check if you have Exchange tools installed or available on the
current PowerShell session.

.DESCRIPTION
The presence of Exchange tools are checked by trying to execute "Get-ExBanner", one of the basic Exchange
cmdlets that runs when the Exchange Management Shell is called.

Just use Test-ExchTools in your script to make the script exit if not launched from an Exchange
tools PowerShell session...

.EXAMPLE
Test-ExchTools
=> will exit the script/program si Exchange tools are not installed
#>
    Try
    {
        #Get-command Get-ExBanner -ErrorAction Stop
        Get-command Get-Mailbox -ErrorAction Stop
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
