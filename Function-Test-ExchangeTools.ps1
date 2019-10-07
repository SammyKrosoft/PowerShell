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


Function Test-Exchange2016Tools {
    $Exch2016InstalledStatus = $false
    Try {
        # Checking if Set-OutlookAnywhere comes with parameter "ExternalClientAuthenticationMethod", it's Exchange 2013/2016 tools
        $GetHelpTest = get-help Set-OutlookAnywhere -parameter ExternalClientAuthenticationMethod -ErrorAction Stop
        $Exch2016InstalledStatus = $true
        $message = "Exchange 2016 tools detected"
        Write-Host $message -ForegroundColor Blue -BackgroundColor Green
    } 
    catch {
        $Exch2016InstalledStatus = $false
        $message = "Exchange 2016 tools not detected"
        Write-Host $message -BackgroundColor Red -ForegroundColor Blue
    } 
    Finally {
        Write-Host "Done"
    }
    Return $Exch2016InstalledStatus
}


Function Test-Exchange2010Tools {
    $Exch2010InstalledStatus = $false
    Try {
        # Checking if Set-OutlookAnywhere comes with parameter "ExternalClientAuthenticationMethod", it's Exchange 2013/2016 tools
        $GetHelpTest = get-help Set-OutlookAnywhere -parameter ClientAuthenticationMethod -ErrorAction Stop
        $Exch2010InstalledStatus = $true
        $message = "Exchange 2010 tools detected"
        Write-Host $message -ForegroundColor Blue -BackgroundColor green
    } 
    catch {
        $Exch2010InstalledStatus = $false
        $message = "Exchange 2010 tools not detected"
        Write-Host $message -BackgroundColor Red -ForegroundColor Blue
    } 
    Finally {
        Write-Host "Done"
    }
    Return $Exch2010InstalledStatus
}