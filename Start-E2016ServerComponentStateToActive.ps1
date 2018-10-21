<#
.NOTES
Written by Adnan Rafique @ExchangeITPro
Modified by Samuel Drey @Microsoft
V1.1 08.06.2014
.SYNOPSIS
Bring componet to active state.
.DESCRIPTION
Bring component state to active state.
.PARAMETER Server
Specifies the DAG node Server name to be bring the component to active state.
.EXAMPLE
PS> .\SetExchangeComponentToActive.ps1 -Server Server1
#>

#Requires -version 3.0

#[CmdletBinding()]
#Param(
#    [Parameter(Mandatory = $true,
#    HelpMessage="Enter the name of Server to set component to active.")]
#    [string]$Server
#)

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
        Get-command Get-MAilbox -ErrorAction Stop
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

If (!(Test-ExchTools)){exit}

cls
$E2016NamesList = @()
$E2016 = Get-ExchangeServer | ? {$_.AdminDisplayVersion -match "15.1"} 
Foreach ($item in $E2016){$E2016NamesList += $($item.Name)}

$counter = 0
Foreach ($Server in $E2016){
    write-progress -id 1 -Activity "Activating all components" -Status "Server $Server" -PercentComplete $($Counter/$($E2016.Count)*100)
    $Counter++

    #Get the status of component 
    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
    #$ComponentStateStatus | ft Component,State -Autosize
    $ACtiveComponents = ($ComponentStateStatus | ? {$_.State -eq "Active"}).count
    $InactiveComponents = ($ComponentStateStatus | ? {$_.State -eq "Inactive"}).count

    Write-Host "There are $ACtiveComponents active compponents, and $InactiveComponents inactive components"

    If ($InactiveComponents -le 2){
        Write-Host "There are only $InactiveComponents, everything looks good ... here are the list of components:"
        $ComponentStateStatus | ft Component,State -Autosize
        Continue
    } Else {
        Write-host "More than 2 components are not active :"
        $ComponentStateStatus | ? {$_.State -eq "Active"}
        Write-host "... trying to re-activate all components..." 
        #Designates that the server is out of maintenance mode
        Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "setting ServerWideOffline..." -PercentComplete 0
        Set-ServerComponentState $($Server.Name) -Component ServerWideOffline -State Active -Requester Functional
        Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "ServerWideOffline done ... now setting Monitoring" -PercentComplete 33
        Set-ServerComponentState $($Server.Name) -Component Monitoring -State Active -Requester Functional
        Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "Monitoring done ... now setting RecoveryACtionsEnabled" -PercentComplete 66
        Set-ServerComponentState $($Server.Name) -Component RecoveryActionsEnabled -State Active -Requester Functional
        Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "All 3 components set !" -PercentComplete 100
    }
    #Get the status of components
    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
    $ComponentStateStatus | ft Component,State -Autosize
    $ACtiveComponents = ($ComponentStateStatus | ? {$_.State -eq "Active"}).count
    $InactiveComponents = ($ComponentStateStatus | ? {$_.State -eq "Inactive"}).count

    Write-Host "There are now $ACtiveComponents active compponents, and $InactiveComponents inactive components"
    If ($InactiveComponents -lt 2) {Write-host "There are still some inactive components ... please troubleshoot !" -BackgroundColor Red -ForegroundColor Yellow} Else {Write-Host "$Server is now completely out of maintenance mode and component are active and functional." -ForegroundColor Yellow}
}

write-progress -id 1 -Activity "Activating all components" -Status "All done !" -PercentComplete $($Counter/$($E2016.Count)*100)
sleep 1