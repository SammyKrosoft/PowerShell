﻿<#
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
PS> .\Start-E2016ServerComponentStateToActive.ps1 -CheckOnly

Will check
#>

#Requires -version 3.0

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)][switch]$HybridServer,
    [Parameter(Mandatory = $false)][switch]$CheckOnly
)

Function Title1 ([string]$title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-host
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host    
}

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

If ($CheckOnly) {
    Title1 "Check only specified - will just list inactive components without trying to activate ..."
} Else {
    Title1 "CheckOnly NOT specified ... will try to activate everything if more than 2 components are inactive..."
}

$E2016NamesList = @()
$E2016 = Get-ExchangeServer | ? {$_.AdminDisplayVersion -match "15.1"} 
Foreach ($item in $E2016){$E2016NamesList += $($item.Name)}

$counter = 0
Foreach ($Server in $E2016){
    Title1 $Server
    write-progress -id 1 -Activity "Activating all components" -Status "Server $Server" -PercentComplete $($Counter/$($E2016.Count)*100)
    $Counter++

    #Get the status of component 
    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
    #$ComponentStateStatus | ft Component,State -Autosize
    If (!($HybridServer)){
        Write-Host "This is an On-Premises only environment (aka not Hybrid, not synchronizing with the cloud). We don't need ForwardSyncDaemon and ProvisioningRPS Components - leaving these as-is"
        $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive" -and $_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
        $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active" -and $_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
    } Else {
        Write-Host "You specified the -HybridServer parameter, indicating that this is an On-Premises environment syncinc with O365. All Server Components need to be active..."
        $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive"}
        $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active"}
    }
    
    $NbActiveComponents = $ACtiveComponents.Count
    $NbInactiveComponents = $InactiveComponents.Count

    Write-Host "There are $NbActiveComponents active components, and $InactiveComponents inactive components on server $($Server.Name)"

    If ($NbInactiveComponents -gt 0){
        Write-Host "There are only $NbInactiveComponents, everything looks good ... here are the list of inactive components:"
        $InactiveComponents | ft Component,State -Autosize
        Continue
    } Else {
        Write-host "More than 2 components are not active - we have $NbInactiveComponents ..."
        $InactiveComponents | ft Component
        If (!($CheckOnly)){
            Write-host "... trying to re-activate all inactive components..." 
                $Counter1 = 0
                Foreach ($Component in $InactiveComponents) {
                Write-progress -id 2 -ParentId 1 -Activity "Setting component states" -Status "setting $($Component.Component)..." -PercentComplete ($Counter1/$NbInactiveComponents*100)
                $Command = "Set-ServerComponentState $($Server.Name) -Component $($Component.Component) -State Active -Requester Functional" 
                Write-host "Running the following command: `n$Command" -BackgroundColor Blue -ForegroundColor White
                Invoke-Expression $Command
                $Counter1++
              }
        } Else {
            Write-Host "Checking only..."
        }
    }
    #Get the new status of components
    $ComponentStateStatus = Get-ServerComponentState ($Server.Name) 
    If (!($HybridServer)){
        Write-Host "This is an On-Premises only environment (aka not Hybrid, not synchronizing with the cloud). We don't need ForwardSyncDaemon and ProvisioningRPS Components - leaving these as-is"
        $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive" -and $_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
        $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active" -and $_.Component -ne "ForwardSyncDaemon" -and $_.Component -ne "ProvisioningRps"}
    } Else {
        Write-Host "You specified the -HybridServer parameter, indicating that this is an On-Premises environment syncinc with O365. All Server Components need to be active..."
        $InactiveComponents = $ComponentStateStatus | ? {$_.State -eq "Inactive"}
        $ACtiveComponents = $ComponentStateStatus | ? {$_.State -eq "Active"}
    }

    Write-Host "There are now $($ACtiveComponents.count) active components, and $($InactiveComponents.count) inactive components"
    If ($($InactiveComponents.count) -gt 2) {Write-host "There are still some inactive components ... please troubleshoot !" -BackgroundColor Red -ForegroundColor Yellow} Else {Write-Host "$Server is now completely out of maintenance mode and component are active and functional." -ForegroundColor Yellow}
}

write-progress -id 1 -Activity "Activating all components" -Status "All done !" -PercentComplete $($Counter/$($E2016.Count)*100)
sleep 1