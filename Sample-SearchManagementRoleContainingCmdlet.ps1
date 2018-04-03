<#
.SYNOPSIS
    Just prints the permissions needed to run the designed cmdlet

.DESCRIPTION

This script dumps the permissions necessary to run the deisgned cmdlet. Just put the cmdlet name as a parameter, and check the Role needed for this.

.PARAMETER CmdLet
Put the CMDLet name as an input of this script


.INPUTS
The cmdlet you want to know in where Management Role it is...

.OUTPUTS
The script returns you the table of the Role(s) that contain the cmdlet you want to run

.EXAMPLE

C:\PS> .\Search-ManagementRoleContainingCmdlet.ps1 -CmdLet Add-ADPermission

Role                                                                                            RoleAssigneeType RoleAssigneeName                                       
----                                                                                            ---------------- ----------------                                       
Active Directory Permissions                                                                           RoleGroup Organization Management                                


The script took 0.1496705 seconds to execute...

.EXAMPLE

C:\PS> .\Full-Name.ps1 "Jane" "Doe"
Your full name is Jane Doe

.LINK

https://docs.microsoft.com/en-us/powershell/exchange/exchange-server/find-exchange-cmdlet-permissions?view=exchange-ps

.LINK

https://github.com/SammyKrosoft
#>
Param(
    [string]$Cmdlet = "Add-ADPermission"
)

<# ------- SCRIPT_HEADER (Only Get-Help comments above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1.0"
# Log or report file definition
$LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$LogOrReportFile2 = "$((Get-Location).Path)\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>



$Perms = Get-ManagementRole -Cmdlet $Cmdlet
$Perms | Select Name | Foreach {Get-ManagementRoleAssignment -Role $_.Name -Delegating $false } | Ft Role,RoleAssigneeType, RoleAssigneeName


<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
