<#
.SYNOPSIS
    Dumps the permissions needed to run the designed cmdlet

.DESCRIPTION
    This script dumps the permissions necessary to run the deisgned cmdlet. Just put the cmdlet name as a parameter, and check the Role needed for this.

.PARAMETER CmdLet
    Put the CMDLet name as an input of this script

.INPUTS
    The cmdlet you want to know in where Management Role it is...

.OUTPUTS
    The script returns you the table of the Role(s) that contain the cmdlet you want to run

.EXAMPLE
.\Search-ManagementRoleContainingCmdlet.ps1 -CmdLet Add-ADPermission

Role                                                                                            RoleAssigneeType RoleAssigneeName                                       
----                                                                                            ---------------- ----------------                                       
Active Directory Permissions                                                                           RoleGroup Organization Management                                

The script took 0.1496705 seconds to execute...

.LINK
https://docs.microsoft.com/en-us/powershell/exchange/exchange-server/find-exchange-cmdlet-permissions?view=exchange-ps

.LINK
https://github.com/SammyKrosoft
#>
[CmdletBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][string]$Cmdlet = "Add-ADPermission",
    [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1.0"
<# Version changes
-> v1.0
#>
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
#$LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
#$LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
$Perms = Get-ManagementRole -Cmdlet $Cmdlet
$Perms | Select-Object Name | ForEach-Object {Get-ManagementRoleAssignment -Role $_.Name -Delegating $false } | Format-Table Role,RoleAssigneeType, RoleAssigneeName

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
