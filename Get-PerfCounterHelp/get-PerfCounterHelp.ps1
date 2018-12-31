<#
.SYNOPSIS
    Get the counter category help as well as the counter help text.
    Thanks to David Martin : https://stackoverflow.com/users/1035521/david-martin

.DESCRIPTION
    Longer description of what this script does

.PARAMETER FirstNumber
    This parameter does blablabla

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    None for now

.EXAMPLE
.\Do-Something.ps1
This will launch the script and do someting

.EXAMPLE
.\Do-Something.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Do-Something.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 0, ParameterSetName = "NormalRun")][string]$CategoryName = "Processor",
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$CategoryInstance = "_Total",
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][string]$CounterName = "% Processor Time",
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
$ScriptVersion = "0.1"
<# Version changes
v0.1 : first script version
v0.1 -> v0.5 : 
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$ScriptPath\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
Function Lines {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $False, Position = 0)][int32]$Iterations = 35
    )
    For ($i=0 ; $i -le $Iterations ;$i++) {Write-host "-" -NoNewline}
    Write-host
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
cls
# Call the static method to get the Help for the category
$categoryHelp = [System.Diagnostics.PerformanceCounterCategory]::GetCategories() | ?{$_.CategoryName -like $categoryName} | select -expandproperty  CategoryHelp

# Create an instance so that GetCounters() can be called
$pcc = new-object System.Diagnostics.PerformanceCounterCategory($categoryName)
$counterHelp = $pcc.GetCounters($categoryInstance) | ?{$_.CounterName -like $counterName} | select -expandproperty CounterHelp

Lines
Write-host "Category Help " -BackgroundColor yellow -ForegroundColor Blue
Lines
Write-host "$categoryHelp"
Lines
Write-host "Counter Help " -BackgroundColor Blue -ForegroundColor Yellow
Lines
Write-host "$counterHelp"
<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
