<#
.SYNOPSIS
Quick description of this script

.DESCRIPTION
Longer description of what this script does

.PARAMETER FirstName
This parameter does blablabla

.PARAMETER LastName
This parameter does blablabla

.INPUTS
None. You cannot pipe objects to that script.

.OUTPUTS
None for now

.EXAMPLE

C:\PS> .\Full-Name.ps1
Your full name is Merlin the Wizard

.EXAMPLE

C:\PS> .\Full-Name.ps1 -FirstName "Jane" -LastName "Doe"
Your full name is Jane Doe

.EXAMPLE

C:\PS> .\Full-Name.ps1 "Jane" "Doe"
Your full name is Jane Doe

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
https://github.com/SammyKrosoft
#>
Param(
    [String]$FirstName,
    [String]$LastName
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
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

<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>

<# /EXECUTIONS #>

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
