<#
.SYNOPSIS

Prints your first name and last name.
Get this help from header by typing Get-Help .\YourScript.ps1 -Full

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
#Your message
$Message = "Hello world !"


<# /DECLARATIONS #>



<# -------------------------- FUNCTIONS -------------------------- #>
function Function1
{
    param($FirstName = "John", $LastName = "Doe")
    Write-Host "Your full name is $FirstName $LastName ... can I call you $FirstName ?"
}

<# /FUNCTIONS #>




<# -------------------------- EXECUTIONS -------------------------- #>
Write-Debug $Message
Function1

<# /EXECUTIONS #>

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
