<# ---------------------------- SCRIPT_HEADER ---------------------------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()
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



<# DECLARATIONS #>
#Your message
$Message = "Hello world !"


<# /DECLARATIONS #>



<# FUNCTIONS #>
function Function1{
    param($FirstName="John",$LastName="Doe")
    Write-Host "Your full name is $FirstName $LastName ... can I call you $FirstName ?"
 }

 <# /FUNCTIONS #>




<# EXECUTIONS #>
Write-Debug $Message
Function1

<# /EXECUTIONS #>

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ---------------------------- #>
