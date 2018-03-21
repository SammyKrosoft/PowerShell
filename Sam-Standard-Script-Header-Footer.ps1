<# SCRIPT_HEADER #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1.0"
<# /SCRIPT_HEADER #>

<#----------------------------------------------------------[Declarations]----------------------------------------------------------#>
#Your message
$Message = "Hello world !"

<#----------------------------------------------------------[Functions]----------------------------------------------------------#>
function _Progress{
    param($PercentComplete,$Status)
    Write-Progress -id 1 -activity "Working !" -status $Status -percentComplete ($PercentComplete)
}

<#----------------------------------------------------------[Execution]----------------------------------------------------------#>
Write-Debug $Message



<# SCRIPT_FOOTER #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $StopWatch.Elapsed.TotalSeconds seconds to execute..."
<# /SCRIPT_FOOTER #>
