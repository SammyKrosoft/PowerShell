<#
.SYNOPSIS
This script loops through a computer status, and display reachability status change.

.PARAMETER Computer
The Computer to monitor reachability

.PARAMETER FileName
If specified, logs the ping change into a file

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
        [Parameter(Mandatory = $False, Position = 0, ParameterSetName = "NormalRun")][string]$ComputerName = $($ENV:ComputerName),
        [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$FileName = $null,
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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

Function IsComputerReachable {
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $False, Position = 0, ParameterSetName = "NormalRun")][string]$ComputerName = $($ENV:ComputerName),
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    If (Test-Connection $ComputerName -Count 1 -ErrorAction SilentlyContinue) {
        Return $True
    } Else {
        Return $False
    }
}

Function Show-ComputerStatusChange ($ComputerName,$SleepTime=2,$FileName=$null) {
    $FirstStatus = IsComputerReachable $ComputerName
    $CurrentTest = $FirstStatus
    $CurrentTime = $((Get-Date).ToString())

    Write-Host "Beginning..."
    if ($FileName) {
            $FileLine = "$CurrentTime#$ComputerName#is up#$CurrentTest"
            Add-Content -Value $FileLine -Path $FileName -Force
    }

    Write-host "$ComputerName" -ForegroundColor Yellow -NoNewline
    Write-Host " at " -NoNewLine
    Write-Host "$CurrentTime" -foregroundcolor Cyan -NoNewline
    Write-Host ", is reachable => " -NoNewLine
    If ($CurrentTest -eq $True) {$StatusColor = "Green"} Else {$StatusColor = "Red"}
    Write-Host "$CurrentTest" -ForegroundColor $StatusColor

    While ($True) {
        $LastTest = $CurrentTest
        $LastTime = $CurrentTime
        Sleep $SleepTime
        $CurrentTest = IsComputerReachable $ComputerName
        $CurrentTime = $((Get-Date).ToString())
        
        # write-host ("Last Test and time") + $LastTime + $LastTest
        # Write-Host ("Current Test and time") + $CurrentTime + $CurrentTest
        If ($CurrentTest -ne $LastTest){
            Write-Host "Status changed - Last status was = $LastTest, and it changed at $CurrentTime to => $CurrentTest" -BackgroundColor Yellow -ForegroundColor Blue
        }
    }
}

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
cls
Show-ComputerStatusChange -ComputerName $ComputerName -FileName $OutputReport
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
