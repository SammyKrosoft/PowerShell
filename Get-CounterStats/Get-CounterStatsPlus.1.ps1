<#
.SYNOPSIS
Script: Get-CounterStatsPlus
Original Authors: Prashanth and Praveen
Modified by : Samuel Drey aka SammyKrosoft

This script will collect the specific counters value from the multiple target machines/servers 
which will be used to analayze the performance of target servers.

.DESCRIPTION
This script will collect the specific counters value from the multiple target machines/servers 
which will be used to analayze the performance of target servers.

The script will query a defined set of counters that you define there :

$Counter = @"
Processor(_total)\% processor time 
\MSExchange RpcClientAccess\RPC Averaged Latency
\MSExchange RpcClientAccess\RPC Requests
Memory\Available MBytes 
PhysicalDisk(*)\Avg. Disk sec/Transfer 
Network Interface(*)\Bytes Total/sec
"@ 

Hint : Chase counters definitions using Powershell ! 
Example:
Get-Counter -ListSet *Memory* | Select -ExpandProperty Counter | ? {$_ -like "*available*"}
Will get you:
\Memory\Available Bytes
\Memory\Available KBytes
\Memory\Available MBytes
Then just copy and paste these on the $Counter = @() definition in the script ... cool eh !




.PARAMETER ServersTXTFile
    This parameter specified the file containing the list of servers to get Perfmon samples from.
    By default it will look for a "servers.txt" file in the same directory as the script.

.PARAMETER NumberOfSamples
    This parameter specifies how many counter samples we need to dump. Default is 5.

.PARAMETER OutputFile
    This parameter specifies the Output file. If not specified, the output file name will be built
    after the script's name, with the date and time appended, and will be stored on the same 
    directory where the script is located.

.PARAMETER CheckVersion
    This parameter Checks the script's version.

.INPUTS
    You need to have a file to import from.

.OUTPUTS
    A CSV file which name is constructed with the scripts name appended with the date and time
    of the execution.

.EXAMPLE
.\Get-CounterStatsPlus.ps1
Will execute and dump the counters stats for 5 default samples on a list of servers defined in the C:\Temp\Servers.txt file.
The detault output file will be named after the script's file with the date and time appended, on the same directory where
the script itself is located (Get-CounterStatsPlus.ps1_Date_Time.csv)

.EXAMPLE
.\Get-CounterStatsPlus.ps1 -ServersTXTfile C:\temp\Myservers.txt -NumberOfSamples 20 -OutputFile c:\ExportRequestISsue.csv
Will execute the counters stats for servers list defined in the C:\temp\Myservers.txt, for 20 samples, and store the
results in the output file specified here :C:\ExportRequestIssue.csv

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string]$ServersTXTfile = ".\servers.txt",
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][int]$NumberOfSamples = 5,
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")][string]$OutputFile,
    [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
$ScriptVersion = "1"
<# Version changes
v1 : first script version
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
#$ScriptLog = "$PSScriptRoot\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
$Answer = ""
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
#region Functions region
#Function to have the customized output in CSV format

function Global:Convert-HString {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] [String]$HString
        )

    Begin 
    {Write-Verbose "Converting Here-String to Array"}
    Process 
    {
        $HString -split "`n" | ForEach-Object {
            $ComputerName = $_.trim()
            if ($ComputerName -notmatch "#")
            {$ComputerName}    
        }
    }#Process
    End 
    {
        # Nothing to do here.
    }
}#Convert-HString

#Performance counters declaration
function Get-CounterStats { 
    param(
        [String[]]$ComputerName = $Env:ComputerName
    ) 

$Counter = @"
Processor(_total)\% processor time 
Memory\Available MBytes 
PhysicalDisk(*)\Avg. Disk sec/Transfer 
Network Interface(*)\Bytes Total/sec
"@ 

    (Get-Counter -ComputerName $ComputerName -Counter (Convert-HString -HString $Counter)).counterSamples | ForEach-Object {
        $path = $_.path
        $PropertyHash=@{
                computerName=($Path -split "\\")[2];
                #WholeCounter = ($path  -split "\\")[-2,-1] -join "-";
                Instance = $_.InstanceName ;
                Value = [Math]::Round($_.CookedValue,2) 
                datetime=(Get-Date -format "yyyy-MM-d hh:mm:ss")
        }

        If (($path  -split "\\")[3] -eq $null -or ($path  -split "\\")[3] -eq "") { 
            $PropertyHash.Add('CounterCategory',$(($path  -split "\\")[4]))
            $PropertyHash.Add('CounterName',$(($path  -split "\\")[5]))
        } Else {
            $PropertyHash.Add('CounterCategory',$(($path  -split "\\")[3]))
            $PropertyHash.Add('CounterName',$(($path  -split "\\")[4]))
        }

New-Object PSObject -Property $PropertyHash
    }
}

function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

#endregion functions region
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
If (IsEmpty $OutputFile){$OutputFile = $OutputReport}

If (!(Test-Path $ServersTXTfile)){
    $MsgErrFileNotFound = "The file $ServersTXTfile is incorrect or doesn't exist ... `nDo you want to gather counters from the local machine ? (Y/N)"
    while ($Answer -ne "Y" -AND $Answer -ne "N") {
        cls
        Write-Host $MsgErrFileNotFound -BackgroundColor Yellow -ForegroundColor Red
        $Answer = Read-host
        If($Answer -eq "N"){Exit} Else {$Servers = $($Env:COMPUTERNAME)}
    }
} Else {
    [string[]]$servers = get-content $ServersTXTFile
}

Write-Host "Gathering performance counters for $($Servers -Join ", ")"
Write-Host "That's a total of $($Servers.count) servers"

#Collecting counter information for target servers
For ($ReRun = 1;$ReRun -le $NumberOfSamples;$ReRun ++){
    Write-Progress -Id 1 -Activity "Gathering $NumberOfSamples counters" -Status "Sample $ReRun of $NumberOfSamples" -PercentComplete ($($rerun/$NumberOfSamples*100))
    Get-CounterStats -ComputerName $Servers |Select-Object computerName,datetime,CounterCategory,CounterName,Instance,Value | Export-Csv -Path $OutputFile -Append -NoTypeInformation
}

Write-Host "File exported : $outputFile"
notepad $OutputFile

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>
$OutputFile = $null

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
