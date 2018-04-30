<#
.SYNOPSIS
Script: Get-CounterStats
Author: Prashanth and Praveen

This script will collect the specific counters value from the multiple target machines/servers 
which will be used to analayze the performance of target servers.

.DESCRIPTION
This script will collect the specific counters value from the multiple target machines/servers 
which will be used to analayze the performance of target servers.

.PARAMETER ServersTXTFile
    This parameter specified the file containing the list of servers to get Perfmon samples from.
    By default it will look for a "servers.txt" file in the same directory as the script.

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
Will execute and dump the counters stats on a list of servers defined in the C:\Temp\Servers.txt file.

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
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][int]$NumberOfSamples = 2,
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
If ($CheckVersion) {Write-Host "SCRIPT NAME :$ScriptName `nSCRIPT VERSION :$ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
#$ScriptLog = "$PSScriptRoot\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
#region Functions region
#Function to have the customized output in CSV format
function Export-CsvFile {
    [CmdletBinding(DefaultParameterSetName='Delimiter', SupportsShouldProcess = $true, ConfirmImpact='Medium')]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][System.Management.Automation.PSObject]${InputObject},
        [Parameter(Mandatory=$true, Position=0)][Alias('PSPath')][System.String]${Path},
        [Switch]${Append},
        [Switch]${Force},
        [Switch]${NoClobber},
        [ValidateSet('Unicode','UTF7','UTF8','ASCII','UTF32','BigEndianUnicode','Default','OEM')][System.String]${Encoding},
        [Parameter(ParameterSetName='Delimiter', Position=1)][ValidateNotNull()][System.Char]${Delimiter},
        [Parameter(ParameterSetName='UseCulture')][Switch]${UseCulture},
        [Alias('NTI')][Switch]${NoTypeInformation}
    )

    begin
    {
        # This variable will tell us whether we actually need to append
        # to existing file
        $AppendMode = $false
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
                {$PSBoundParameters['OutBuffer'] = 1}
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Csv',[System.Management.Automation.CommandTypes]::Cmdlet)
            #String variable to become the target command line
            $scriptCmdPipeline = ''
            # Add new parameter handling
            #Process and remove the Append parameter if it is present
            if ($Append) {
                $PSBoundParameters.Remove('Append') | Out-Null
                    if ($Path) {
                        if (Test-Path $Path) {        
                        # Need to construct new command line
                        $AppendMode = $true
                    if ($Encoding.Length -eq 0) {
                        # ASCII is default encoding for Export-CSV
                        $Encoding = 'ASCII'
                    }

            # For Append we use ConvertTo-CSV instead of Export
            $scriptCmdPipeline += 'ConvertTo-Csv -NoTypeInformation '

            # Inherit other CSV convertion parameters
            if ( $UseCulture ) {$scriptCmdPipeline += ' -UseCulture ' }
            if ( $Delimiter ) {$scriptCmdPipeline += " -Delimiter '$Delimiter' "} 

            # Skip the first line (the one with the property names) 
            $scriptCmdPipeline += ' | Foreach-Object {$start=$true}'
            $scriptCmdPipeline += '{if ($start) {$start=$false} else {$_}} '

            # Add file output
            $scriptCmdPipeline += " | Out-File -FilePath '$Path' -Encoding '$Encoding' -Append "
            
            if ($Force) {$scriptCmdPipeline += ' -Force'}
            if ($NoClobber) {$scriptCmdPipeline += ' -NoClobber'}   
                        }
                    }
            } 
    
        $scriptCmd = {& $wrappedCmd @PSBoundParameters }

        if ( $AppendMode ) {
            # redefine command line
            $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock($scriptCmdPipeline)
        } else {
            # execute Export-CSV as we got it because
            # either -Append is missing or file does not exist
            $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock([string]$scriptCmd)
        }

        # standard pipeline initialization
        $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
        $steppablePipeline.Begin($PSCmdlet)
        } catch {throw}
    }

    process
    {
        try {
            $steppablePipeline.Process($_)
        } catch {throw}
    }

    end
    {
        try {$steppablePipeline.End()} catch {throw}
    }

}

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
            {
                $ComputerName
            }    
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
        [String[]]$ComputerName = $ENV:ComputerName
    ) 

    $Object =@()

$Counter = @"
Processor(_total)\% processor time 
\MSExchange RpcClientAccess\RPC Averaged Latency
\MSExchange RpcClientAccess\RPC Requests
Memory\Available MBytes 
PhysicalDisk(*)\Avg. Disk sec/Transfer 
Network Interface(*)\Bytes Total/sec
"@ 


    (Get-Counter -ComputerName $ComputerName -Counter (Convert-HString -HString $Counter)).counterSamples | ForEach-Object {
        $path = $_.path

        $PropertyHash=@{
                computerName=($Path -split "\\")[2];
                WholeCounter = ($path  -split "\\")[-2,-1] -join "-";
                Item = $_.InstanceName ;
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
    $MsgErrFileNotFound = "The file $ServersTXTfile is incorrect or doesn't exist ... please try again with another file or the correct path."
    Write-Host $MsgErrFileNotFound -BackgroundColor Yellow -ForegroundColor Red
    Exit
} Else {
    [string[]]$servers = get-content $ServersTXTFile
}

Write-Host "Gathering performance counters for $($Servers -Join ", ")"
Write-Host "That's a total of $($Servers.count) servers"
#exit

#Collecting counter information for target servers
#foreach($server in $Servers){
For ($ReRun = 1;$ReRun -le $NumberOfSamples;$ReRun ++){
    Write-Progress -Id 1 -Activity "Gathering $NumberOfSamples counters" -Status "Sample $ReRun of $NumberOfSamples" -PercentComplete ($($rerun/$NumberOfSamples*100))
    Get-CounterStats -ComputerName $Servers |Select-Object computerName,WholeCounter,CounterCategory,CounterName,Item,Value,datetime | Export-Csv -Path $OutputFile -Append -NoTypeInformation
}

Write-Host "File exported : $outputFile"
notepad $OutputFile

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
