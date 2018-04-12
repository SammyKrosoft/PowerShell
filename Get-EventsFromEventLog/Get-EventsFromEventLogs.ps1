<#
.SYNOPSIS
    Get specific events from any computer, local or remote, or computer list.

.DESCRIPTION
    This script gathers events from a filter defined in a Hash Table stored in the $FilterHashProperties variable, and prints it out to the screen.

.PARAMETER FirstNumber
    This parameter does blablabla

.PARAMETER SecondNumber
    This parameter does blablabla

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    None for now

.EXAMPLE
    Add default numbers 1 + 2
C:\PS> .\Add-Numbers.ps1
3

.EXAMPLE
    Add 14 with 23
C:\PS> .\Add-Numbers.ps1 -FirstNumber 14 -SecondNumber 23
37

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
Param(
    [Parameter(Mandatory = $False, Position = 1)] $EventLogName = {'Application', 'System'},
    [Parameter(Mandatory = $False, Position = 2)] $EventToCheck
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
#$LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
#$LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

<# -------------------------- DECLARATIONS -------------------------- #>
$computers = @()
$FilterHashProperties = $null
$computer = $null
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
cls
Write-Host "Starting script..."

#$Computers = Get-ExchangeServer
#$Computers = "Server-01", "Server-02", "Server-03", "Server-04"
#$COmputers = Get-Content C:\temp\MyServersList.txt
$Computers = "127.0.0.1"

If ($EventToCheck -eq "None" -or $EventToCheck -eq "" -or $EventToCheck -eq $Null)
{
    $EventToCheck = Read-Host "Which eventID are you looking for ? "
}


$FilterHashProperties = @{
    LogName = $EventLogName;
    #LogName = 'Application', 'System';
    #LogName = 'system';
    ID      = $EventToCheck;
}

Foreach ($computer in $computers)
{
    Write-host "Checking Computer $Computer" -BackgroundColor yellow -ForegroundColor Blue
    Try
    {
        $LastEvent = Get-WinEvent -ComputerName $Computer -Logname 'Application' -oldest -MaxEvents 1
        Write-host "Event logs on $computer goes as far as $($LastEvent.TimeCreated)"
        Try
        {
            $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -MaxEvents 80 -Computer $Computer -ErrorAction Stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ID, Message, ProviderName
            Write-host "Found at least $($Events.count) events ! Here are the 80 last ones :"
            $Events | Select -first 80 | ft -a
        }
        Catch
        {
            Write-Host "No such events with EventID = $($FilterHashProperties.ID) in the $($FilterHashProperties.LogName) event log on this computer..." -ForegroundColor Green
        }
        Finally
        {
            Write-Host "OK_"
        }
    }
    Catch
    {
        Write-Host "Error accessing Event Logs of $computer" -ForegroundColor Red
    }
}

<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>



