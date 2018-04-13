<#
.SYNOPSIS
    Get specific events from any computer, local or remote, or computer list.

.DESCRIPTION
    This script gathers events from a filter defined in a Hash Table stored in 
    the $FilterHashProperties variable, and prints it out to the screen.

.PARAMETER Computers
    By default, the script will search for events in the local computer (defined as 127.0.0.1).
    You can specify a remote computer (NOTE: you must have the Administrative rights in the remote machine), or
    you can also specify a list of computers, in the form of strings separated by commas like:
    -Computers "Server1", "Server2", "Server3", "Server4"

.PARAMETER EventLogName
    This parameter can be a string like Application, or an array of strings like ('Application', 'System') for example.
    By default, the EventLogName parameter is set to ('Application', 'System'), but you can specify "Application" to
    search on the Application Log only, or "System" to search in the System Log only, or "Security", etc...

.PARAMETER EventIDToCheck
    This parameter determines which Event number to check. It can be a single number, or an array of numbers.
    For a single number just type the event ID you're looking for, and for an array of numbers, type the numbers
    you want the script to search, separated by commas.

.PARAMETER NumberOfLastEventsToGet
    Indicates how many events you want the script to dump. 
    By default the script outputs the 30 last events that you searched for.

.INPUTS
    The name of the Event Log you want to search in (see EventLogName parameter) and the ID of the event you're looking
    for (see EventToCheck parameter)

.OUTPUTS
    Shows the events found on the console...

.EXAMPLE
    Search for 
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
    [Parameter(Mandatory = $False, Position = 1)] $Computers = "127.0.0.1",
    [Parameter(Mandatory = $False, Position = 2)] $EventLogName = ('Application', 'System'),
    [Parameter(Mandatory = $False, Position = 3)] $EventIDToCheck,
    [Parameter(Mandatory = $False, Position = 4)] $NumberOfLastEventsToGet = 30
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
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
cls
Write-Debug "Starting script..."

#$Computers = Get-ExchangeServer
#$Computers = "Server-01", "Server-02", "Server-03", "Server-04"
#$COmputers = Get-Content C:\temp\MyServersList.txt
#$Computers = "127.0.0.1"

If ($EventIDToCheck -eq "None" -or $EventIDToCheck -eq "" -or $EventIDToCheck -eq $Null)
{
    $EventIDToCheck = Read-Host "Which eventID are you looking for ? "
}


$FilterHashProperties = @{
    LogName = $EventLogName;
    ID      = $EventIDToCheck;
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
            $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -MaxEvents $NumberOfLastEventsToGet -Computer $Computer -ErrorAction Stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ID, Message, ProviderName
            Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones :"
            $Events | Select -first $NumberOfLastEventsToGet | ft -a
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



