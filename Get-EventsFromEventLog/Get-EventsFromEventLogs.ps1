<#
.SYNOPSIS
    Get specific events from any computer, local or remote, or from a computer list.

.DESCRIPTION
    This script gathers events from a computer or a list of computers, from
    the Application, System or Security or all of these Event logs types.
    You just have to specify which event ID or IDs (e.g. 105, 1020, 67 or just 105
    or any number), and spit the events list on the screen.
    
    By default, if no computers are specified, the script will search on the local
    computer.

.PARAMETER Computers
    By default, the script will search for events in the local computer (defined as 127.0.0.1).
    You can specify a remote computer (NOTE: you must have the Administrative rights in the remote machine), or
    you can also specify a list of computers, in the form of strings separated by commas like:
    -Computers "Server1", "Server2", "Server3", "Server4"

.PARAMETER EventLogName
    Which Event Log to look at => default will look on Application and System logs.
    This parameter can be a string like Application, or an array of strings like ('Application', 'System') for example.
    By default, the EventLogName parameter is set to ('Application', 'System'), but you can specify "Application" to
    search on the Application Log only, or "System" to search in the System Log only, or "Security", etc...

.PARAMETER EventIDToCheck
    This parameter determines which Event number to check. It can be a single number, or an array of numbers.
    For a single number just type the event ID you're looking for, and for an array of numbers, type the numbers
    you want the script to search, separated by commas.
    Example: Get-EventsFromEventLogs.ps1 -EventIDToCheck 2121
    or Get-EventsFromEventLogs.ps1 -EventIDToCheck 2121, 2242, 2080

.PARAMETER NumberOfLastEventsToGet
    Indicates how many events you want the script to dump. 
    By default the script outputs the 30 last events that you searched for.
    If there are less than 30 events (or the number you specified), it will dump all the existing events, 
    which can be less than 30 (or the number you specified)

.PARAMETER ExportToFile
    This is a SWITCH that, if specified, will store the results in a CSV file.
    This file will be placed on the directory where the script is located, and named :
        "GetEventsFromEventLogs_EventID1-EventID2-XXXXX_Year-MONTH-DAY-Hour-Minute-Second.csv"
    Example:
        "GetEventsFromEventLogs_916-105_2018-04-13-09-52-08.csv"
    AND it will be opened automatically in notepad once the script finishes.
    
.INPUTS
    The name of the Event Log you want to search in (see EventLogName parameter) and the ID of the event you're looking
    for (see EventToCheck parameter)

.OUTPUTS
    Shows the events found on the console...

.EXAMPLE
.\Get-EventsFromEventLogs.ps1
    Launching the script without options will :
    - Ask you which event(s) you wish to search for (separated by commas if you want multiple event IDs to search)
    - Search the local computer
    - Search the Application and System logs
    - Get 30 events of the type specified

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -NumberOfLastEventsToGet 10 -EventIDToCheck 916,105 -ExportToFile
    - Search for the 10 last events (-NumberOfLastEventsToGet 10) 
    - Search for event IDs 916 and 105
    - As no Event Log name (Application, System, Security, etc...) were specified, 
    the script will look inside the Application AND System logs by default.
    - We asked the script to look for Event IDs 916 and 105 (-EventIDTocheck 916, 105)
    
    The exported file will be named GetEventsFromEventLogs_916-105_2018-04-13-10-01-55.csv
    as I ran the script on 13th April 2018 at 10h01 and 55 seconds in the morning.

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -NumberOfLastEventsToGet 30 -EventIDToCheck 26 -EventLogName Application
    - Search for the last 30 events (-NumberOfLastEventsToGet 30)
    - Search for Event ID 26 only
    - Search in the Application Log only
    - We don't output any file, just print the results on the screen

.NOTES
    More examples to be documented as the script gain experience over the usage...

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
Param(
    [Parameter(Mandatory = $False, Position = 1)] $Computers = ("127.0.0.1"),
    [Parameter(Mandatory = $False, Position = 2)] $EventLogName = ('Application', 'System'),
    [Parameter(Mandatory = $False, Position = 3)] $EventIDToCheck,
    [Parameter(Mandatory = $False, Position = 4)] [int]$NumberOfLastEventsToGet = 30,
    [Parameter(Mandatory = $False, Position = 4)] [Switch]$ExportToFile
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
# $EventsReport = "$PSScriptRoot\GetEventsFromEventLogs_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
#$LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

<# -------------------------- DECLARATIONS -------------------------- #>
$FilterHashProperties = $null
$Answer = ""
$Events4All = @()
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
#cls
Write-Host "Starting script..."

#$Computers = Get-ExchangeServer
#$Computers = "Server-01", "Server-02", "Server-03", "Server-04"
#$COmputers = Get-Content C:\temp\MyServersList.txt
#$Computers = "127.0.0.1"

While ($EventIDToCheck -eq "None" -or $EventIDToCheck -eq "" -or $EventIDToCheck -eq $Null -or $EventIDToCheck -eq 0)
{
    $EventIDToCheck = Read-Host "Which eventID are you looking for ? "
    If ($EventIDToCheck -eq "None" -or $EventIDToCheck -eq "" -or $EventIDToCheck -eq $Null -or $EventIDToCheck -eq 0)
        {Write-Host "Invalid value - please enter an integer or a list of integers comma separated like 2121, 2242, 2080..." -BackgroundColor Blue -ForegroundColor Yellow}
}

while ($Answer -ne "Y" -AND $Answer -ne "N") {
    cls
    Write-Host "Event log names         :   $EventLogName"
    Write-Host "Computers               :   $Computers"
    Write-Host "Event ID to check       :   $EventIDToCheck"
    Write-Host "Number of events to get :   $NumberOfLastEventsToGet"
If($ExportToFile){
    Write-Host "Write into a file       :   YES" -ForegroundColor yellow
    } Else {
    Write-Host "Write into a file       :   NO" -ForegroundColor yellow
    }
    Write-Host "`nContinue (Y/N) ?" -BackgroundColor Red -ForegroundColor Blue
    $Answer = Read-host
    If ($Answer -eq "Y"){$StopWatch.Reset();$StopWatch.start()} 
    IF ($Answer -eq "N"){exit}
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
            Foreach ($Event in $Events) {
                $Event.Message = $Event.Message.Replace("`r","#")
            }
            Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones :"
            $Events | Select -first $NumberOfLastEventsToGet | ft -a
            $Events4All += $Events
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

If ($ExportToFile){
    $EventsReport = "$PSScriptRoot\GetEventsFromEventLogs_$($EventIDToCheck -join "-")_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
    $Events4all | Export-Csv -NoTypeInformation $EventsReport
    notepad $EventsReport
}

<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...)
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>



