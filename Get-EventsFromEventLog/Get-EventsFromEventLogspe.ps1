<#
.SYNOPSIS
    Searches and Get specific events from any computer, local or remote, or from a computer list.

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

.PARAMETER EventID
    This parameter determines which Event number to check. It can be a single number, or an array of numbers.
    For a single number just type the event ID you're looking for, and for an array of numbers, type the numbers
    you want the script to search, separated by commas.
    Example: Get-EventsFromEventLogs.ps1 -EventID 2121
    or Get-EventsFromEventLogs.ps1 -EventID 2121, 2242, 2080

.PARAMETER EventSource
    This parameter determine which Event source to search for. This is optional. To search for all events of type
    "Outlook" for example on your workstation, type -EventSource "Outlook"

.PARAMETER EventLevel
    With this additionnal parameter, you can filter your search on Event Logs on the Level of events:

    Name                           Value                                           
    ----                           -----                                          
    Verbose                        5                                              
    Informational                  4                                              
    Warning                        3                                              
    Error                          2                                              
    Critical                       1                                              
    LogAlways                      0        


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
    Shows the events found on the console and/or exported on a file.

.EXAMPLE
.\Get-EventsFromEventLogs.ps1
    Launching the script without options will :
    - Ask you which event(s) you wish to search for (separated by commas if you want multiple event IDs to search)
    - Search the local computer
    - Search the Application and System logs
    - Get 30 events of the type specified

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -Computers MyServers -EventLevel Error
    This will collect the Error events (the last 30 errors by default) from the computer named MyServers. 
    It won't store it into a file as we didn't call the "-ExportToFile" parameter, just dump into the screen
    to have an idea if your server is okay or if it's full of errors

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -Computers SRV-EX-01,SRV-EX-02,SRV-EX-03 -EventLevel Error -ExportToFile
    This will collect the Warning, Error, Critical events on computers SRV-EX-01, 02 and 03. The results
    will be dumped into a file labelled GetEventsFromEventLogs-Date-time.csv as we specified the
    ExportToFile parameter.
    Note that the computers list can come from a txt file as well (see next example)

    .EXAMPLE
.\Get-EventsFromEventsLogs.ps1 -Computers $(Get-Content .\ServersList.txt) -EventLevel Error,Critical -ExportToFile
    This will collect Error and Critical events on computers list defined in the "ServersList.txt" file on the current 
    directory from where you launched the script (.\ refers to the current user directory, NOT the directory where the
    script is) and store it into a file.

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -NumberOfLastEventsToGet 10 -EventID 916,105 -ExportToFile
    - Search for the 10 last events (-NumberOfLastEventsToGet 10) 
    - Search for event IDs 916 and 105
    - As no Event Log name (Application, System, Security, etc...) were specified, 
    the script will look inside the Application AND System logs by default.
    - We asked the script to look for Event IDs 916 and 105 (-EventID 916, 105)
    
    The exported file will be named GetEventsFromEventLogs_916-105_2018-04-13-10-01-55.csv
    as I ran the script on 13th April 2018 at 10h01 and 55 seconds in the morning.

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -NumberOfLastEventsToGet 30 -EventID 26 -EventLogName Application
    - Search for the last 30 events (-NumberOfLastEventsToGet 30)
    - Search for Event ID 26 only
    - Search in the Application Log only
    - We don't output any file, just print the results on the screen

.EXAMPLE 
.\Get-EventsFromEventLogs.ps1 -EventSource "Outlook"
    - Search all events generated by the "Outlook" application (all Event IDs, all Level (Info, Warning, etc...))
    - Search in Application and System (because I didn't specify which event log)
    - Search the last 30 events of type "Outlook" - if there are less, it will just print less
    - We don't output any file because I didn't specify the -ExportToFile parameter

MachineName     LogName         TimeCreated             LevelDisplayName    Id      Message
-----------     -------         -----------             ----------------    --      -------
12345678901     Application     4/13/2018 11:57:06 AM   Information          63     La demande de service web Exchange GetAppManifestssuccède à.</0w>
12345678901     Application     4/13/2018 7:57:00 AM    Information          63     La demande de service web Exchange GetAppManifestssuccède à.</0w>
12345678901     Application     4/13/2018 7:56:59 AM    Information          63     Outlook a détecté une notification de modification pour vos applications et va t...
12345678901     Application     4/13/2018 7:56:55 AM    Information          45     Outlook a chargé le(s) complément(s) suivant(s) :...

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -EventSource "disk","Outlook" -EventLevel Warning -NumberOfLastEventsToGet 1000
    - Search all events which source are "Disk" and "Outlook"
    - Search only "Warning" events of the above defined sources
    - All Event IDs of these (because I didn't specify any ID to filter)
    - Get the 1000 last events of the above criteria
    - didn't specify the -ExportToFile so will just display to screen

.EXAMPLE
.\Get-EventsFromEventLogs.ps1 -EventSource "disk" -NumberOfLastEventsToGet 1000 -EventLevel Critical,Warning,Error -ExportToFile
    - Search all events about the "disk"
    - Search only Critical, Warning and Error events
    - Search the 1000 last events about the above criteria
    - Export into a file (like GetEventsFromEventLogs_None_2018-04-14-04-34-27.csv)

.NOTES
    More examples to be documented as the script gain experience over the usage...

.COMPONENT
    Get-WinEvent cmdlet

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.diagnostics/get-winevent

.LINK
    https://github.com/SammyKrosoft

#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] $Computers = ("127.0.0.1"),
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][ValidateSet("Application","System","Security")] [array]$EventLogName = ('Application', 'System'),
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")] [array]$EventID="All",
    [Parameter(Mandatory = $False, Position = 4, ParameterSetName = "NormalRun")] [array]$EventSource="All",
    [Parameter(Mandatory = $False, Position = 5, ParameterSetName = "NormalRun")][ValidateSet("All","Information","Warning","Error","Critical", "Verbose")] [array]$EventLevel="All",
    [Parameter(Mandatory = $False, Position = 6, ParameterSetName = "NormalRun")] [int]$NumberOfLastEventsToGet = 30,
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] [Switch]$ExportToFile,
    [Parameter(Mandatory = $False, Position = 8, ParameterSetName = "NormalRun")] [Boolean]$Confirm = $true,
    [Parameter(Mandatory = $False, Position = 9, ParameterSetName = "NormalRun")] [switch]$DebugScript,
    [Parameter(Mandatory = $false, Position = 10, ParameterSetName = "CheckVersionOnly")][Switch]$CheckVersion
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
$ScriptVersion = "1.4.4spe"
<# Version changes :
v1.4.4spe -> Added speech
v1.4.4 -> corrected examples
v1.4.3 -> added a test on Powershell version (using $PSVersionTable) to check whether we can use
$PSSCriptRoot variable or $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition instead
1.4.1 -> 1.4.2
Added the CheckVersion switch
1.4 -> 1.4.1
Fixed file name generation when not supplying any EventID and no EventSource -> putting "Last_X_" inside the csv name
1.3 -> 1.4
Oddly, had some events with description = $null on my test machine => I had a stop error on replace carriage with # sign
and then added the condition $Event.Message -eq $null to not replace anything to fix this error
1.2.1 -> 1.3
simplified the script requirements : if no parameters specified, search and dump all events !
1.2 -> 1.2.1
replaced "None" with "All" when we don't specify a filter parameter (because when $EventSouce = nothing, we basically
search for all event sources)
#>
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
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
function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

Function IsPSV3 {
    <#
    .DESCRIPTION
    Just printing Powershell version and returning "true" if powershell version
    is Powershell v3 or more recent, and "false" if it's version 2.
    .OUTPUTS
    Returns $true or $false
    .EXAMPLE
    IsPSVersionV3
    #>
    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        Return $false
    } Else {
        Return $true
        }
}


Function Say {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    #$InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
    #$InstalledVoices
    #Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior')
    $Speak.SelectVoiceByHints(0,0,0,'en')
        $Speak.Speak($Msg)
}

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
cls
Write-Host "Starting script..."
$msg = "Welcome to the Events Collection Script !"
Say $msg
#$Computers = Get-ExchangeServer
#$Computers = "Server-01", "Server-02", "Server-03", "Server-04"
#$COmputers = Get-Content C:\temp\MyServersList.txt
#$Computers = "127.0.0.1"


while ($Answer -ne "Y" -AND $Answer -ne "N") {
    cls
    If (IsEmpty $EventSource -and IsEmpty $EventID -and IsEmpty $EventLevel){
    Write-host "No Event ID, Event Source or Event Level specified, we will search for all the last $NumberOfLastEventsToGet events on each machine`nor the local machine if you didn't specify the -Computers parameter" -BackgroundColor yellow -ForegroundColor blue
    }
    Write-Host "Event log names         :   $($EventLogName -join ", ")"
    Write-Host "Computers               :   $($Computers -join ", ")"
    Write-Host "Event ID to check       :   $($EventID -join ", ")"
    Write-Host "Event Source to check   :   $($EventSource -join ", ")"
    Write-Host "Event Level to check    :   $($EventLevel -join ", ")"
    Write-Host "Number of events to get :   $NumberOfLastEventsToGet"
If($ExportToFile){
    Write-Host "Write into a file       :   YES" -ForegroundColor yellow
} Else {
    Write-Host "Write into a file       :   NO" -ForegroundColor yellow
    }
    If ($Confirm){
        $msg = "Hit Yes or No to continue with the above options"
        Say $msg
        Write-Host "`nContinue (Y/N) ?" -BackgroundColor Red -ForegroundColor Blue
        $Answer = Read-host
    } Else {
        $Answer = "Y"
    }
    
    If ($Answer -eq "Y"){
        $msg = "Please wait while starting to collect the events"
        write-host $msg
        Say $msg
       
        $StopWatch.Reset()
        $StopWatch.start()
       }  
    IF ($Answer -eq "N"){
        Write-host "Exiting..."
        $msg = "Bye, come back soon!"
        Say $msg
        exit
    }
}

# Note for remember - sample of Hash table definition:
# $Hash= @{LogName='application'; ProviderName='outlook';Level = 2,4}
# Which gives:
# Name                           Value
# ----                           -----
# ProviderName                   outlook
# Level                          {2, 4}
# LogName                        application
#
# Then calling Get-WinEvent -FilterHash $Hash

$FilterHashProperties = @{
    LogName = $EventLogName
}

If (!(IsEmpty $EventSource)){
    $FilterHashProperties.Add('ProviderName',$EventSource)
}

If (!(IsEmpty $EventID)){
    $FilterHashProperties.Add("ID",$EventID)
}

If (!(IsEmpty $EventLevel)){
    for ($i=0;$i -lt $($EventLevel.count);$i++){
        $EventLevel[$i] = switch ($EventLevel[$i]) {
            "LogAlways" {0}
            "Critical" {1}
            "Error" {2}
            "Warning" {3}
            "Information" {4}
            "Verbose" {5}
        }
    }
    $FilterHashProperties.Add('Level',$EventLevel)
}

#Just adding a debug hard coded switch to check my filter...
if ($DebugScript){
    $FilterHashProperties
    $Computer = "127.0.0.1"
    $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -MaxEvents $NumberOfLastEventsToGet -Computer $Computer -ErrorAction stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ProviderName, ID, Message
    Foreach ($Event in $Events) {
        If ($Event.Message -ne $null){
            $Event.Message = $Event.Message.Replace("`r","#")
        }
    }
    Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones :"
    $Events | Select -first $NumberOfLastEventsToGet | ft -a
    exit
}

$msg = ("About to collect events on $($Computers.count)") + $(If ($($Computers.count) -gt 1){" machines"}Else{" machine"})
Write-host $msg
Say $msg

Foreach ($computer in $computers)
{
    $msg = "Checking Computer $Computer"
    Write-host $msg -BackgroundColor yellow -ForegroundColor Blue
    Say $msg
    Try
    {
        $LastEvent = Get-WinEvent -ComputerName $Computer -Logname 'Application' -oldest -MaxEvents 1
        Write-host "Event logs on $computer goes as far as $($LastEvent.TimeCreated)"
        Try
        {
            $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -MaxEvents $NumberOfLastEventsToGet -Computer $Computer -ErrorAction stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ProviderName, ID, Message
            Foreach ($Event in $Events) {
                If ($Event.Message -ne $null){
                    $Event.Message = $Event.Message.Replace("`r","#")
                }
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
        $msg = "Error accessing Event Logs of $computer"
        Write-Host $msg -ForegroundColor Red
        Say $msg
    }
}

$msg = "Found $($Events4all.count) Events in total ..."
Write-host $msg -BackgroundColor blue -ForegroundColor yellow
Say $msg
Write-host "Here are the stats by Event Level :"
$Summary = @()
$Summary = $Events4All | Group-Object LevelDisplayName | select @{Label="Event Level";Expression ={$_.Name}},@{Label = "Number of Events";Expression = {$_.Count}}
$Summary
$msg = "we found"
Say $msg
$Summary | Foreach {
    $msg = ($Summary."Number of Events") + (" events of type ") + ($Summary."Event Level")
    Write-host $msg}


If ($ExportToFile){
    If (!(IsEmpty $EventID)){
        $FileEventLogFirstID = "GetEventsFromEventLogs_$($EventID[0])_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
        If (IsPSV3){
            $EventsReport = "$PSScriptRoot\$FileEventLogFirstID"
        } Else {
            $EventsReport = "$(split-path -parent $MyInvocation.MyCommand.Definition)\$FileEventLogFirstID"
        }
    } Else { 
        If (!(IsEmpty $EventSource)){
            $FileEventLogFirstSource = "GetEventsFromEventLogs_$($EventSource[0])_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
            If (IsPSV3){
                $EventsReport = "$PSScriptRoot\$FileEventLogFirstSource"
            } Else {
                $EventsReport = "$(split-path -parent $MyInvocation.MyCommand.Definition)\$FileEventLogFirstSource"
            }
            
        } Else {
            $FileNumberOfLastEvents = "GetEventsFromEventLogs_Last_$($NumberOfLastEventsToGet)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
            If(IsPSV3){
                $EventsReport = "$PSScriptRoot\$FileNumberOfLastEvents"
            } else {
                $EventsReport = "$(split-path -parent $MyInvocation.MyCommand.Definition)\$FileNumberOfLastEvents"
            }
        }
    }
    $Events4all | Export-Csv -NoTypeInformation $EventsReport
    $msg = "I'm done. I exported the results to a file located on the same directory as the Script"
    Write-host $msg
    Say $msg
    notepad $EventsReport
} Else {
    $msg = "I'm done. No file exported because you didn't specify the ExportToFile script parameter."
    Write-Host $msg
    Say $msg
}
<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...)
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
Say $msg
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>