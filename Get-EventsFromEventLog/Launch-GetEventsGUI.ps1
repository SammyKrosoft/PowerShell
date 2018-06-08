<#
.NOTES
With the help of            :   Jim Moyle @jimmoyle
How-To GUI From Jim Moyle   :   https://github.com/JimMoyle/GUIDemo

#>
$global:GUIversion = "1.2"
<# Release notes
v1.2 -> changed way to call ShowDialog() to avoid crashes
v1.1.1 -> fixed lack of IsPSV3 function ...
#>
#========================================================
#region Functions definitions (NOT the WPF form events)
#========================================================

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
        Write-Host "Sorry, PowerShell v3 or more is required. Exiting."
        Return $false
        Exit
    } Else {
        Return $true
        }
}

Function Get-EventsFromEventLogs {
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
    $ScriptVersion = "1.4.4.1"
    <# Version changes :
    v1.4.4.1 -> update for the GUI version Get-Events function - added out-string to dump events on host
    Write-Host ($Events | Select -first $NumberOfLastEventsToGet | ft -a | out-string)
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
    If ($CheckVersion) {Return $ScriptVersion}
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
    <# /FUNCTIONS #>
    <# -------------------------- EXECUTIONS -------------------------- #>
    #cls
    Write-Host "Starting script..."

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
            Write-Host "`nContinue (Y/N) ?" -BackgroundColor Red -ForegroundColor Blue
            $Answer = Read-host
        } Else {
            $Answer = "Y"
        }
        If ($Answer -eq "Y"){$StopWatch.Reset();$StopWatch.start()} 
        IF ($Answer -eq "N"){exit}
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
        Write-Host ($Events | Select -first $NumberOfLastEventsToGet | ft -a | out-string)
        exit
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
                $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -MaxEvents $NumberOfLastEventsToGet -Computer $Computer -ErrorAction stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ProviderName, ID, Message
                Foreach ($Event in $Events) {
                    If ($Event.Message -ne $null){
                        $Event.Message = $Event.Message.Replace("`r","#")
                    }
                }
                Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones :"
                Write-Host ($Events | Select -first $NumberOfLastEventsToGet | ft -a | out-string)
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

    Write-host "Found $($Events4all.count) Events in total ..." -BackgroundColor blue -ForegroundColor yellow
    Write-host "Here are the stats by Event Level :"
    Write-host ($Events4All | Group-Object LevelDisplayName | ft @{Label="Event Level";Expression ={$_.Name}},@{Label = "Number of Events";Expression = {$_.Count}} | out-string)

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
        notepad $EventsReport
    }
    <# /EXECUTIONS #>
    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...)
    $stopwatch.Stop()
    Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

}

function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

Function Say {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    If ($wpf.chkSpeech.IsChecked -eq $true) {
        $InstalledVoices = @()
        Add-Type -AssemblyName System.Speech
        $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
        # $InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
        # $InstalledVoices
        # Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior',[int32]'Position which voices are ordered','fr/en')
        switch ($wpf.lstBoxLanguage.SElectedItem.Content) {
            "Francais" {$Language = 'fr'}
            "English" {$Language = 'en'}
            "" {$language = 'en'}
            $null {$Language = 'en'}
        }
        $Speak.SelectVoiceByHints(0,0,0,$language)
        $Speak.Speak($Msg)
    }
}

Function WritNSay ($msg) {
    Write-Host $msg
    Say $msg
}

Function Split-ListColon {
    param(
        [string]$StringToSplit,
        [switch]$Noquotes
    )
    $TargetSplit = $StringToSplit.Split(',')
    $ListItems = ""
    If ($NoQuotes){
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += $TargetSplit[$i].trim() + (", ")}
        $ListItems += $TargetSplit[$TargetSplit.Count - 1].trim()
    } Else {
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += ("""") + $TargetSplit[$i].trim() + (""", ")}
        $ListItems += ("""") + $TargetSplit[$TargetSplit.Count - 1].trim() + ("""")
    }
    Return $ListItems
}

Function Update-cmd{
    # [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] 
    # $Computers = ("127.0.0.1"),
    # [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][ValidateSet("Application","System","Security")] 
    # [array]$EventLogName = ('Application', 'System'),
    # [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")] 
    # [array]$EventID="All",
    # [Parameter(Mandatory = $False, Position = 4, ParameterSetName = "NormalRun")] [array]$EventSource="All",
    # [Parameter(Mandatory = $False, Position = 5, ParameterSetName = "NormalRun")][ValidateSet("All","Information","Warning","Error","Critical", "Verbose")] 
    # [array]$EventLevel="All",
    # [Parameter(Mandatory = $False, Position = 6, ParameterSetName = "NormalRun")] [int]$NumberOfLastEventsToGet = 30,
    # [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] [Switch]$ExportToFile,
    # [Parameter(Mandatory = $False, Position = 8, ParameterSetName = "NormalRun")] [Boolean]$Confirm = $true,
    # [Parameter(Mandatory = $False, Position = 9, ParameterSetName = "NormalRun")] [switch]$DebugScript,
    # [Parameter(Mandatory = $false, Position = 10, ParameterSetName = "CheckVersionOnly")][Switch]$CheckVersion

    $command = "Get-EventsFromEventLogs"
    If ($($wpf.txtCSVComputersList.Text) -ne ""){
        $TextBoxList = Split-ListColon -StringToSplit $wpf.txtCSVComputersList.Text
        $command += (" -Computers ") + ($TextBoxList)
    }

    If ($($wpf.chkAppLog.IsChecked) -or $($wpf.chkSystemLog.IsChecked) -or $($wpf.chkSecurityLog.IsChecked)){
        [string[]]$LogsToSearch = @()
        If($wpf.chkAppLog.IsChecked){$LogsToSearch += "Application"}
        If($wpf.chkSystemLog.IsChecked) {$LogsToSearch += "System"}
        If($wpf.chkSecurityLog.IsChecked) {$LogsToSearch += "Security"}
        $LogsToSearch = $LogsToSearch -join ", "
        $Command += (" -EventLogName ") + $LogsToSearch
    }

    If ($($wpf.chkLevelInformation.IsChecked) -or $($wpf.chkLevelWarning.IsChecked) -or $($wpf.chkLevelError.IsChecked) -or $($wpf.chkLevelCritical.IsChecked)){
        [string[]]$EventLevelToSearch = @()
        If ($wpf.chkLevelInformation.IsChecked){$EventLevelToSearch += "Information"}
        If ($wpf.chkLevelWarning.IsChecked){$EventLevelToSearch += "Warning"}
        If ($wpf.chkLevelError.IsChecked){$EventLevelToSearch += "Error"}
        If ($wpf.chkLevelCritical.IsChecked){$EventLevelToSearch += "Critical"}
        $EventLevelToSearch = $EventLevelToSearch -join ", "
        $command += (" -EventLevel ") + $EventLevelToSearch
    }

    If ($($wpf.chkSaveToFile.IsChecked)){
        $command += " -ExportToFile"
    }

    If ($($wpf.txtEventIDs.Text) -ne ""){
        $TextBoxList = Split-ListColon -StringToSplit $wpf.txtEventIDs.Text -NoQuotes
        $command += (" -EventID ") + ($TextBoxList)
    }

    If ($($wpf.txtEventSources.Text) -ne ""){
        $TextBoxList = Split-ListColon -StringToSplit $wpf.txtEventSources.Text
        $command += (" -EventSource ") + ($TextBoxList)
    }

    If (($($wpf.txtNumberOfEvents.Text) -ne "30") -and ($($wpf.txtNumberOfEvents.Text) -ne "")){
        $command += (" -NumberOfLastEventsToGet ") + ($wpf.txtNumberOfEvents.Text)
    }

    $command += " -Confirm `$false"
    # Populate the cmdlet text box that it's gonna use...
    $wpf.txtCommand.text = $command
}

#========================================================
# End of Functions definitions (note the WPF form events)
#endregion
#========================================================

#========================================================
#region WPF form definition and load controls
#========================================================

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
# $inputXML = Get-Content -Path "C:\Users\Kamehameha\Documents\GitHub\PowerShell\Get-EventsFromEventLog\VisualStudio2017WPFDesign\Launch-EventsCollector-WPF\Launch-EventsCollector-WPF\MainWindow.xaml"
$inputXML = @"

<Window x:Name="EventCollectWindow" x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="SearchAndCollect" Height="501.903" Width="800" ShowActivated="False">
    <Grid Margin="0,0,0,0">
        <Grid.Background>
            <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                <GradientStop Color="#FFC9D47E"/>
                <GradientStop Color="#FFEB5E5E" Offset="1"/>
            </LinearGradientBrush>
        </Grid.Background>
        <CheckBox x:Name="chkAppLog" Content="Application Log" HorizontalAlignment="Left" Margin="371,28,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtCSVComputersList" HorizontalAlignment="Left" Height="147" Margin="10,68,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="317"/>
        <CheckBox x:Name="chkSystemLog" Content="System Log" HorizontalAlignment="Left" Margin="371,48,0,0" VerticalAlignment="Top"/>
        <Label Content="Computers List (Comma Separated)" HorizontalAlignment="Left" Margin="10,42,0,0" VerticalAlignment="Top" Width="202"/>
        <CheckBox x:Name="chkLevelInformation" Content="Information" HorizontalAlignment="Left" Margin="498,28,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkLevelWarning" Content="Warning" HorizontalAlignment="Left" Margin="498,48,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkLevelError" Content="Error" HorizontalAlignment="Left" Margin="498,68,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkLevelCritical" Content="Critical" HorizontalAlignment="Left" Margin="498,88,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtNumberOfEvents" HorizontalAlignment="Left" Height="35" Margin="332,180,0,0" TextWrapping="Wrap" Text="30" VerticalAlignment="Top" Width="104"/>
        <TextBlock HorizontalAlignment="Left" Margin="332,143,0,0" TextWrapping="Wrap" Text="Events to collect per computer" VerticalAlignment="Top" Width="104"/>
        <TextBox x:Name="txtCommand" HorizontalAlignment="Left" Height="91" Margin="10,338,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="775" IsReadOnly="True"/>
        <Label Content="Function Command Line we'll launch" HorizontalAlignment="Left" Margin="10,307,0,0" VerticalAlignment="Top" Width="240"/>
        <Button x:Name="btnRun" Content="Run" HorizontalAlignment="Left" Height="30" Margin="149,434,0,0" VerticalAlignment="Top" Width="161"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" Margin="471,434,0,0" VerticalAlignment="Top" Width="160" Height="30"/>
        <CheckBox x:Name="chkSpeech" Content="Speech" HorizontalAlignment="Left" Margin="681,28,0,0" VerticalAlignment="Top"/>
        <ListBox x:Name="lstBoxLanguage" HorizontalAlignment="Left" Height="47" Margin="681,48,0,0" VerticalAlignment="Top" Width="70" IsSynchronizedWithCurrentItem="False" IsEnabled="False" SelectedIndex="1">
            <ListBoxItem Content="Francais"/>
            <ListBoxItem Content="English"/>
        </ListBox>
        <CheckBox x:Name="chkSecurityLog" Content="Security Log" HorizontalAlignment="Left" Margin="371,68,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtEventIDs" HorizontalAlignment="Left" Height="32" Margin="483,143,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302"/>
        <TextBox x:Name="txtEventSources" HorizontalAlignment="Left" Height="62" Margin="483,206,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302"/>
        <Label Content="Event IDs to look for (comma separated)" HorizontalAlignment="Left" Margin="483,117,0,0" VerticalAlignment="Top" Width="302"/>
        <Label Content="Event Sources to look for (comma separated)" HorizontalAlignment="Left" Margin="483,180,0,0" VerticalAlignment="Top" Width="302"/>
        <CheckBox x:Name="chkSaveToFile" Content="Save events to file" HorizontalAlignment="Left" Margin="483,295,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblGUIVer" Content="GUI version" HorizontalAlignment="Left" Margin="10,242,0,0" VerticalAlignment="Top" Background="#FFC1B621"/>
        <Label x:Name="lblFUNCVer" Content="Event collector function version" HorizontalAlignment="Left" Margin="10,268,0,0" VerticalAlignment="Top" Background="#FF66D71F"/>

    </Grid>
</Window>

"@
$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#========================================================
# END of WPF form definition and load controls
#endregion
#========================================================

#========================================================
#region WPF EVENTS definition
#========================================================

#region Buttons
$wpf.btnRun.add_Click({
    $msg = "Running the command"
    WritNSay $msg
    Invoke-expression $wpf.txtCommand.text
})

$wpf.btnCancel.add_Click({
    $msg = "Exiting..."
    WritNSay $msg
    $wpf.EventCollectWindow.Close()
})
# End of Buttons region
#endregion

#region Speech management region
$wpf.chkSpeech.add_Checked({
    $wpf.lstBoxLanguage.Isenabled = $true
    If ($($wpf.lstBoxLanguage.SelectedItem.content) -eq "Francais") {
        $msg = "Narrateur activé - merci de m'enlever mon baillon !"
    } Else {
        $msg = "Narrator activated - thanks for unmuting me !"
    }
    WritNsay $msg
})

$wpf.chkSpeech.add_UnChecked({
    $wpf.lstBoxLanguage.Isenabled = $false
})

$wpf.lstBoxLanguage.add_SelectionChanged({
    $msg = "Language = $($wpf.lstBoxLanguage.SelectedItem.content)"
    If ($($wpf.lstBoxLanguage.SelectedItem.content) -eq "Francais") {
        $msg = "Langue Francaise sélectionnée !"
    } Else {
        $msg = "English Language selected !"
    }
    WritNsay $msg})

# End of speech management region
#endregion

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.EventCollectWindow.Add_Loaded({
    $wpf.lblGUIVer.content += $global:GUIversion
    $wpf.lblFUNCVer.content += (" ") + (Get-EventsFromEventLogs -CheckVersion)
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.EventCollectWindow.Add_ContentRendered({
    Update-cmd
})
$wpf.EventCollectWindow.add_Closing({
    $msg = "bye bye !"
    WritNSay $msg
})
# End of load, draw and closing form events
#endregion

#region Text Changed events
$wpf.txtCSVComputersList.add_TextChanged({
    Update-cmd
})
$wpf.txtNumberOfEvents.add_TextChanged({
    Update-cmd
})
$wpf.txtEventIDs.add_TextChanged({
    Update-cmd
})
$wpf.txtEventSources.add_TextChanged({
    Update-cmd
})


#End of Text Changed events
#endregion

#region Clicked on Checkboxes events
$wpf.chkAppLog.add_Click({
    Say "Application Log"
    Update-cmd
})
$wpf.chkSystemLog.add_Click({
    Say "System Log"
    Update-cmd
})
$wpf.chkSecurityLog.add_Click({
    Say "Security Log"    
    Update-cmd
})
$wpf.chkLevelInformation.add_Click({
    Say "Information events"
    Update-cmd
})
$wpf.chkLevelWarning.add_Click({
    Say "Warning events"
    Update-cmd
})
$wpf.chkLevelError.add_Click({
    Say "Error events"
    Update-cmd
})
$wpf.chkLevelCritical.add_Click({
    Say "Critical events"
    Update-cmd
})
$wpf.chkSaveToFile.add_Click({
    Say "Save to file"
    Update-cmd
})
# End of Clicked on Checkboxes events
#endregion

#=======================================================
#End of Events from the WPF form
#endregion
#=======================================================

IsPSV3 | out-null

# Load the form:
# Older way >>>>> $wpf.EventCollectWindow.ShowDialog() | Out-Null >>>> generates crash if run multiple times
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.EventCollectWindow.Dispatcher.InvokeAsync({
    $wpf.EventCollectWindow.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null
