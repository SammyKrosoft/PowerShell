<#
.SYNOPSIS
    Imports mailbox Full Access, Send As and Send On Behalf of permissions
    from a CSV file.

.DESCRIPTION
    This script Imports mailbox Full Access, Send As and Send On Behalf of
     permissions from a CSV file. The input CSV file must have the following headers :
     
     "DisplayName","PrimarySMTPAddress","SendAsPermissions","FullAccessPermissions","SendOnBehalfPermissions"
     
     The "DisplayName" header is optional, we just need to be able to idendify the 
     mailbox into which we need to assign Full Access and/or Send-As and/or Send On Behalf 
     permissions. Usually the SMTP address is enough to uniquely identify a mailbox, but
     since the script uses standard Exchange Management Shell cmdlets, we can also use any
     other value that uniquely identifies a mailbox.

     As per Get-Mailbox help, you can use any value that uniquely identifies the mailbox.
        For example:
            Name
            Display name
            Alias
            Distinguished name (DN)
            Canonical DN
            <domain name>\<account name>
            Email address
            GUID
            LegacyExchangeDN
            SamAccountName
            User ID or user principal name (UPN)


.PARAMETER InputFile
    Choose the CSV file form you want to import the permissions map from.
    This CSV file must have the following headers:

    "DisplayName","PrimarySMTPAddress","SendAsPermissions","FullAccessPermissions","SendOnBehalfPermissions"

.PARAMETER CheckVersion
    This is just to check the script version.

.INPUTS
    You must specify an input CSV file (see InputFile parameter)

.OUTPUTS
    Script Log ...

.EXAMPLE
    .\Import-MailboxFASAPermissions.ps1 -InputFile C:\temp\ContosoPermissionsMAP.csv
    Will parse the ContosoPermissionsMAP.csv file and apply the permissions defined
    in the file to the mailboxes defined in this input CSV file.

.EXAMPLE
    .\Launch-YourScript.ps1 -CheckVersion
    Like in all my other scripts, this only dumps the script's version

.NOTES
    None

.LINK
    Get-Mailbox

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][string]$InputFile=".\sample.csv",
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
If ($CheckVersion) {Write-Host "SCRIPT NAME :$ScriptName `nSCRIPT VERSION :$ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$($ScriptPath)\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$($ScriptPath)\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
[array]$CSVFileRequiredHeaders = "PrimarySMTPAddress", "SendAsPErmissions", "FullAccessPermissions", "SendOnBehalfPermissions"

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
#region Functions
function Write-Log {
    <# 
     .SYNOPSIS
      Function to log input string to file and display it to screen
    
     .DESCRIPTION
      Function to log input string to file and display it to screen. Log entries in the log file are time stamped. Function allows for displaying text to screen in different colors.
    
     .PARAMETER String
      The string to be displayed to the screen and saved to the log file
    
     .PARAMETER Color
      The color in which to display the input string on the screen
      Default is White
      Valid options are
        Black
        Blue
        Cyan
        DarkBlue
        DarkCyan
        DarkGray
        DarkGreen
        DarkMagenta
        DarkRed
        DarkYellow
        Gray
        Green
        Magenta
        Red
        White
        Yellow
    
     .PARAMETER LogFile
      Path to the file where the input string should be saved.
      Example: c:\log.txt
      If absent, the input string will be displayed to the screen only and not saved to log file
    
     .EXAMPLE
      Write-Log -String "Hello World" -Color Yellow -LogFile c:\log.txt
      This example displays the "Hello World" string to the console in yellow, and adds it as a new line to the file c:\log.txt
      If c:\log.txt does not exist it will be created.
      Log entries in the log file are time stamped. Sample output:
        2014.08.06 06:52:17 AM: Hello World
    
     .EXAMPLE
      Write-Log "$((Get-Location).Path)" Cyan
      This example displays current path in Cyan, and does not log the displayed text to log file.
    
     .EXAMPLE 
      "$((Get-Process | select -First 1).name) process ID is $((Get-Process | select -First 1).id)" | Write-Log -color DarkYellow
      Sample output of this example:
        "MDM process ID is 4492" in dark yellow
    
     .EXAMPLE
      Write-Log 'Found',(Get-ChildItem -Path .\ -File).Count,'files in folder',(Get-Item .\).FullName Green,Yellow,Green,Cyan .\mylog.txt
      Sample output will look like:
        Found 520 files in folder D:\Sandbox - and will have the listed foreground colors
    
     .LINK
      https://superwidgets.wordpress.com/2014/12/01/powershell-script-function-to-display-text-to-the-console-in-several-colors-and-save-it-to-log-with-timedate-stamp/
    
     .NOTES
      Function by Sam Boutros
      v1.0 - 08/06/2014
      v1.1 - 12/01/2014 - added multi-color display in the same line
      v1.2 - 8 August 2016 - updated date time stamp format, protect against bad LogFile name
    
    #>
    
        [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
        Param(
            [Parameter(Mandatory=$true,
                       ValueFromPipeLine=$true,
                       ValueFromPipeLineByPropertyName=$true,
                       Position=0)]
                [String[]]$String, 
            [Parameter(Mandatory=$false,
                       Position=1)]
                [ValidateSet("Black","Blue","Cyan","DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","Red","White","Yellow")]
                [String[]]$Color = "Green", 
            [Parameter(Mandatory=$false,
                       Position=2)]
                [String]$LogFile = $ScriptLog,
            [Parameter(Mandatory=$false,
                       Position=3)]
                [Switch]$NoNewLine
        )
    
    
        $LegalFileNameCharSet = "^[" + [Regex]::Escape("A-Za-z0-9^&'@{}[],$=!-#()%.+~_") + "]+$"
        if ($String.Count -gt 1) {
            $i=0
            foreach ($item in $String) {
                if ($Color[$i]) { $col = $Color[$i] } else { $col = "White" }
                Write-Host "$item " -ForegroundColor $col -NoNewline
                $i++
            }
            if (-not ($NoNewLine)) { Write-Host " " }
        } else { 
            if ($NoNewLine) { Write-Host $String -ForegroundColor $Color[0] -NoNewline }
                else { Write-Host $String -ForegroundColor $Color[0] }
        }
    
        if ($LogFile.Length -gt 2 -and !($LogFile -match $LegalFileNameCharSet)) {
            "$(Get-Date -format 'dd MMMM yyyy hh:mm:ss tt'): $($String -join " ")" | Out-File -Filepath $Logfile -Append 
        } else {
            Write-debug "Log: Missing -LogFile parameter or bad LogFile name. Will not save input string to log file.."
        }
    }

function _Progress {
    param(
        [parameter(position = 0)] $Id = 1,
        [parameter(position = 1)] $PercentComplete=100,
        [parameter(position = 2)] $Activity = "Working...",
        [parameter(position = 3)] $Status="In Progress..."
        )

    Write-Progress -id $Id -activity $Activity -status $Status -PercentComplete ($PercentComplete)
    }

Function Test-ExchTools(){
    Try
    {
        Get-command Get-mailbox -ErrorAction Stop
        $ExchInstalledStatus = $true
        $Message = "Exchange tools are present !"
        Write-Host $Message -ForegroundColor Blue -BackgroundColor Red
    }
    Catch [System.SystemException]
    {
        $ExchInstalledStatus = $false
        $Message = "Exchange Tools are not present !"
        Write-Host $Message -ForegroundColor red -BackgroundColor Blue
        Exit
    }
    Return $ExchInstalledStatus
}
    
function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}


    <#
    .EXAMPLE
    ValidateHeadersFromCSV -FilePathAndName ".\sample.csv" -CSVFilerequiredHeaders "PrimarySMTPAddress", "SendAsPErmissions", "FullAccessPermissions", "SendOnBehalfPermissions"
    #>

    Function ValidateHeadersFromCSV {
        Param(
            [Parameter(Mandatory=$true, Position = 0, ParameterSetName = "NormalRun")][string]$FilePathAndName,
            [Parameter(Mandatory =$true, Position = 1, ParameterSetName = "NormalRun")][string[]]$CSVFilerequiredHeaders
        )
        $DuplicateHeader = $false
        $MissingHeader = $false
        If (!(Test-Path $FilePathAndName)){
            $MsgErrFileNotFound = "The file $InputFile is incorrect or doesn't exist ... please try again with another file or the correct path."
            Write-Host $MsgErrFileNotFound -BackgroundColor Yellow -ForegroundColor Red
            Return $false
        } Else {
            #Get the first line of the CSV file => THIS is what we will validate
            [string[]]$HeadersFromFile = (Get-content -Path $FilePathAndName | Select -first 1).Split(",")
            $HeadersFromFile = $HeadersFromFile.TrimStart()
            # $CSVFilerequiredHeaders
            # $CSVFilerequiredHeaders.count
            # $HeadersFromFile;
            # $HeadersFromFile.count
    
            # exit
            #Putting message in a variable to do localization in the future (French + English)
            $msgValidatingCSVHeader = "Validating the CSV headers..."
            Write-host $msgValidatingCSVHeader -BackgroundColor yellow -ForegroundColor blue
            #Parsing CSV required file headers coming from parameter
            #3 cases : 1 matching header in the file for each required header, 1 of the headers is missing, or we have duplicate headers 
            Foreach ($RequiredHeader in $CSVFilerequiredHeaders) {
                Write-Host "Checking $RequiredHeader" -BackgroundColor Green
                #Header counter to identify whether we have no matches, one match, or more than one
                $HeaderMatch = 0
                #We compare each CSV required header with each header of the file -> 3 cases : 1 match (wanted), 0 matches (CSV file not valid) or more than 1 matches (duplicates in CSV headers, CSV File not valid) 
                Foreach ($FileHeader in $HeadersFromFile){
                    if($($FileHeader) -eq $RequiredHeader -or $($FileHeader) -eq """$RequiredHeader"""){$HeaderMatch++}
                }
                If ($HeaderMatch -eq 1){
                    $msgFound1Match = "Ok"
                    Write-Host $msgFound1Match -BackgroundColor green -ForegroundColor black
                } ElseIf($headerMatch -eq 0) {
                    $msgErrMissingRequiredHeader = "$RequiredHeader not found in file or there are trailing space characters after $RequiredHeader! Please correct your CSV or select another CSV file."
                    Write-Host $msgErrMissingRequiredHeader -ForegroundColor Red
                    $MissingHeader = $true
                    [array]$MissingHeaderDetails += $RequiredHeader
                } Else {
                    $msgErrDuplicateRequiredHeader = "Cannot have more than 1 header named $RequiredHeader - please correct your CSV or select another CSV."
                    Write-Host  $msgErrDuplicateRequiredHeader -ForegroundColor Red
                    $DuplicateHeader = $true
                    [array]$DuplicateHeaderDetails += $RequiredHeader
                }
            }
        }
        If ($Missingheader -or $DuplicateHeader){
            If ($MissingHeader){
                $msgMissingHeader = "Missing Headers in file or space characters after Headers in the file: $($MissingHeaderDetails -join ", "), please use a CSV file with proper headers"
                Write-Host $msgMissingHeader -BackgroundColor yellow -ForegroundColor red
            }
            If ($DuplicateHeader){
                $msgDuplicateHeader = "Duplicate Headers in file : $($DuplicateHeaderDetails -join ", "), please use a CSV file with proper headers"
                Write-Host $msgDuplicateHeader -BackgroundColor yellow -ForegroundColor red
            }
            return $False
        }Else{
            Return $True
        }
    }

<# /FUNCTIONS #>
#endregion Functions
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>

If (!(Test-Path $InputFile)){
    $MsgErrFileNotFound = "The file $InputFile is incorrect or doesn't exist ... please try again with another file or the correct path."
    Write-Host $MsgErrFileNotFound -BackgroundColor Yellow -ForegroundColor Red
    Exit
} Else {
    If(!(ValidateHeadersFromCSV -FilePathAndName $InputFile -CSVFilerequiredHeaders $CSVFileRequiredHeaders)){exit}
    $PermissionsMAP = Import-CSV $InputFile
}



Foreach ($Item in $PermissionsMAP) {
    If(!(Isempty $Item.PrimarySMTPAddress)){
        $strCmd1 = "$LoadCurrentMAilbox = Get-Mailbox  $item.PrimarySMTPAddress"
        Write-Host $strCmd1
    }
    If(!(Isempty $Item.SendAsPermissions)){
        $ListOfSendAsTemp = $Item.SendAsPermissions -split ";"
        $ListOfSendAsTemp
        Foreach ($item in $ListOfSendAsTemp){
            $strCmd2 = "Get-Mailbox $Item"
        }
        #
        # Process permissions : for that mailbox (Get-Mailbox )
        #

    }
    If(!(Isempty $Item.FullAccessPermissions)){
        $ListOfFullAccessTemp = $Item.FullAccessPermissions -split ";"
        $ListOfFullAccessTemp
    }
    If(!(Isempty $Item.SendOnBehalfPermissions)){
        $ListOfSendOnBehalfTemp = $Item.SendOnBehalfPermissions -split ";"
        $ListOfSendOnBehalfTemp
    }

}


<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>
$PermissionsMAP = $null
<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
