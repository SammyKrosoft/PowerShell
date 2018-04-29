<#
.SYNOPSIS
    Export Exchange Mailbox Send As, Full Access, and Send On Behalf permissions
    in a CSV file in order to later import them in another environment using the 
    output CSV file.

.DESCRIPTION
    Export following Exchange Mailbox permissions in a CSV file 
    - Send As
    - Full Access
    - Send On Behalf To
    in order to be able to import them later in another environment using 
    the output CSV file.

    To import back the permissions, use the associated Import-MailboxFASAPermissions.ps1 
    script.

    Since the Send As and Full Access permissions can be granted to non-mailbox or
    non-mail enabled users, these are stored in the CSV in the form of DOMAIN\Alias.
    On the other hand, the Send On Behalf permission can be granted to mailbox-enabled users
    or mail-enabled users or security groups only, but for some reason it is stored in the form of 
    DOMAIN\OU1\Sub-OU1\Sub-OU2\Name - then, the script is designed to convert these to 
    PrimarySMTPAddress of these users => that way :
    > Not only we are sure that each entry represents a unique user
    > But also it will be easier for the IMPORT script to import these permissions back.

.PARAMETER OutputFile
    Sets the file to which we want to store the results.
    By default, the script will generate a CSV report with the name of the script, 
    with the date and time appended to it.

.PARAMETER SharedMailboxes
    This indicates the script 

.PARAMETER ResourceMailboxes

.PARAMETER CheckVersion
    This parameter just dumps the script version.

.INPUTS
    The script will scan all the mailboxes, but database by database to avoid to use
    all the RAM of the machine from which it's executed. 

.OUTPUTS
    A CSV file with either a name that you specify with the OutputFile parameter, or if not,
    the name of the script, containing the users Display Names, primary SMTP addresses,
    and the list of Send-As, Full Access and SendOnBehalfTo for each of these mailboxes.
    
    If the Send-As, Full Access and SendOnBehalfTo are multi-values, they are stored in the columns
    as semi-colon separated values, like Value1;value2;value3;...
    
    => when processing each permissions set, just use something like $ImportedCSV.SendAsPermissions -split ";" 
    or $ImportedCSV.SendAsPermissions.Split(";") ... 

.EXAMPLE
.\Export-MailboxFASAPermissions.ps1
    Will run the script and export the mailbox Display Names, primary SMTP Addresses, and all the
    Send As, Full Access and Send On Behalf To permissions on a CSV file.

.EXAMPLE
.\Export-MailboxFASAPermissions.ps1 -OutputFile C:\temp\EnvironmentPermissions.csv
    Will run the script and export permissions for all mailboxes, in the file specified on the 
    OutputFile parameter : C:\temp\EnvironmentPermissions.csv

.NOTES
    This script can be use alone to export a permissions map, but the output it is intended to be used 
    with the Import-MailboxFASAPermissions.ps1 script.

    "Sens As" permissions
        . Stored in the form of "DOMAIN\Alias"
        . Is set with Add-ADPermission
        . https://docs.microsoft.com/en-us/powershell/module/exchange/active-directory/Add-ADPermission?view=exchange-ps

    "Full Access" Permissions
        . Stored in the form of "DOMAIN\Alias" as well
        . Is set with Add-MailboxPermission
        . https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/Add-MailboxPermission?view=exchange-ps

    "Send On Behalf Of" permissions
        . Stored in the form of "Domain.com/OU_Name/Sub_OU/Name"
        . Is set with Set-Mailbox
        . https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/Set-Mailbox?view=exchange-ps
        . -GrantSendOnBehalfTo parameter accepts one or more values from the below :
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

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")][switch]$SharedMailboxes,
    [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][switch]$ResourceMailboxes,
    [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "NormalRun")][string]$OutputFile,
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
$ScriptVersion = "0.1 Alpha"
<# Version changes
v0.1 - first script version
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME :$ScriptName `nSCRIPT VERSION :$ScriptVersion";exit}
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition #<-- that's for Powershell 2.0
$OutputReport = "$($ScriptPath)\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$($ScriptPath)\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
[array]$report = @()
$Databases = $null
$DBProgressCount = $null
$Mailboxes = @()
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
    param([parameter(Mandatory = $false, Position = 1)] $PercentComplete = 100)
    Write-Progress -id 1 -activity "Working..." -status "In progress..." -PercentComplete ($PercentComplete)
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
<# /FUNCTIONS #>
#endregion Functions
<# -------------------------- EXECUTIONS -------------------------- #>
Test-ExchTools

If (IsEmpty $OutputFile) {$OutputFile = $OutputReport}

$Databases = Get-MailboxDatabase
#Databases = "DB01"

$DBProgressCount = 0
Foreach ($Database in $Databases){
    $DBProgressCount++
    _Progress ($DBProgressCount/$($Databases.count)*100)
    
    $Mailboxescommand = "Get-Mailbox -resultsize unlimited -database $Database"
    If ($ResourceMailboxes -or $SharedMailboxes) {
        $MailboxesCommand += " -RecipientTypeDetails "
        $combo = @()
        If ($ResourceMailboxes){$Combo += "ResourceMailbox"}
        If ($SharedMailboxes){$Combo += "SharedMailbox"}
        $combo = $Combo -join ","
        $MailboxesCommand += $combo
        $MailboxesCommand
        exit
    }
    
    
    If ($ResourceMailboxes -and !$SharedMailboxes){
        $Mailboxes = Get-Mailbox -resultsize unlimited -database $Database -RecipientTypeDetails RoomMailbox,EquipmentMailbox
    } ElseIf ($SharedMailboxes){
        $Mailboxes = Get-Mailbox -resultsize unlimited -database $Database -RecipientTypeDetails SharedMailbox
    } Else {
        $Mailboxes = Get-Mailbox -resultsize unlimited -database $Database
    }
    #$strMailboxes = "Discovery Search Mailbox","RoomTest1 Ottawa" #,"User10", "User1","Test Canada",  "Room 1 - 85 Sparks"
    #$Mailboxes = @()
    #$Mailboxes += $strMailboxes | Get-Mailbox

    Foreach ($Mailbox in $Mailboxes) {
        Write-Host "Working on mailbox $($Mailbox.DisplayName) which Primary SMTP is $($Mailbox.primarySMTPAddress.ToString())" -ForegroundColor Blue -BackgroundColor Yellow
        $SendAs=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self")}
        $FullAccess=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false")}
        $SendOnBehalf = Get-Mailbox $mailbox | Select Alias, @{Name='GrantSendOnBehalfTo';Expression={[string]::join(";", ($_.GrantSendOnBehalfTo))}}
        #Initializing a new Powershell object to store our discovered properties
        $Obj = New-Object PSObject
        #Populating basic mailbox info to bind with SendAs / FullMailbox / SendOnBehalf permissions
        $Obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $Mailbox.DisplayName
        $obj | Add-Member -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $Mailbox.PrimarySMTPAddress.ToString()
		
        If (IsEmpty $SendAs){
            Write-Host "No custom Send As permissions detected"
            $Obj | Add-Member -MemberType NoteProperty -Name "SendAsPermissions" -Value ""
        } Else {
            Write-Host "Found one or more SendAs Permission ! Dumping ..." -ForegroundColor Blue -BackgroundColor green
            [array]$UsersWithSendAs = @()
            ForEach($SAright in $SendAs){$UsersWithSendAs += ($SARight.User.ToString())}
            $strUsersWithSendAs = $UsersWithSendAs -join ";"
            $Obj | Add-Member -MemberType NoteProperty -Name "SendAsPermissions" -Value $strUsersWithSendAs
        }

        If (IsEmpty $FullAccess){
            Write-Host "No custom Full Access permissions detected"
            $Obj | Add-Member -MemberType NoteProperty -Name "FullAccessPermissions" -Value ""
        }  else {
            Write-Host "Found one or more Full Access Permission ! Dumping ..." -ForegroundColor Blue -BackgroundColor green
            [array]$UsersWithFullAccess = @()
            ForEach ($FARight in $FullAccess) {$UsersWithFullAccess += ($FARight.User.ToString())}
            $strUsersWithFullAccess = $UsersWithFullAccess -join ";"
            $Obj | Add-Member -MemberType NoteProperty -Name "FullAccessPermissions" -Value $strUsersWithFullAccess
        }
        
        If (IsEmpty ($SendOnBehalf.GrantSendOnBehalfTo)){
            Write-Host "No custom SendOnBehalf permissions detected"
            $Obj | Add-Member -MemberType NoteProperty -Name "SendOnBehalfPermissions" -Value ""
        } else {
            Write-Host "Found one or more SendOnBehalf Permission ! Dumping ..." -ForegroundColor Blue -BackgroundColor green
            $TableOfSendOnBehalfToConvert = $($SendOnBehalf.GrantSendOnBehalfTo) -Split (";")
            $SMTPAddressesOfSendOnBehalf = @()
            Foreach ($entry in $TableOfSendOnBehalfToConvert) {
                #Since the GrantSendOnBehalfTo entries HAVE to be mailbox-enabled users or mail enabled user or groups,
                #Getting primary SMTP address for each object, and storing these as a string separated by semicolon
                #to replace the string of DOMAIN/OU1/OU2/Name separated by semicolon
                $SMTPAddressesOfSendOnBehalf += (Get-Mailbox $Entry).primarySMTPAddress
            }
            $SendOnBehalfConverted = $SMTPAddressesOfSendOnBehalf -join ";"
            $Obj | Add-Member -MemberType NoteProperty -Name "SendOnBehalfPermissions" -Value $SendOnBehalfConverted
        }
        #Appending the current object into the $report variable (it's an array, remember)
        $report += $Obj

        #Cleaning the variables now before the next loop...
        $SendOnBehalfConverted = $null
        $obj = $Null
        $SMTPAddressesOfSendOnBehalf = $null
        $TableOfSendOnBehalfToConvert = $null
        $SendOnBehalfConverted = $null
        $SendAs = $null
        $FullAccess = $null
        $SendOnBehalf = $null
        #... add more later
    }
}

# Get mailbox forward to from mailboxes:Change the items below that are in bold to fit your needs.
# Get-Mailbox -Filter {ForwardingAddress -ne $Null} |Select Alias, ForwardingAddress | Export-Csv -NoType -encoding "unicode" C:\*location*\MailboxesForwardTo.csv
# Get mailbox grant send on behalf to:Change the items below that are in bold to fit your needs.
#Get-Mailbox -Filter {GrantSendOnBehalfTo -ne $Null} |Select Alias, @{Name='GrantSendOnBehalfTo';Expression={[string]::join(";", ($_.GrantSendOnBehalfTo))}} | Export-Csv -NoType -encoding "unicode" C:\*location*\MailboxesSendOnBehalf.csv
Write-host "saving file in $OutputFile"
$Report | export-csv -NoTypeInformation $OutputFile
Notepad $OutputFile

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>
$Report = $null
$OutputReport = $null
$obj = $null
$strUsersWithSendAs = $null
$strUsersWithFullAccess = $null
$UsersWithSendAs = $null
$UsersWithFullAccess = $null
$SendOnBehalf = $null
$FullAccess = $null
$SendAs = $null
$Mailboxes = $null
$Databases = $Null
<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
