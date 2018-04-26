<#
.SYNOPSIS
    ************ SCRIPT IN PROGRESS NOT TESTED ****************
    Export Exchange Mailbox Full Access permissions in a CSV file in order to import them in another 
    environment using the output CSV file.

.DESCRIPTION
    Longer description of what this script does
    SCRIPT IN PROGRESS NOT TESTED

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
SCRIPT IN PROGRESS NOT TESTED

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][int]$FirstNumber = 1,
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][int]$SecondNumber = 2,
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
v0.1 - first script version
#>
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$OutputReport = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$PSScriptRoot\$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
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

    Write-Progress -id 1 -activity $Activity -status $Status -PercentComplete ($PercentComplete)
    }

    <# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
$Databases = Get-MailboxDatabase
$DBProgressCount = 0
Foreach ($Database in $Databases){
    $DBProgressCount++
    _Progress ($DBProgressCount/$Databases.count*100) "Processing mailboxes database by database" "Current database : $($Database.name)"
    Get-Mailbox -resultsize unlimited -database $Database
}

# Get mailbox forward to from mailboxes:Change the items below that are in bold to fit your needs.
# Get-Mailbox -Filter {ForwardingAddress -ne $Null} |Select Alias, ForwardingAddress | Export-Csv -NoType -encoding "unicode" C:\*location*\MailboxesForwardTo.csv

# Get mailbox grant send on behalf to:Change the items below that are in bold to fit your needs.
#Get-Mailbox -Filter {GrantSendOnBehalfTo -ne $Null} |Select Alias, @{Name='GrantSendOnBehalfTo';Expression={[string]::join(";", ($_.GrantSendOnBehalfTo))}} | Export-Csv -NoType -encoding "unicode" C:\*location*\MailboxesSendOnBehalf.csv

# From other
$SendAs=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self")} | Sort-Object name
$FullAccess=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false")}
$SendOnBehalf = Get-Mailbox $mailbox -Filter {GrantSendOnBehalfTo -ne $Null} |Select Alias, @{Name='GrantSendOnBehalfTo';Expression={[string]::join(";", ($_.GrantSendOnBehalfTo))}}



<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
