<#
.SYNOPSIS
    Quick description of this script

.DESCRIPTION
    Longer description of what this script does

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
$ScriptVersion = "1.0"
<# Version changes
-> v1.0
#>
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
# $LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
# $LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
# NOTE: This script was designed in Powershell 2.0 and we want to get
# the script path directory so that we can store our files in the Script's directory
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$MyLogFile = "$scriptPath\ADandMailboxPermissionsSetting-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
$UserToAddOnPermissions = "ServiceAccount"
# Below is the file that contains the list of mailboxes that we want to modify
$UsersToChangeFilePath = "$scriptPath\UsersToChange.txt"
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function Log
{

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')] 
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipeLineByPropertyName = $true,
            Position = 0)]
        [String[]]$String, 
        [Parameter(Mandatory = $false,
            Position = 1)]
        [ValidateSet("Black", "Blue", "Cyan", "DarkBlue", "DarkCyan", "DarkGray", "DarkGreen", "DarkMagenta", "DarkRed", "DarkYellow", "Gray", "Green", "Magenta", "Red", "White", "Yellow")]
        [String[]]$Color = "Green", 
        [Parameter(Mandatory = $false,
            Position = 2)]
        [String]$LogFile = $MyLogFile,
        [Parameter(Mandatory = $false,
            Position = 3)]
        [Switch]$NoNewLine
    )


    $LegalFileNameCharSet = "^[" + [Regex]::Escape("A-Za-z0-9^&'@{}[],$=!-#()%.+~_") + "]+$"
    if ($String.Count -gt 1)
    {
        $i = 0
        foreach ($item in $String)
        {
            if ($Color[$i]) { $col = $Color[$i] } else { $col = "White" }
            Write-Host "$item " -ForegroundColor $col -NoNewline
            $i++
        }
        if (-not ($NoNewLine)) { Write-Host " " }
    }
    else
    { 
        if ($NoNewLine) { Write-Host $String -ForegroundColor $Color[0] -NoNewline }
        else { Write-Host $String -ForegroundColor $Color[0] }
    }

    if ($LogFile.Length -gt 2 -and !($LogFile -match $LegalFileNameCharSet))
    {
        "$(Get-Date -format 'dd MMMM yyyy hh:mm:ss tt'): $($String -join " ")" | Out-File -Filepath $Logfile -Append 
    }
    else
    {
        Write-Debug "Log: Missing -LogFile parameter or bad LogFile name. Will not save input string to log file.."
    }
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
#Load cmdlet
Write-Debug "Loading Exchange cmdlets..."
get-PSSnapin -registered *exchange* | add-pssnapin

if (Test-Path $UsersToChangeFilePath)
{
    $AllMailboxesToChange = Get-Content $UsersToChangeFilePath
    ForEach ($MailboxToChange in $AllMailboxesToChange)
    {
        Write-Debug "getting mailbox to process"
        Write-Debug "mailbox to process is $MailboxToChange"
        Write-Debug "Now getting Statistics for $MailboxToChange ..."
        $MailboxToChangeGet = Get-MailboxStatistics -Identity $MailboxToChange
        $MailboxToChangeGet
        $MailboxToChangeGet = $MailboxToChangeGet.DisplayName
        Try
        {
            Write-Debug "Adding Impersonation right to $MailboxToChange"
            Add-ADPermission -Identity $MailboxToChangeGet -User $UserToAddOnPermissions -extendedRights ms-Exch-EPI-May-Impersonate -ErrorAction Stop

            Write-Debug "Adding Send-As right to $MailboxToChange"
            add-AdPermission -Identity $MailboxToChangeGet -user $UserToAddOnPermissions -ExtendedRights “send as” -ErrorAction Stop

            Write-Debug "Adding FullAccess right to $MailboxToChange"
            add-MailboxPermission -Identity $MailboxToChangeGet -User $UserToAddOnPermissions -AccessRights "FullAccess" -ErrorAction Stop
                        
            Log "#SUCCESS#Mailbox #$MailboxToChange# successfully processed"
                        
        }

        Catch
        {

            Log "#ERROR#Mailbox #$MailboxToChange# Failed to process Add-ADPermission / Add-MailboxPermission"

        }

        Finally
        {

            Write-Debug "Printing the permissions of $MailboxToChange to check it has impersonation and/or Send As rights ..."
            #Get-AdPermission $MailboxToChangeGet | Where { $_.ExtendedRights -like 'ms-Exch-EPI-May-Impersonate' -or $_.ExtendedRights -like "*Send-As*"} #| Format-Table identity, User, Deny, IsInherited, ExtendedRights -AutoSize
            Get-AdPermission $MailboxToChangeGet | Where { $_.user -like "*$UserToAddOnPermissions*" -and ($_.ExtendedRights -like 'ms-Exch-EPI-May-Impersonate' -or $_.ExtendedRights -like "*Send-As*")} | Select Identity, User, ExtendedRights, Deny
            Write-Debug "Now getting the Full Access permission because Get-ADPermission only dumps the Extended Rights of a user over a mailbox, and Full mailbox access is NOT an Extended Right ..."
            Write-Debug "We use the Get-MailboxPermissions cmdlet to see if a user has Full Access right" 
            #The below line dumps all users who have " FullAccess" on the mailbox...
            #Get-MailboxPermission $MailboxToChange | Where { $_.AccessRights -like '*FullAccess*'} #| ft Identity, user, AccessRights
            #The below line dumps only user we added "FullAccess" to the mailbox...
            Get-MailboxPermission $MailboxToChangeGet | Where { $_.AccessRights -like '*FullAccess*' -and $_.user -like "*$UserToAddOnPermissions*"} #| ft Identity, user, AccessRights
        }
    }     
                
}
Else
{
    #If the users files does NOT exist ... just exit ...
    Write-Host "The File $UsersToChangeFilePath does NOT exist, create it first !!" -BackgroundColor Red -ForegroundColor Yellow
    Break
}
<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
