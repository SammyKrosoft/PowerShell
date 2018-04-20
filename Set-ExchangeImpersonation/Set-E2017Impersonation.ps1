<#
.SYNOPSIS
    This script sets the impersonation rights as well as Send-As and Full Mailbox Access
    to a service account, on a list of mailboxes.

.DESCRIPTION
    IMPORTANT: This script is the Impersonation method for Exchange 2007 only 
    See links section or type Get-Help Set-E2007Impersonation.ps1 -Online

    This script sets the impersonation rights as well as Send-As and Full Mailbox Access to a service account
     on Exchange 2007, on a list of mailboxes provided in a file. There are 2 parameters, which have sample default values.

    From the TechNet:
    =================
    Microsoft Exchange Server 2007 provides two Active Directory directory service 
    extended permissions that are used to determine which callers can perform Exchange Impersonation
    calls and which accounts can be impersonated by the caller.

.PARAMETER UserToAddOnPermissions
    This parameter specifies the service account that will need to access the other
    mailboxes - it will be the account that will "impersonate" these mailboxes to
    connect to these.

.PARAMETER UsersToChangeFilePath
    This parameter specifies the file containing the list of mailboxes that we want
    to impersonate with the account specified on the UserToAddOnPermissions parameter

.INPUTS
    Input needed is the file containing the list of mailboxes to impersonate.

.OUTPUTS
    None.

.EXAMPLE
.\Set-ExchangeImpersonation.ps1

This will enable the default "ServiceAccount" user to impersonate the mailboxes listed 
in the default path .\MailboxesToImpersonate.txt.

.EXAMPLE
.\Set-ExchangeImpersonation.ps1 -UserToAddOnPermissions "BES-Service-Account" -UsersToChangeFilePath "C:\temp\MailboxesToImpersonate.txt"

This will enable the "Bes-Service-Account" to impersonate the mailboxes listed in the "C:\temp\MailboxesToImpersonate.txt" file.

.NOTES
None

.LINK
    https://msdn.microsoft.com/en-us/library/bb204095(v=exchg.80).aspx

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "NormalRun")][string]$UserToAddOnPermissions = "ServiceAccount",
    [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "NormalRun")][string]$UsersToChangeFilePath = "$(split-path -parent $MyInvocation.MyCommand.Definition)\MailboxesToImpersonate.txt",
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
# NOTE: This script was designed in Powershell 2.0 and we want to get
# the script path directory so that we can store our files in the Script's directory
#$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$MyLogFile = "$scriptPath\ADandMailboxPermissionsSetting-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
# Using variables defined in the parameters section
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

Function Test-ExchTools()
{
    Try
    {
        Get-command Get-mailbox -ErrorAction Stop
        $ExchInstalledStatus = $true
    }
    Catch [System.SystemException]
    {
        $ExchInstalledStatus = $false
    }
    Return $ExchInstalledStatus
}

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Clear-Host
If (!(Test-ExchTools)) {
    Write-Host "Exchange Management Tools are not installed or you don't have enough permissions to run these.`nExiting..." -BackgroundColor red -ForegroundColor Yellow
    Exit
    # Write-Debug "Loading Exchange cmdlets..."
    # get-PSSnapin -registered *exchange* | add-pssnapin
}

if (Test-Path $UsersToChangeFilePath)
{
    $AllMailboxesToChange = Get-Content $UsersToChangeFilePath
    ForEach ($MailboxToImpersonate in $AllMailboxesToChange)
    {
        Write-Debug "getting mailbox to process"
        Write-Debug "mailbox to process is $MailboxToImpersonate"
        Write-Debug "Now getting Statistics for $MailboxToImpersonate ..."
        $MailboxDisplayName = (Get-MailboxStatistics -Identity $MailboxToImpersonate).DisplayName
        # $MailboxDisplayName
        # $MailboxDisplayName = $MailboxDisplayName.DisplayName
        Try
        {
            Write-Debug "Adding Impersonation right to $MailboxToImpersonate"
            Add-ADPermission -Identity $MailboxDisplayName -User $UserToAddOnPermissions -extendedRights ms-Exch-EPI-May-Impersonate -ErrorAction Stop

            Write-Debug "Adding Send-As right to $MailboxToImpersonate"
            add-AdPermission -Identity $MailboxDisplayName -user $UserToAddOnPermissions -ExtendedRights “send as” -ErrorAction Stop

            Write-Debug "Adding FullAccess right to $MailboxToImpersonate"
            add-MailboxPermission -Identity $MailboxDisplayName -User $UserToAddOnPermissions -AccessRights "FullAccess" -ErrorAction Stop
                        
            Log "#SUCCESS#Mailbox #$MailboxToImpersonate# successfully processed"
                        
        }
        Catch
        {
            Log "#ERROR#Mailbox #$MailboxToImpersonate# Failed to process Add-ADPermission / Add-MailboxPermission"
        }
        Finally
        {
            Write-Debug "Printing the permissions of $MailboxToImpersonate to check it has impersonation and/or Send As rights ..."
            #Get-AdPermission $MailboxDisplayName | Where { $_.ExtendedRights -like 'ms-Exch-EPI-May-Impersonate' -or $_.ExtendedRights -like "*Send-As*"} #| Format-Table identity, User, Deny, IsInherited, ExtendedRights -AutoSize
            Get-AdPermission $MailboxDisplayName | Where { $_.user -like "*$UserToAddOnPermissions*" -and ($_.ExtendedRights -like 'ms-Exch-EPI-May-Impersonate' -or $_.ExtendedRights -like "*Send-As*")} | Select Identity, User, ExtendedRights, Deny
            Write-Debug "Now getting the Full Access permission because Get-ADPermission only dumps the Extended Rights of a user over a mailbox, and Full mailbox access is NOT an Extended Right ..."
            Write-Debug "We use the Get-MailboxPermissions cmdlet to see if a user has Full Access right" 
            #The below line dumps all users who have " FullAccess" on the mailbox...
            #Get-MailboxPermission $MailboxToImpersonate | Where { $_.AccessRights -like '*FullAccess*'} #| ft Identity, user, AccessRights
            #The below line dumps only user we added "FullAccess" to the mailbox...
            Get-MailboxPermission $MailboxDisplayName | Where { $_.AccessRights -like '*FullAccess*' -and $_.user -like "*$UserToAddOnPermissions*"} #| ft Identity, user, AccessRights
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
