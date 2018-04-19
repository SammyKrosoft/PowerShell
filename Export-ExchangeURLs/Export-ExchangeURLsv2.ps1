<#
.SYNOPSIS
    This script dumps the URLs of all your Exchange servers in a CSV file

.DESCRIPTION
	This script exports the URLs of all the Exchange servers in a CSV file by default,
	or dumps the info to the screen (you can then redirect to a file if you wish) if you
	use the -DoNotExport switch with the script.

.PARAMETER DoNotExport
	This parameter tells the script not to export to a file. The results will just be
	dumped to the screen.

.PARAMETER CheckVersion
	This parameter dumps the current script version - the script stops processing after displaying the
	version if this parameter is specified, no matter what other parameter is also specified.

.INPUTS
    None.

.OUTPUTS
    Exports a CSV file, or the screen.

.EXAMPLE
.\Export-ExchangeURLsv2.ps1
Will export all URLs for all Exchange servers' virtual directories into a file named ExchangeURLs_Day_Date_Time.csv

.EXAMPLE
.\Export-ExchangeURLsv2.ps1 -DoNotExport
Will just print the results into the PowerShell console.

.NOTES
None

.LINK
    https://github.com/SammyKrosoft
#>
[CmdletBinding(DefaultParameterSetName="NormalRun")]
Param(
	[Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [switch]$DoNotExport,
	[Parameter(Mandatory = $False, Position = 2, ParameterSetName = "checkversion")] [Switch] $CheckVersion
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
$ScriptVersion = "2"
<# Version History
v1.0 -> v2
Added export of Outlook Anywhere with External Hostname (E2010, E2013, E2016) and Internal Hostname (not existing in E2010)
Fixed output file name
Added -DoNoExport switch, to not export to a file...
#> 
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
# $LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
# $LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
If ($CheckVersion) {Write-Host "SCript Version v$ScriptVersion";exit}
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
#Loading Exchange 2010 snapins enabling script to be executed on a basic Powershell session
#Note: you must have Exchange Admin tools installed on the machine where you run this.
Add-PSSnapin microsoft.exchange.management.powershell.admin -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Setup -erroraction 'SilentlyContinue'  | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Support -erroraction 'SilentlyContinue'  | OUT-NULL
#For Exchange 2007 and 2013, add the corresponding modules/snapins, or simply execute the script into an Exchange MAnagement Shell :-)

#Getting all Exchange servers in an array
#Note: you can target only one server, or get servers list from a file,
#just change the $Servers = @(Get-ClientAccessServer) line with $Servers = @(Get-content ServersList.txt) for example to get servers from a list...
$Servers = @(Get-ClientAccessServer)

#Initializing counters to setup a progress bar based on the number of servers browsed
# (more useful in an environment where you have dozen of servers - had 45 in mine)
	$Counter=0
    $Total=$Servers.count	
#Initializing the variable where I'll put all the results of my object browsing
    $report = @()
#For each server discovered in the "$Servers = Get-ClientAccessServer" line, 
# grab the Virtal Directories properties and store it in a custom Powershell object, 
# and then add this object in the $report array variable to eventually dump the whole result in a text (CSV) file.
foreach( $Server in $Servers)
{
    #$Computername=$Server.Name   <- not needed for now
	#This is to print the progress bar incrementing on each server (increment is later in the script $Counter++ it is...
    $Pct=($Counter/$Total)*100    
    Write-Progress -Activity "Processing Server $Server" -status "Server $Counter of $Total" -percentcomplete $pct
	#For the current server, get the main vDir settings (including AutodiscoverServiceInternalURI which is important to determine 
	#whether the Autodiscover service will be hit using the Load Balancer (recommended).
	$EAS = Get-ActiveSyncVirtualDirectory -Server $Server -ADPropertiesOnly | Select Name, InternalURL,externalURL
	$OAB = Get-OabVirtualDirectory -Server $Server -ADPropertiesOnly | Select Name,internalURL,externalURL
	$OWA = Get-OwaVirtualDirectory -Server $Server -ADPropertiesOnly | Select Name,InternalURL,externalURL
	$ECP = Get-EcpVirtualDirectory -Server $Server -ADPropertiesOnly | Select Name,InternalURL,externalURL
	$AutoDisc = get-ClientAccessServer $Server | Select name,identity,AutodiscoverServiceInternalUri
	$EWS = Get-WebServicesVirtualDirectory -Server $Server -ADPropertiesOnly | Select NAme,identity,internalURL,externalURL
    $OA = Get-OutlookAnywhere -Server $Server -ADPropertiesOnly | Select Name,InternalHostName, ExternalHostName
    #If you want to dump more things, use the below line as a sample:
	#$ServiceToDump = Get-Whatever -Server $Server | Select Property1, property2, ....   <- don't need the "Select property", you can omit this, it will just get all attributes...

   	#Initializing a new Powershell object to store our discovered properties
    $Obj = New-Object PSObject

	#the below is a template if you need to dump more things into the final report
	#just replace the "ServiceToDump" string with the service you with to dump - don't forget to 
	#Get something above like the $Service = Get-whatever -Server
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-vDirNAme" -Value $ServiceToDump.Name
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-InternalURL" -Value $ServiceToDump.InternalURL
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-ExernalURL" -Value $ServiceToDump.ExternalURL	
		
	$Obj | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Server.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "EAS-vDirNAme" -Value $EAS.Name
    $Obj | Add-Member -MemberType NoteProperty -Name "EAS-InternalURL" -Value $EAS.InternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "EAS-ExternalURL" -Value $EAS.ExternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "OAB-vDirNAme" -Value $OAB.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "OAB-InternalURL" -Value $OAB.InternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "OAB-ExernalURL" -Value $OAB.ExternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "OWA-vDirNAme" -Value $OWA.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "OWA-InternalURL" -Value $OWA.InternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "OWA-ExernalURL" -Value $OWA.ExternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "ECP-vDirNAme" -Value $ECP.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "ECP-InternalURL" -Value $ECP.InternalURL
	$Obj | Add-Member -MemberType NoteProperty -Name "ECP-ExernalURL" -Value $ECP.ExternalURL	
	$Obj | Add-Member -MemberType NoteProperty -Name "AutoDisc-vDirNAme" -Value $AutoDisc.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "AutoDisc-URI" -Value $AutoDisc.AutodiscoverServiceInternalURI
	$Obj | Add-Member -MemberType NoteProperty -Name "EWS-vDirNAme" -Value $EWS.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "EWS-InternalURL" -Value $EWS.InternalURL
    $Obj | Add-Member -MemberType NoteProperty -Name "EWS-ExernalURL" -Value $EWS.ExternalURL
    $Obj | Add-Member -MemberType NoteProperty -Name "OutlookAnywhere-InternalHostName(NoneForE2010)" -Value $OA.InternalHostName
    $Obj | Add-Member -MemberType NoteProperty -Name "OutlookAnywhere-ExternalHostNAme(E2010+)" -Value $OA.ExternalHostName
		
		
		#Appending the current object into the $report variable (it's an array, remember)
        $report += $Obj
		
		#Incrementing the Counter for the progress bar
        $Counter++
    }
	
	
	If (!($DoNotExport)){
		#Building the file name string using date, time, seconds ...
		$DateAppend = Get-Date -Format "ddd-dd-MM-yyyy-\T\i\m\e-HH-mm-ss"
		$CSVFilename=$PSScriptRoot+"\ExchangeURLs_"+$DateAppend+".csv"
		#Exporting the final result into the output file (see just above for the file string building...
		$report | Export-csv -notypeinformation -encoding Unicode $CSVFilename
		Notepad $CSVFilename
	} Else {
		Write-Host "Won't create a file because you specified the -DoNotExport parameter" -ForegroundColor Yellow -BackgroundColor Blue
		Write-Host "Just dumping to the screen this time ..." -ForegroundColor DarkBlue -BackgroundColor red
		$Report
	}

<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
