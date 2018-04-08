<#
.SYNOPSIS
For each CAS server, assign the AD site on its scope.

.DESCRIPTION
For each CAS server in your Exchange organization, get the sites list from a CSV file
and for each CAS, get the corresponding column in the CSV file, and use the values of
that column to assign the list of AD sites to that CAS.

You need to pre-populate a CSV file that has as column header the name of each of your 
CAS server, and the value under each one will be the AD sites list you want to assign
to that CAS. 

For a given AD site, you have some CAS servers, and these CAS will all have the same
AD sites list.

We can optimize this script in the future to get the AD site of each CAS server, and 
when looping through each CAS server, we check which site it belong to, and depending
on that site it belongs to, we add a list of client-only AD sites to that CAS.
That would imply Get-ClientAccessServer, get the site (you will have to get-ClientAccessServer
then for each CAS, Get-ExchangeServer, as the "Site" property doesn't come with Get-ClientAccessServer)
there must be a simpler method to get the AD site name of each CAS server => Research needed here...

.PARAMETER MyFile
The path of the CSV file containing each CAS server hostname and which has a list of AD sites as values under headers

.PARAMETER Execute
Switch to actually execute the AutoDiscover Scope setting instead of just showing what we would do

.PARAMETER UseSampleCASServers
Use this switch, without the -EXECUTE parameter to test the script on sample "CAS1, CAS2, CAS3 and CAS4" server names.
NOTE: this will create a sample CSV file with CAS1, CAS2, CAS3, CAS4 as headers, and fake sites - just to demo the script
with sample values.

.PARAMETER SiteScopeConfigBeforeExecuting
This parameter is optionnal and will let you define the name and path of the output file for your AutoDiscoverSiteScope
current configuration. By default, a file with AutoDiscoverSiteScope_Config_Dump_ prefix and date and time stamp will be
created with the AutodiscoverSiteScope before executing the script inside.

.INPUTS
None. You cannot pipe objects to that script.

.OUTPUTS
Output description to be completed...

.EXAMPLE

C:\PS> .\Sample-CASSiteScopes.ps1 -UseSampleCASServers
Will generate a sample CSV file from within the script with CAS1, CAS2, CAS3 and CAS4, and fake AD Site names, just to show you how the script work a little bit.
The Sample-AutoDiscoverSiteScope.csv will be located on the same directory where you are executing the script, and will show you how your CSV file for your real server must be configured.

.EXAMPLE

C:\PS> .\Sample-CASSiteScopes.ps1
This will launch the script against your production servers using the default <Script Directory>\Classeur.csv file containing your CAS servers (as file header) and your AD sites for each CAS server (under each CAS server header)
This will NOT execute the AD Site scope setting because we don't specify the -EXECUTE parameter -> that's to test if we' re good before executing the changes.
If Classeur.CSV doesn't exist, the script will tell you and exit.

.EXAMPLE

C:\PS> .\Sample-CASSiteScopes.ps1 -Execute
Same as the above example, but this time it will execute the commands and assign the sites to the CAS servers, as defined in the default CLASSEUR.CSV.
If Classeur.CSV does not exist, the script will tell you and exit.

.EXAMPLE

C:\PS> .\Sample-CASSiteScopes.ps1 -CSVFileName "C:\temp\AutreCSVFile.csv" -Execute
Will not only display the command that will set the AD sites from default CSV file (note : no -CSVFileName parameter used, then
will take the default C:\temp\Classeur1.csv file if it exists (otherwise if Classeur1.csv doesn't exist, will output an error message),
but will also Execute the actual command to set the CAS AutodiscoverSiteScope (aka Site Affinity) as per the map defined in the file defined
in the -CSVFileName parameter (in this example, C:\temp\AutreCSVFile.csv)

.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK

https://github.com/SammyKrosoft
#>
Param(
    [String]$MyFile = "$($PSScriptroot)Classeur.csv",
    [switch]$UseSampleCASServers,
    [switch]$Execute,
    [string]$SiteScopeConfigBeforeExecuting = "$($PSScriptroot)AutoDiscoverSiteScope_Config_Dump_$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').csv"
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
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
$Answer = ""
#Doing Conditional declarations here ... if specifying -GetRealCASServers, get the real servers
If ($UseSampleCASServers) {
        #Putting these manually for the demo
        $MyCASServers = "CAS1", "CAS2", "CAS3", "CAS4"
        $SampleCSVData = @"        
CAS1,CAS2,CAS3,CAS4
AD_DatecenterSite1,AD_DatecenterSite1,AD_DatacenterSite3,AD_DatacenterSite3
AD_ClientSite1,AD_ClientSite1,AD_ClientSite6,AD_ClientSite6
AD_ClientSite2,AD_ClientSite2,AD_ClientSite7,AD_ClientSite7
AD_ClientSite3,AD_ClientSite3,AD_ClientSite8,AD_ClientSite8
"@
        $MyFile = "$($PSScriptroot)Sample-AutoDiscoverSiteScope.csv"

        $SampleCSVData | out-file $MyFile
        
        
} else {
        #Get your CAS servers using Get-ClientAccessServer and put these in $MyCASServers
        $MyCasServers = @()
        $MyCASServers = Get-ClientAccessServer | Select Name
        $MyCASServers = $MyCASServers | Foreach {$_.Name}
}


<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>

Try 
{
    #Import your CSV where you have your CAS1 - CAS2 - CAS3 - CAS4 servers with list of AD sites in each
    #use the -MyFile parameter with the script to replace the default C:\Temp\Classeur1.csv with your own file
    $CASWithSitesScopes = import-csv $MyFile -ErrorAction Stop
}
Catch
{
    cls
    Write-Host "Failed to load $MyFile, exiting..." -BackgroundColor yellow -ForegroundColor blue
    Break
}

$Answer = ""
while ($Answer -ne "Y" -AND $Answer -ne "N") {
    cls
    If (!$UseSampleCASServers) {
        Write-Host "Here is the current AutodiscoverSiteScope configuration:" -BackgroundColor Yellow -ForegroundColor Red
        $ConfigBeforeChange = Get-ClientAccessServer | select Name, AutodiscoverSiteScope
        $ConfigBeforeChange | ft -a 
        $ConfigBeforeChange | ft -a > $SiteScopeConfigBeforeExecuting # For now exports as printed on screen, not pure CSV format ... To be fixed later
        # $ConfigBeforeChange | Export-CSV -NoTypeInformation $SiteScopeConfigBeforeExecuting <-- exports with "Microsoft.Exchange.Data.MultiValuedProperty`1[System.String]"
        Write-Host "Configuration dumped into $SiteScopeConfigBeforeExecuting file for later reference and in case of misconfiguration" -BackgroundColor Yellow -ForegroundColor Red
    }
    Write-Host "Validating the Script options"  -BackgroundColor Red -ForegroundColor Blue
    Write-Host "The file we'll use to populate AutodiscoverSiteScope is   : $MyFile" -BackgroundColor Blue -ForegroundColor Yellow
    Write-Host "Use Sample CAS Servers                                    : $UseSampleCASServers"  -BackgroundColor Blue -ForegroundColor Yellow
    Write-Host "Execute the AutoDiscoverSiteScope                         : $Execute"  -BackgroundColor Blue -ForegroundColor Yellow
    Write-Host "`nContinue (Y/N) ?" -BackgroundColor Red -ForegroundColor Blue
    $Answer = Read-host
    If ($Answer -eq "Y"){$StopWatch.Reset();$StopWatch.start()} else {exit}
}

Foreach ($CAS in $MyCASServers) {
    Write-host "Currently processing $CAS"
    $SiteScope = @()
    $CASWithSitesScopes | Foreach {If (($($_.$CAS) -ne $Null) -AND ($($_.$CAS) -ne 0) -AND ($($_.$CAS) -ne "")){$SiteScope += $_.$CAS}}
    $Command = "Set-ClientAccessServer $CAS -AutodiscoverSiteScope `$SiteScope"
    Write-Host $Command -BackgroundColor yellow -ForegroundColor Red
    Write-Host "Where SiteScope has $($SiteScope.count) entries :"
    $SiteScope | Foreach {Write-host $_ -BackgroundColor green -ForegroundColor Yellow}
    If (!$Execute) {
        Write-Host "Testing only ... not executing actual command."
    } Else {
        Invoke-Expression $Command
    }
}

If ($Execute) {
    Write-Host "Here is the !!NEW!! AutodiscoverSiteScope configuration:" -BackgroundColor Yellow -ForegroundColor red
    $ConfigAfterChange = Get-ClientAccessServer | select Name,AutodiscoverSiteScope
    $ConfigAfterChange | ft -a
}

<# /EXECUTIONS #>

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
