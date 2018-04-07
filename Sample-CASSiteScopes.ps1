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


.INPUTS
None. You cannot pipe objects to that script.

.OUTPUTS
Output description to be completed...

.EXAMPLE

C:\PS> .\CASSiteScopes.ps1
Sets the AD sites from default CSV file (if exists) to the CAS servers


.EXAMPLE

C:\PS> .\CASSiteScopes.ps1 -CSVFileName "C:\temp\AutreCSVFile.csv"

Display the command that will sets the AD sites from CSV file specified C:\temp\AutreCSVFile.csv (if exists) to the CAS servers

.EXAMPLE

C:\PS> .\CASSiteScopes.ps1 -Execute

Will not only display the command that will set the AD sites from default CSV file (note : no -CSVFileName parameter used, then
will take the default C:\temp\Classeur1.csv file if it exists (otherwise if Classeur1.csv doesn't exist, will output an error message),
but will also Execute the actual command to set the CAS AutodiscoverSiteScope (aka Site Affinity) as per the map defined in the Classeur1.csv
file

.EXAMPLE

C:\PS> .\CASSiteScopes.ps1 -CSVFileName "C:\temp\AutreCSVFile.csv" -Execute

Same as the above example, will execute the command but this time with the file specified with the -CSVFileName...


.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK

https://github.com/SammyKrosoft
#>
Param(
    [String]$MyFile = "C:\temp\Classeur1.csv",
    [switch]$Execute
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
#Get your CAS servers using Get-ClientAccessServer and put these in $MyCASServers
#$MyCASServers = Get-ClientAccessServer | Select Name
#Putting these manually for the demo
$MyCASServers = "CAS1", "CAS2", "CAS3", "CAS4"


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

Foreach ($CAS in $MyCASServers) {
    Write-host "Currently processing $CAS"
    [array]$SiteScope = @()
    Foreach ($Site in $CASWithSitesScopes.$Cas) {If ($Site -ne $Null -AND $Site -ne 0 -AND $Site -ne ""){$SiteScope += $Site}}
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


<# /EXECUTIONS #>

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
