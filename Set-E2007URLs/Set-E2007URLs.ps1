
#region Functions definition #

function Global:Convert-HString {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] [String]$HString
        )

    <#NOTE: This function is from Ben Wilkinson - https://gallery.technet.microsoft.com/scriptcenter/917c2357-2911-4c79-bd06-ab95714de2d4#>

    Begin 
    {Write-Verbose "Converting Here-String to Array"}
    Process 
    {
        $HString -split "`n" | ForEach-Object {
            $Item = $_.trim()
            #NOTE: below is to enable the use of hashtag to comment aka ignore #lines in your txt file...
            if ($Item -notmatch "#")
            {$Item}    
        }
    }#Process
    End 
    {
        # Nothing to do here.
    }
}#Convert-HString


function AllUrls ($Server,$Internal,[switch]$Test) {
    $Command1 = "Set-WebServicesVirtualDirectory -Identity ""$server\EWS (Default Web Site)"" -InternalURL https://$Internal/ews/exchange.asmx -ExternalURL `$null"
    $Command2 = "set-OWAVirtualDirectory -Identity ""$server\owa (Default Web Site)"" -InternalURL https://$Internal/owa -ExternalURL `$null"
    $Command3 = "Set-ActiveSyncVirtualDirectory -Identity ""$server\Microsoft-Server-ActiveSync (Default Web Site)"" -InternalURL https://$Internal/Microsoft-Server-ActiveSync -ExternalURL `$null"
    $Command4 = "Set-OabVirtualDirectory -Identity ""$server\OAB (Default Web Site)"" -InternalURL https://$Internal/oab -ExternalURL `$null"

    if ($Test) {
        Write-Host "#Would execute the following commands :" -BackgroundColor Blue -ForegroundColor Yellow
        Write-Host $Command1
        Write-Host $Command2
        Write-Host $Command3
        Write-Host $Command4
    } Else {
        Invoke-Command $Command1
        Invoke-Command $Command2
        Invoke-Command $Command3
        Invoke-Command $Command4
    }
}

function URI ($Server, $Autodiscover,[switch]$Test){
    if ($AutoDiscover -ne "null") {
        $command = "set-clientaccessserver $server -AutoDiscoverServiceInternalUri https://$AutoDiscover/autodiscover/autodiscover.xml"
        if ($test){
            Write-Host "#Would execute the following command :"  -BackgroundColor Blue -ForegroundColor Yellow
            Write-Host $command
        } Else {
            Write-Host "Executing the Autodiscover setting command..."  -BackgroundColor Blue -ForegroundColor Yellow
            Invoke-Expression $command
            Write-host "Done!"
        }

    } else {
        $Command = set-clientaccessserver $server -AutoDiscoverServiceInternalUri $null
        if ($test){
            Write-Host "#Would have executed :"
            Write-host $command
        } ELse {
            Write-Host "Executing ..."
            Invoke-Expression $command
        }
    }
}

# Enf of Functions definition
#endregion

$NCRServers = @"
NJES1S5101
NJES1S5102
NJES1S5104
NJES1S5151
NJES1S6252
NJES1S6252-NEW
"@

$NATIONALServers = @"
NJES1S1103
NJES1S1109
NJES1S1111
NJES1S1112
NJES1S1503
"@

$NATIONALServers = Convert-HString $NATIONALServers
$NCRServers = Convert-HString $NCRServers


$NCRInternalURL = "webmail.ci.gc.ca"
$NAtionalInternalURL = "ActiveSync1.ci.gc.ca"
$AutodiscoverURIForNCRSite = "Autodiscover.ci.gc.ca"
$AutoDiscoverURIForNATIONALSite = $AutodiscoverURIForNCRSite

Write-Host "#SERVERS IN NCR : $($NCRServers -join ",")" -BackgroundColor Yellow -ForegroundColor Red
Foreach ($Server in $NCRServers) {
    AllUrls -Server $Server -Internal $NCRInternalURL -Test
#    URI -Server $Server -Autodiscover $AutodiscoverURIForNCRSite -Test
}

Write-Host "#SERVERS IN NATIONAL : $($NATIONALServers -join ",")" -BackgroundColor Yellow -ForegroundColor Red

Foreach ($Server in $NATIONALServers) {
    AllURLs -Server $Server -Internal $NAtionalInternalURL -Test
#    URI -Server $Server -Autodiscover $AutoDiscoverURIForNATIONALSite -Test
}

