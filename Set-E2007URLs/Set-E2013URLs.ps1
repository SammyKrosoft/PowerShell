
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
    $Command2bis = "set-ECPVirtualDirectory -Identity ""$server\ecp (Default Web Site)"" -InternalURL https://$Internal/ecp -ExternalURL `$null"
    $Command3 = "Set-ActiveSyncVirtualDirectory -Identity ""$server\Microsoft-Server-ActiveSync (Default Web Site)"" -InternalURL https://$Internal/Microsoft-Server-ActiveSync -ExternalURL `$null"
    $Command4 = "Set-OabVirtualDirectory -Identity ""$server\OAB (Default Web Site)"" -InternalURL https://$Internal/oab -ExternalURL `$null"
    $Command5 = "Get-OutlookAnywhere | Set-OutlookAnywhere -InternalHostname $Internal -InternalClientsRequireSsl `$true"

    if ($Test) {
        Write-Host "#Would execute the following commands :" -BackgroundColor Blue -ForegroundColor Yellow
        Write-Host $Command1
        Write-Host $Command2
        Write-Host $Command2bis
        Write-Host $Command3
        Write-Host $Command4
        Write-Host $Command5
    } Else {
        Invoke-Expression $Command1
        Invoke-Expression$Command2
        Invoke-Expression $Command2bis
        Invoke-Expression $Command3
        Invoke-Expression $Command4
        Invoke-Expression $Command5
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

$NCRE2013Servers = @"
NJES1S6503
NJES1S6504
NJES1S6505
NJES1S6506
"@

$NCRE2013Servers = Convert-HString $NCRE2013Servers


$NCRInternalURL = "webmail.ci.gc.ca"
$AutodiscoverURIForNCRSite = "Autodiscover.ci.gc.ca"

Write-Host "#SERVERS IN NCR : $($NCRE2013Servers -join ",")" -BackgroundColor Yellow -ForegroundColor Red
Foreach ($Server in $NCRE2013Servers) {
#    AllUrls -Server $Server -Internal $NCRInternalURL -Test
    URI -Server $Server -Autodiscover $AutodiscoverURIForNCRSite -Test
}