
#region Functions definition #
function AllUrls ($Server,$Internal,[switch]$Test) {
    Set-WebServicesVirtualDirectory -Identity "$server\EWS (Default Web Site)" -internalurl https://$Internal/ews/exchange.asmx -externalurl $null
    set-OWAVirtualDirectory -Identity "$server\owa (Default Web Site)" -InternalURL https://$Internal/owa -externalurl $null
    Set-ActiveSyncVirtualDirectory -Identity "$server\Microsoft-Server-ActiveSync (Default Web Site)" -InternalURL https://$Internal/Microsoft-Server-ActiveSync -externalurl $null
    Set-OabVirtualDirectory -Identity "$server\OAB (Default Web Site)" -InternalURL https://$Internal/oab -externalurl $null
}

function URI ($Server, $Autodiscover,[switch]$Test){
    if ($AutoDiscover -ne "null") {
        $command = "set-clientaccessserver $server -AutoDiscoverServiceInternalUri https://$AutoDiscover/autodiscover/autodiscover.xml"
        if ($test){
            Write-Host"Would execute the following command to set Autodiscover:"
            Write-Hsot $command
        } Else {
            Write-Host "Executing the Autodiscover setting command..."
            Invoke-Expression $command
            Write-host "Done!"
        }

    } else {
        $Command = set-clientaccessserver $server -AutoDiscoverServiceInternalUri $null
        if ($test){
            Write-Host "Would have executed :"
            Write-host $command
        } ELse {
            Write-Host "Executing ..."
            Invoke-Expression $command
        }
    }
}

# Enf of Functions definition
#endregion

$NCRServers = NJES1S5101, NJES1S5102, NJES1S5104, NJES1S5151, NJES1S6252, NJES1S6252-NEW
$NATIONALServers = NJES1S1105, NJES1S1103, NJES1S1109, NJES1S1111, NJES1S1112, NJES1S1503, NJES1S1105 

$NCRInternalURL = "webmail.ci.gc.ca"
$NATIONALIntenralURL = "ActiveSync1.ci.gc.ca"

Foreach ($Server in $NCRServers) {
    AllUrls -Server $Server -Internal $NCRInternalURL
    URI -Server $Server -Autodiscover $NCRInternalURL
}

Foreach ($Server in $NATIONALServers) {
    AllURLs -Server $Server -Internal $NATIONALIntenralURL
    URI -Server $Server -Autodiscover $NATIONALIntenralURL
}

