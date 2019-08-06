$TryCount = 0
$Done = $false
do{
    # It takes a while after enabling mailbox until settings can be applied. So we need to retry.
    try{
        # If we need to execute a setting several times.
        if ($MailboxSetting.LoopOver){
            # We have a loop value (array).
            foreach ($LoopValue in $MailboxSetting.LoopOver){
                # Copy parameter as we have to change a value (loop value).
                $TempParams = $Params.PsObject.Copy()                               
                @($Params.getenumerator()) |? {$_.Value -match '#LOOPVALUE#'} |% {$TempParams[$_.Key]=$LoopValue} 
                $res = & $MailboxSetting.Command -ErrorAction Stop @TempParams -WhatIf:$RunConfig.TestMode
            }
        } else {
            $res = & $MailboxSetting.Command -ErrorAction Stop @Params -WhatIf:$RunConfig.TestMode
        }
        # Write-Log "Setting command $($MailboxSetting.Command) executed successfully"
        $Done = $true
    } catch{
        $tryCount++
        $res = Write-Error -err $error -msg "Error applying mailbox settings, account: $($AccountDetails.sAMAccountName), retry count: $($TryCount)" -Break $false
        Start-Sleep -s $(($Retires-$TryCount) * 5)

        try{
            # We may have lost the Kerberos ticket, reconnect to Exchange.
            Disconnect-Exchange
            Connect-Exchange
        } catch {}
    } 
} while ((!$done) -and ($tryCount -lt $Retires))


# O365 renewal : REFRESHTOKEN