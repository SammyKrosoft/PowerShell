#Re-using a function from my good Serkan Varoglu >> http:\\Mshowto.org >> http:\\Get-Mailbox.org
 
function _Progress
{
    param($PercentComplete,$Status)
    Write-Progress -id 1 -activity "Report for Mailboxes" -status $Status -percentComplete ($PercentComplete)
}
