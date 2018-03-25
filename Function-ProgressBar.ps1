
function _Progress
{
    param($PercentComplete=100,$Status="In Progress...")

<# 
.SYNOPSIS
Simple progress bar function using Write-Progress but simplify a bit the call
   
.DESCRIPTION
Re-using a function from my good Serkan Varoglu >> http:\\Mshowto.org >> http:\\Get-Mailbox.org
 
.NOTES
To report % progress, use a $Counter variable that you increment at each loop iteration, divide by the Total number 
of items of the collection you're looping in, multiplied by 100:
$PercentComplete = $Counter/$TotalItems*100

.LINK
http:\\Mshowto.org

.LINK
http:\\Get-Mailbox.org

#>

Write-Progress -id 1 -activity "Working !" -status $Status -percentComplete ($PercentComplete)
}
