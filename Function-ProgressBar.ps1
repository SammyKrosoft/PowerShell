<# 
.SYNOPSIS
Simple progress bar function using Write-Progress but simplify a bit the call
   
.DESCRIPTION
Re-using a function from my good Serkan Varoglu
 
.NOTES
To report % progress, use a $Counter variable that you increment at each loop iteration, divide by the Total number 
of items of the collection you're looping in, multiplied by 100:
$PercentComplete = $Counter/$TotalItems*100

.LINK
http:\\Mshowto.org

.LINK
http:\\Get-Mailbox.org

#>


function _Progress {
    param(
        [parameter(position = 0)] $PercentComplete=100,
        [parameter(position = 1)] $Activity = "Working...",
        [parameter(position = 2)] $Status="In Progress..."
        )

    Write-Progress -id 1 -activity $Activity -status $Status -PercentComplete ($PercentComplete)
    }