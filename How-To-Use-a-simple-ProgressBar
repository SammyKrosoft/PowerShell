################### STEPS FOR USING A PROGRESS BAR ###################

#NOTE: You can copy paste that whole post to an ISE or a NOTEPAD and save it as a .ps1 script - the explanations are put in comments for that purpose 🙂
 
#STEP 1 - Get your main objects collection you're going to browse and for which you wish to monitor progression - it can be a Get-Item, or an Import-CSV or a manual list for example - $Mylist = "Server1", "Server2", "Server3", "Server4"
$MyListOfObjectsToProcess = "Server1","Server2","Server3","Server4","Server5","Server6"


#STEP 2 - We need to know how much  items we will have to process (to calculate the % done later on)
$TotalItems = $MyListOfObjectsToProcess.count
 

#STEP 3 - We need to initialize a counter to know what % is done
$Counter = 0

 
#STEP 4 - Process each item (Foreach $Item) of your List Of Objects To Process (in $MyListOfObjectsToProcess)
Foreach ($Item in $MyListOfObjectsToProcess) {

#STEP 5 - We increment the counter
    #  NOTE 1 : if we initialize the counter to $Counter = 1, then we can increment the counter at the end of the Foreach block - in this example we initialized the counter to $Counter = 0 then we increment it right away to go to "1" first ...
    #  NOTE 2 :Developpers like to start at "0" so usually you'll see counters increment at the end of ForEach-like loops - I'm just not following the sheep here, just because I'm French :-p
    $Counter++
 
#STEP 6 - Here is the core : Write-Progress - it comes with 3 mandatory properties : ACTIVITY, STATUS and PERCENTCOMPLETE
    #  ACTIVITY is the title of the progress bar
    #  STATUS is to show on screen which element it is currently processing (Processing item #2 of #20000)
    #  PERCENTCOMPLETE is the progress bar itself and it has to be a percentage (hence the $($Counter / $TotalRecipients * 100) as we calculate the percentage on the fly...

    Write-Progress -Activity "Processing Items" -Status "Item $Counter of $TotalItems" -PercentComplete $($Counter / $TotalItems * 100)

  
#YOUR ROUTINE - Here whatever you want to process on each $Item - that doesn't change anything to the Write-Progress, apart from the time it takes for the progress bar to go to the next % done...
#  In the below example it's just sleeping for 1 second, then outputting each $Item we're processing for the Demo purposes ...

    sleep 1
    Write-Host $Item

}

  

<# Fore more information visit:
 
Powershell progress bar basics :
https://technet.microsoft.com/en-us/library/ff730943.aspx
 

A little bit more with Write-Progress:
https://technet.microsoft.com/en-us/library/2008.03.powershell.aspx

#>
