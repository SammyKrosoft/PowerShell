<#

.Description
PREREQUISITES :

1- the account with which you launch the below export script MUST be a member of a role group
that has the "Mailbox Import Export"  rights. On PCO, that role group is "Mailbox Exports"
NOTE: If such a Role Group does NOT exist, see Sam's blog post on the LINK section below...

2- The UNC path where you export the PST must start with double slashes, and it seems that it works
only when you export on Exchange servers (Exchange Trusted Subsystem account seems to be needed as local admin
of the machine where you export the PST)


.EXAMPLE
[PS] C:\Users\super-ssc-sm\Documents>.\Export-FolderToPST.ps1

Will export the folder that you hard coded on the variable $FolderToExport
from the mailbox that is specified in the variable $MailboxToExport
and it will put it on the UNC path specified in the variable $UNCFilePathToExportThePST
The Request will be able to be retrieved using Get-MailboxExportRequest with the "-Name" parameter
and the name provided in the $ExportRequestName variable.

.EXAMPLE
[PS] C:\Users\super-ssc-sm\Documents>.\Export-FolderToPST.ps1 -ExportRequestName "MyExportRequest01"

Will create an Export request to dump a mailbox inside a PST file, and the Export request job will be called "MyExportRequest01".
The mailbox it will search in will be the one hard coded in the $MailboxToExport, the UNC Path specified on the $UNCFilePathToExportThePST,
and the folder specified on the $FolderToExport variable.

.LINK
https://blogs.technet.microsoft.com/samdrey/2011/02/16/exchange-2010-rbac-issue-mailbox-import-export-new-mailboximportrequest-and-other-management-role-entries-are-missing-not-available/


#>
param(
    [string]$FolderToExport = "Restore/Restore - My Mailbox/Inbox/*",
    [string]$MailboxToExport = "My Mailbox Name",
    [string]$UNCFilePathToExportThePST = "\\YourExchangeExportServer\C$\temp\Restored-$(get-date -f yyyy-MM-dd-hh-mm-ss).pst",
    [string]$ExportRequestName = "MyExportRequest"
    )

#Removing previous Mailbox Export request that had the same name as the name provided
#Note for the future: we can develop a simple routing that checks for existing $ExportRequestName, and if it exists, exit the script with instruction to specify another name...
Remove-MailboxExportRequest $ExportRequestName
    
#Write-Host "Checking if Exchange can find $MAilboxToExport ..." 
#Note for the future: we can test if the mailbox targetted exists, and if it doesn't, exit the script...
#Get-mailbox $MailboxToExport

Write-Host "Trying to export data from $MailboxToExport and targetting folder $FolderToExport ..."
New-MailboxExportRequest -Name $ExportRequestName -IncludeFolders $FolderToExport -Mailbox $MailboxToExport -Filepath $UNCFilePathToExportThePST

#Getting the status of the newly created Export Request...
Get-MailboxExportRequest -name $ExportRequestName 
