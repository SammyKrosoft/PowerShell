<#
.SYNOPSIS

Prints your first name and last name.
Get this help from header by typing Get-Help .\YourScript.ps1 -Full

.DESCRIPTION

Just a dummy script that prints your first name and last name.
Takes any strings for first name and last name.

.PARAMETER FirstName
Specifies the First Name. "Merlin" is the default.

.PARAMETER LastName
Specifies the last name. "the Wizard" is the default.

.INPUTS

None. You cannot pipe objects to that script.

.OUTPUTS

System.String. The script (Full-Name.ps1 or whatever you name it) returns a string with the full
name.

.EXAMPLE

C:\PS> .\Full-Name.ps1 -FirstName "John" -LastName "Doe"
Your full name is John Doe

.EXAMPLE

C:\PS> .\Full-Name.ps1 -FirstName "Jane" -LastName "Doe"
Your full name is Jane Doe

.EXAMPLE

C:\PS> .\Full-Name.ps1 "Jane" "Doe"
Your full name is Jane Doe

.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK

Set-Item
#>

param ([string]$FirstName="Merlin",[string]$LastName = "the Wizard")
$name = $FirstName + " " + $LastName
Write-host "Your full name is $name"
