<#
.SYNOPSIS

Prints your first name and last name.
Get this help from header by typing Get-Help .\Your-Script.ps1 -Full

.DESCRIPTION

The Your-Script.ps1 script prints your first name and last name
from parameters when calling the script. It's a demo-purposed script
that's why you'll see parameters taken from the script, then a function
that concatenates 2 strings and prints "Your full name is John Doe", which
I call using the parameters value taken from the script. It's definitely
overkill but that's to demo Comments defined at the script level and then
using function to use the parameters defined at the script level ...

.PARAMETER FirstName
Specifies the First Name. "Merlin" is the default.

.PARAMETER LastName
Specifies the First Name. "the Wizard" is the default.

.INPUTS
None. You cannot pipe objects to Your-Script.ps1.

.OUTPUTS
None. Your-Script.ps1 does not generate any output.

.EXAMPLE

C:\PS> .\Your-Script.ps1

.EXAMPLE

C:\PS> .\Your-Script.ps1 -FirstName "Jane" -LastName "Doe"
Your full name is Jane Doe

.EXAMPLE

C:\PS> .\Your-Script.ps1 "Jane" "Doe"
Your full name is Jane Doe
#>

param ([string]$FirstName="Merlin", [string]$LastName="the Wizard")

function Print-FullName ([string]$FName, [string]$LName) { 
  $FullName = $FName + " " + $LName
  Write-Host "Your Full Name is $FullName"
}

Print-FullName $FirstNAme $LastName
