<#
.SYNOPSIS

Adds a file name extension to a supplied name.
Get this help from header by typing Get-Help .\YourScript.ps1 -Full

.DESCRIPTION

Adds a file name extension to a supplied name.
Takes any strings for the file name or extension.

.PARAMETER Name
Specifies the file name.

.PARAMETER Extension
Specifies the extension. "Txt" is the default.

.INPUTS

None. You cannot pipe objects to Add-Extension.

.OUTPUTS

System.String. Add-Extension returns a string with the extension
or file name.

.EXAMPLE

C:\PS> extension -name "File"
File.txt

.EXAMPLE

C:\PS> extension -name "File" -extension "doc"
File.doc

.EXAMPLE

C:\PS> extension "File" "doc"
File.doc

.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK

Set-Item
#>

param ([string]$Name,[string]$Extension = "txt")
$name = $name + "." + $extension
$name
