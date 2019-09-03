<#
.SYNOPSIS
How to use WPF MessageBox

.DESCRIPTION
Message Box WPF Class:
https://msdn.microsoft.com/en-us/library/ms602949.aspx



.INPUTS
None. You cannot pipe objects to that script.

.OUTPUTS
System.String. The script (Full-Name.ps1 or whatever you name it) returns a string with the full
name.

.EXAMPLE
.\Full-Name.ps1
Your full name is Merlin the Wizard

.LINK
http://aka.ms/sammy
http://github.com/sammykrosoft

#>

# Always load WPF assembly to be able to use "[System.Windows.MessageBox]"
Add-Type -AssemblyName presentationframework, presentationcore

# Option #1 - only a message
$msg = "Test"
[system.windows.MessageBox]::show($msg)

# Option #2 - a message and a title
$msg = "Test"
$Title = "Title"
[System.Windows.MessageBox]::Show($msg,$Title)

# Option #3 - a message, a title and a button
# More info : https://msdn.microsoft.com/en-us/library/ms598690.aspx
$msg = "Test"
$Title = "Title"
$Button = "YesNoCAncel"
[System.Windows.MessageBox]::Show($msg,$Title, $Button)

# Option #4 - a message, a title, buttons, and an icon
# More info : https://msdn.microsoft.com/en-us/library/system.windows.messageboximage.aspx
$msg = "Do you want to continue ?"
$Title = "Question..."
$Button = "YesNo"
$Icon = "Question"
[System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)

# Option #5 - a message, a title, buttons, an icon, and a default result
$msg = "Do you want to continue ?"
$Title = "Question..."
$Button = "YesNo"
$Icon = "Question"
$DefaultResult = "Ok"
[System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon, $defaultResult)

# Option #6 - using a function to call the box each time we need it without retyping the same code
Function MsgBox {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0)][String]$msg = "Do you want to continue ?",
        [Parameter(Position=1)][String]$Title = "Question...",
        [Parameter(Position=2)]
            [ValidateSet("OK","OKCancel","YesNo","YesNoCancel")]
                [String]$Button = "YesNo",
        [Parameter(Position=3)]
            [ValidateSet("Asterisk","Error","Exclamation","Hand","Information","None","Question","Stop","Warning")]
                [String]$Icon = "Question"
    )
    Add-Type -AssemblyName presentationframework, presentationcore
    [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)

}