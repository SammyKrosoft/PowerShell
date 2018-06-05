<# 
.SYNOPSIS 
Creates a WPF Message Box with the supplied message and title and returns the response. 
(C) Chris Carter
https://gallery.technet.microsoft.com/scriptcenter/Create-a-New-WPF-Message-1b14ff1b
 
.DESCRIPTION
New-MessageBox creates a new WPF Message Box with the supplied message and title and returns the response.  The desired button configuration, the icon accompanying the message, as well as the default button can also be set.   
 
The default is to capture the button that was clicked by the user, but it can be discarded using the Quiet parameter.  Also, when the button configuration has two buttons, the OutputBool parameter can be used to convert the default response type of Windows.MessageBoxResult to a True or False value.  
 
New-MessageBox will accept objects along the pipeline for the Message and Title Parameters.  Aliases have been used to also accomodate errors stored in the $Error variable. 
 
.PARAMETER Message 
The message to display in the main body of the popup. 
 
.PARAMETER Title 
The Title to display in the title bar of the popup. 
 
.PARAMETER Button 
The buttons to be displayed on the popup. The valid choices are OK, OKCancel, YesNo, YesNoCancel. The default is OK. 
 
.PARAMETER Icon 
The icon to use in the popup. The valid choices are Asterisk, Error, Exclamation, Hand, Information, None, Question, Stop, Warning. The default is Information.  
 
.PARAMETER DefaultButton 
Sets the default button to be activated with the Enter key. The valid choices are OK, Yes, No, Cancel, None. The default is None.  
 
.PARAMETER Quiet 
Using the Quiet parameter will force New-MessageBox to have no ouputs. 
 
.PARAMETER OutputBool 
Setting the OutputBool parameter will cause the result to be converted to a Boolean value. 
Example:  If the popup has Okay and Cancel buttons, using OutputBool would cause the output of True for Okay and False for Cancel. 
 
.INPUTS 
System.String 
You can pipe string objects to New-MessageBox or any object that has an appropriate ToString() method. 
 
System.Management.Automation.ErrorRecord 
You can pipe ErrorRecord objects to New-MessageBox that bind on the FullyQualifiedErrorID and Exception properties. 
 
System.Management.Automation.ParseException 
You can pipe ParseException objects to New-MessageBox that bind on the Source and Message properties. 
 
.OUTPUTS 
System.Windows.Forms.DialogResult 
 
System.Boolean 
The OutputBool parameter will force New-MessageBox to output a boolean value. 
 
None 
The Quiet parameter will force New-MessageBox to have no outputs. 
 
.EXAMPLE 
PS C:\> New-MessageBox -Message "Test" -Title "Test Title" 
 
This command will show a message box with the message of "Test" and a title of "Test Title" with the Information icon and an OK button. 
.EXAMPLE 
PS C:\> New-MessageBox -Message "Test" -Title "Test Title" -Buttons "YesNo" -Icon "Question" 
 
This command will display a message box with the message of "Test" and a title of "Test Title" with the Question mark icon and the Yes and No Buttons. 
.EXAMPLE 
PS C:\> $Error | New-MessageBox -Icon "Error" 
 
This command will pipe the errors in $Error to New-MessageBox and display a message box for each error with an Error icon. 
.NOTES 
It is recommended that if you want to use this in GUI scripts, wrap it in a function declaration and include it in the script. 
Author:  Chris Carter 
Version: 1.0 
 
.LINK 
https://msdn.microsoft.com/en-us/library/ms598711(v=vs.110).aspx 
 
.COMPONENT 
PresentationFramework 
#> 
 
#Requires -Version 2.0 
[CmdletBinding(DefaultParameterSetName="Quiet")] 
 
Param( 
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
        [Alias("Caption","Description","Exception")] 
        [String]$Message, 
 
    [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
        [Alias("FullyQualifiedErrorId","FullName","Source")] 
        [String]$Title, 
 
    [Parameter(Position=2)] 
        [ValidateSet("OK","OKCancel","YesNo","YesNoCancel")] 
        [String]$Button="OK", 
 
    [Parameter(Position=3)] 
        [ValidateSet("Asterisk","Error","Exclamation","Hand","Information","None","Question","Stop","Warning")] 
        [String]$Icon="Information", 
 
    [Parameter(Position=4)] 
        [ValidateSet("Yes","No","OK","Cancel","None")] 
        [String]$DefaultButton="None", 
 
    [Parameter(ParameterSetName="Quiet")][Switch]$Quiet, 
 
    [Parameter(ParameterSetName="Bool")][Switch]$OutputBool 
) 
 
Begin { 
    try {[System.Windows.Window]} 
    catch{Add-Type -AssemblyName PresentationCore,PresentationFramework} 
 
    $Button = [System.Windows.MessageBoxButton]::$Button 
    $Icon = [System.Windows.MessageBoxImage]::$Icon 
    $DefaultButton = [System.Windows.MessageBoxResult]::$DefaultButton 
 
    if ($OutputBool) { 
        if ($Button -match "OK$|oCancel$") { 
            Write-Error "The OutputBool parameter is only valid when there are only two buttons: OKCancel and YesNo. Boolean output will be unavailble during this execution." 
        } 
        else {$boolValid = $true} 
    }  
} 
 
Process { 
    $result = [System.Windows.MessageBox]::Show($Message,$Title,$Button,$Icon,$DefaultButton) 
 
    switch ($true) { 
        $Quiet {$result = $null} 
        {$boolValid -and ($result -match "^O|^Y")} {$result = $true; break} 
        $boolValid {$result = $false} 
    } 
     
    $result  
}