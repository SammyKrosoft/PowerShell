# Preparing for small messages boxes
# Always load WPF assembly to be able to use "[System.Windows.MessageBox]"
Add-Type -AssemblyName presentationframework, presentationcore

Function MsgBox ($msg = "Do you want to continue ?",$Title = "Information...",$Button = "Ok",$Icon = "Information"){
    [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
}


# Word document to open
$DocFile = "C:\temp\E2016Build.docx"

# Bookmarks list
$BookMark = "DEFAULT_GATEWAY"

#Word Object definition
$WordObject = New-Object -ComObject Word.Application
$WordObject.Visible = $true

$OpenedDocFile = $WordObject.Documents.Open($DocFile)

If ($OpenedDocFile.Bookmarks.Exists($BookMark)){
    MsgBox -msg "BookMark Exists !"
} Else {
    MsgBox -msg "Bookmark does NOT exist !" -Icon "Error"
}