function Show-OpenFileDialog {
<#
.DESCRIPTION
    This function is a PowerShell call to the OpenFileDialog box using WPF.
    We'll see many examples of OpenFileDialog using Windows Forms, but we'll try to
    get off Windows Forms as it's legacy tech.

.LINK
    https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32?view=net-5.0

#>

    param
    ($Title = 'Select a file to use', $Filter = 'Comma Separated|*.csv|Text|*.txt',$InitialDirectory = "c:\temp")
    
    Add-Type -AssemblyName PresentationFramework

    $dialog = New-Object -TypeName 'Microsoft.Win32.OpenFileDialog'
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    If (!(Test-Path $InitialDirectory)){$InitialDirectory = "$($env:Userprofile)\Documents"} #If the default C:\temp doesn't exist, defaults to user's Document folder
    $dialog.InitialDirectory = $InitialDirectory
  
    if ($dialog.ShowDialog() -eq $true)
    {
        Return $dialog.FileName
    }
    else
    {
        Write-Warning 'Cancelled'
    }
}

Show-OpenFileDialog