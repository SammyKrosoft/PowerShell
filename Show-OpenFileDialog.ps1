function Show-OpenFileDialog
{
    param
    ($Title = 'Select a file to use', $Filter = 'Comma Separated|*.csv|Text|*.txt',$InitialDirectory = "c:\temp")

    $dialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog
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