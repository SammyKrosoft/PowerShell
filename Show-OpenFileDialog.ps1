function Show-OpenFileDialog {
<#
.DESCRIPTION
    This function is a PowerShell call to the OpenFileDialog box using WPF.
    We'll see many examples of OpenFileDialog using Windows Forms, but we'll try to
    get off Windows Forms as it's legacy tech.

.EXAMPLE
    PS_>Show-OpenFileDialog
    This will open a dialog box to enable the user to select a file. The output of
    the function is the full path of that file. It's useful for example to Import-CSV
    from a CSV file, or to select an Office document to be opened with PowerShell
    application automation...

.EXAMPLE
    PS_>$FileName = Show-OpenFileDialog
    This will open a dialog box to enable the user to select a file, and the file name
    will be stored in the $FileName variable to be reused as described on the first example.

.EXAMPLE
    PS_>$FileName = Show-OpenFileDialog -Title "Open an .XLSX file to be parsed" -Filter "Excel file|*.xlsx"  -InitialDirectory c:\MyExcelFiles
    This will open a dialog box to select a file, with a customized title, and with a default filter on *.xlsx Excel files. This
    dialog box will open the C:\MyExcelFiles directory to look for files. User can select later any other folder to look for.

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