
#region Code to define WPF form and controls

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"

<Window x:Name="MyForm" x:Class="simple_datagrid_tests.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:simple_datagrid_tests"
        mc:Ignorable="d"
        Title="MainWindow" Height="450" Width="800">
    <Grid>
        <DataGrid x:Name="dataGrid" HorizontalAlignment="Left" Height="402" VerticalAlignment="Top" Width="775" Margin="10,10,0,0"/>

    </Grid>
</Window>

"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#endregion (code to define WPF form and controls)

#region Execution Code

$Array = @{}
$Array = @{"IP Address"="0.0.0.0-255.255.255.255";"Type" = "Single"}

$wpf.dataGrid.ItemsSource = $Array


#endregion Execution Code




#region Load the form:
# Older way >>>>> $wpf.MyForm.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.MyForm.Dispatcher.InvokeAsync({
    $wpf.MyForm.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null
#endregion (Load the form)