
# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"

Here-String pasted from XAML

"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#Get the form name to be used as parameter in functions external to form...
$FormName = $NamedNodes[0].Name


#Define events functions
#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.$FormName.Add_Loaded({
    #Update-Cmd
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.$FormName.Add_ContentRendered({
    #Update-Cmd
})
$wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})

#endregion Load, Draw and closing form events
#End of load, draw and closing form events

#HINT: to update progress bar and/or label during WPF Form treatment, add the following:
# ... to re-draw the form and then show updated controls in realtime ...
$wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null