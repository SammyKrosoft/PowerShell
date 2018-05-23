#========================================================
#region Functions definitions (NOT the WPF form events)
#========================================================

function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

Function Split-ListColon {
    param(
        [string]$StringToSplit,
        [switch]$Noquotes
    )
    $TargetSplit = $StringToSplit.Split(',')
    $ListItems = ""
    If ($NoQuotes){
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += $TargetSplit[$i].trim() + (", ")}
        $ListItems += $TargetSplit[$TargetSplit.Count - 1].trim()
    } Else {
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += ("""") + $TargetSplit[$i].trim() + (""", ")}
        $ListItems += ("""") + $TargetSplit[$TargetSplit.Count - 1].trim() + ("""")
    }
    Return $ListItems
}


#========================================================
#endregion Functions definitions (NOT the WPF form events)
#========================================================

Function Get-Mailboxes {
    $wpf.listBoxMailboxes.items.Clear()
    $Mailboxes = Get-Mailbox "*$($wpf.txtMailboxString.text)*"| Select Name,DisplayName,alias,primarySMTPADdress, emailaddresses
    $Mailboxes | % {
        $wpf.listBoxMailboxes.Items.Add($_.DisplayName)
    }
}

#========================================================
#region WPF form definition and load controls
#========================================================

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
# $inputXML = Get-Content -Path "C:\Users\Kamehameha\Documents\GitHub\PowerShell\Get-EventsFromEventLog\VisualStudio2017WPFDesign\Launch-EventsCollector-WPF\Launch-EventsCollector-WPF\MainWindow.xaml"
$inputXML = @"

<Window x:Name="GetMailboxForm" x:Class="WpfApp4.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp4"
        mc:Ignorable="d"
        Title="Get and Set Mailbox Basic Properties" Height="450" Width="800" ResizeMode="NoResize">
    <Grid>
        <TextBox x:Name="txtMailboxString" HorizontalAlignment="Left" Height="33" Margin="10,54,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="272"/>
        <Label Content="Mailbox(es) to look for" HorizontalAlignment="Left" Margin="10,23,0,0" VerticalAlignment="Top" Width="134"/>
        <ListBox x:Name="listBoxMailboxes" HorizontalAlignment="Left" Height="301" Margin="10,111,0,0" VerticalAlignment="Top" Width="775"/>
        <Button x:Name="btnSearch" Content="Search" Margin="311,54,420,335.5"/>
        <Label Content="Label" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Content="ResultSize" HorizontalAlignment="Left" Margin="500,31,0,0" VerticalAlignment="Top" Width="239"/>
        <TextBox x:Name="txtResultSize" HorizontalAlignment="Left" Height="23" Margin="500,62,0,0" TextWrapping="Wrap" Text="100" VerticalAlignment="Top" Width="184"/>

    </Grid>
</Window>

"@
$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#========================================================
# END of WPF form definition and load controls
#endregion
#========================================================

#========================================================
#region WPF EVENTS definition
#========================================================

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.GetMailboxForm.Add_Loaded({
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.GetMailboxForm.Add_ContentRendered({
    Write-Host "Windows WPF Form Loaded."
})
$wpf.GetMailboxForm.add_Closing({
    Write-Host "Bye !"
})
# End of load, draw and closing form events
#endregion

#region Clicked on Checkboxes events
$wpf.btnSearch.add_Click({
    Write-Host "Clicked on the [Search] button..."
    Get-mailboxes
})

# End of Clicked on Checkboxes events
#endregion


#========================================================
#endregion WPF EVENTS definition
#========================================================


# Load the form:
$wpf.GetMailboxForm.ShowDialog() | Out-Null
