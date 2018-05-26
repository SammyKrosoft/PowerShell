
#region FUNCTIONS other than Form events

Function Run-Action{
    $SelectedAction = $wpf.comboSelectAction.SelectedItem.Content
    Switch ($SelectedAction) {
        "Display Info"  {
            Write-host "Displaying Info"
            Write-Host "Listing selected mailbox names:"
            $SelectedITems = $wpf.GridView.SelectedItems
            $SelectedItems | Foreach{
                Write-Host $_.Name
            }
        }
        "Kill process"  {
            Write-Host "Kill process not implemented yet..."
        }
    }
    Update-Label "Action done."
}

Function Update-Label ($msg) {
    $wpf.lblStatus.Content = $msg
}

Function Working-Label {
        # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.lblStatus.Content = "Working ..."
    # $wpf.WForm.Dispatcher.Invoke("Render",[action][scriptblock]::create({}))
    $wpf.WForm.Dispatcher.Invoke("Render",[action][scriptblock]{})
}
Function Get-Mailboxes {
    $SearchSubstring = ("*") + ($wpf.txtMailboxString.text) + ("*")
    Try {
        #$Mailboxes = Get-Mailbox -ResultSize Unlimited $SearchSubstring -ErrorAction Stop | Select Name,DisplayName,primarySMTPAddress
        $Processes = Get-process -Name $SearchSubstring -ErrorAction Stop 
        #[System.Collections.IENumerable]$Results = @($Mailboxes)
        [System.Collections.IENumerable]$Results = @($Processes)
        $wpf.GridView.ItemsSource = $Results
        $wpf.GridView.Columns | Foreach {
            $_.CanUserSort = $true
        }
        $wpf.lblStatus.Content = "Found $($Results.Count) Mailbox(es)"
        $wpf.lblNbItemsInGrid.Content = $($Results.Count)
    }

    Catch {
        $Mailboxes = $null
        $wpf.lblStatus.Content = "No mailboxes found... try again !"
    }
}

#endregion

#========================================================
#region WPF form definition and load controls
#========================================================

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
# $inputXML = Get-Content -Path "C:\Users\Kamehameha\Documents\GitHub\PowerShell\Get-EventsFromEventLog\VisualStudio2017WPFDesign\Launch-EventsCollector-WPF\Launch-EventsCollector-WPF\MainWindow.xaml"
$inputXML = @"

<Window x:Name="WForm" x:Class="GridView_WPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:GridView_WPF"
        mc:Ignorable="d"
        Title="Search Mailboxes" Height="450" Width="800" ResizeMode="NoResize">
    <Grid>
        <DataGrid x:Name="GridView" HorizontalAlignment="Left" Height="349" Margin="353,31,0,0" VerticalAlignment="Top" Width="410" Grid.ColumnSpan="2"/>
        <TextBox x:Name="txtMailboxString" HorizontalAlignment="Left" Height="23" Margin="10,87,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302"/>
        <Label Content="Search for mailbox (substring of alias, e-mail address, &#xD;&#xA;display name, ...)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,31,0,0" Height="51" Width="302"/>
        <Button x:Name="btnRun" Content="Search" HorizontalAlignment="Left" Margin="10,115,0,0" VerticalAlignment="Top" Width="75"/>
        <Label x:Name="lblStatus" Content="Please start a search..." HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,189,0,0" Width="255" FontStyle="Italic"/>
        <Button x:Name="btnAction" Content="Action" Margin="273,302,446,89.5" IsEnabled="False"/>
        <ComboBox x:Name="comboSelectAction" HorizontalAlignment="Left" Margin="228,337,0,0" VerticalAlignment="Top" Width="120" Height="24" SelectedIndex="0" IsEnabled="False">
            <ComboBoxItem Content="Display Info"/>
            <ComboBoxItem Content="Kill process"/>
        </ComboBox>
        <Label x:Name="lblNbItemsInGrid" Content="0" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="506,385,0,0" Width="55"/>
        <Label Content="Number of Items in Grid:" HorizontalAlignment="Left" Margin="353,385,0,0" VerticalAlignment="Top" Width="148"/>
        <Label Content="Selected:" HorizontalAlignment="Left" Margin="621,385,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblNumberItemsSelected" Content="0" HorizontalAlignment="Left" Margin="684,385,0,0" VerticalAlignment="Top"/>

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

#region Buttons
$wpf.btnRun.add_Click({
    Working-Label
    Get-Mailboxes
})

$wpf.btnAction.add_Click({
    Working-Label
    Run-Action
})
# End of Buttons region
#endregion

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.WForm.Add_Loaded({

})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.WForm.Add_ContentRendered({

})
$wpf.WForm.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})
# End of load, draw and closing form events
#endregion

#region Text Changed events

$wpf.GridView.add_SelectionChanged({
    $Selected = $wpf.GridView.SelectedItems.count
    If ($Selected -eq 0) {
        $wpf.btnAction.IsEnabled = $false
        $wpf.comboSelectAction.IsEnabled = $false
    } ElseIf ($Selected -gt 0) {
        $wpf.btnAction.IsEnabled = $true
        $wpf.comboSelectAction.IsEnabled = $true
    }
    $wpf.lblNumberItemsSelected.Content = $Selected
})
#End of Text Changed events
#endregion


#endregion

#=======================================================
#End of Events from the WPF form
#endregion
#=======================================================


# Load the form:
$wpf.WForm.ShowDialog() | Out-Null