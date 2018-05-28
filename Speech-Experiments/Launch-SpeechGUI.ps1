<# Functions #>

Function Working-Label ($msg,$BG = 1) {
    # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    Switch ($BG) {
        0   {$wpf.lblStatus.Background = "green"}
        1   {$wpf.lblStatus.Background = "red"}
    }
    $wpf.lblStatus.Text = $msg
    # $wpf.WForm.Dispatcher.Invoke("Render",[action][scriptblock]::create({}))
    $wpf.frmSpeechGUI.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function Init-Speech {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    $InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
     write-host $InstalledVoices
}

Function Say {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
     #$InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
     #write-host $InstalledVoices
     # Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior',[int32]'Position which voices are ordered','fr/en')
    switch ($wpf.lstBoxLanguage.SelectedItem.Content) {
        "Francais" {$Language = 'fr'}
        "English" {$Language = 'en'}
        "" {$language = 'en'}
        $null {$Language = 'en'}
    }
    $Speak.rate = $wpf.txtSpeed.Text
    $Speak.SelectVoiceByHints(0,0,0,$language)
    $Speak.Speak($Msg)
    
}
<#\Functions#>

#region Form definition

Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"

<Window x:Name="frmSpeechGUI" x:Class="SpeechGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:SpeechGUI"
        mc:Ignorable="d"
        Title="SpeechGUI" Height="450" Width="800">
    <Grid Background="#FF8081FF">
        <ListBox x:Name="lstBoxLanguage" HorizontalAlignment="Left" Height="52" Margin="10,10,0,0" VerticalAlignment="Top" Width="94" SelectedIndex="1">
            <ListBoxItem Content="Francais"/>
            <ListBoxItem Content="English"/>
        </ListBox>
        <TextBox x:Name="txtInputBox" HorizontalAlignment="Left" Height="184" Margin="10,114,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="775" FontSize="20"/>
        <Label Content="Text to speech:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,83,0,0"/>
        <Button x:Name="btnRun" Content="Run !" HorizontalAlignment="Left" VerticalAlignment="Top" Width="244" Margin="133,322,0,0" Height="66" FontSize="36"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" VerticalAlignment="Top" Width="244" Margin="454,322,0,0" Height="66" FontSize="20"/>
        <Label HorizontalAlignment="Left" Margin="240,32,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtSpeed" HorizontalAlignment="Left" Height="42" Margin="172,47,0,0" TextWrapping="Wrap" Text="5" VerticalAlignment="Top" Width="36" FontSize="20" IsReadOnly="True" TextAlignment="Center" VerticalContentAlignment="Center" HorizontalContentAlignment="Center"/>
        <Button x:Name="btnSlower" Content="&lt;" HorizontalAlignment="Center" VerticalAlignment="Top" Width="34" Margin="133,47,625,0" Height="42" FontSize="24" FontWeight="Bold" Cursor="Hand" UseLayoutRounding="False"/>
        <Button x:Name="btnFaster" Content="&gt;" HorizontalAlignment="Center" VerticalAlignment="Top" Width="34" Margin="213,47,545,0" Height="42" FontSize="24" FontWeight="Bold" Cursor="Hand"/>
        <Label Content="Speed / Vitesse:" FontSize="20" Margin="125,5,503,377"/>
        <TextBlock x:Name="lblStatus" HorizontalAlignment="Left" TextWrapping="Wrap" Text="Ready. Make me speak !" VerticalAlignment="Top" Margin="328,26,0,0" Height="63" Width="327" FontSize="24" FontWeight="Bold" TextAlignment="Center" Background="Lime" TextOptions.TextHintingMode="Fixed"/>

    </Grid>
</Window>

"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#endregion

#region FORM EVENTS HANDLING

#region Buttons
$wpf.btnRun.add_Click({
    Working-Label "Busy. I'm speaking, wait..."
    Say $wpf.txtInputBox.text
    Working-Label "Ready. Make me speak !" 0
})

$wpf.btnCancel.add_Click({
    $msg = "Exiting..."
    Working-Label "Busy. I'm speaking, wait..."
    Say $msg
    Working-Label "Ready. Make me speak !" 0
    $wpf.frmSpeechGUI.Close()
})

$wpf.btnSlower.add_click({
    if ($wpf.txtSpeed.text.ToInt32($Null) -le -10) {
        $wpf.txtSpeed.text = -10
    } Else {
        $wpf.txtSpeed.Text = $wpf.txtSpeed.Text.ToInt32($Null) - 1
    }
})

$wpf.btnFaster.add_click({
    If ($wpf.txtSpeed.Text.ToInt32($Null) -ge 10) {
        $wpf.txtSpeed.Text = 10
    } Else {
        $wpf.txtSpeed.text = $wpf.txtSpeed.Text.ToInt32($Null) + 1
    }
})

# End of Buttons region
#endregion

#region Language Selection box
$wpf.lstBoxLanguage.add_SelectionChanged({
    $msg = "Language = $($wpf.lstBoxLanguage.SelectedItem.content)"
    If ($($wpf.lstBoxLanguage.SelectedItem.content) -eq "Francais") {
        $msg = "Langue Francaise sélectionnée !"
    } Else {
        $msg = "English Language selected !"
    }
    Working-Label "Busy. I'm speaking, wait..."
    Say $msg
    Working-Label "Ready. Make me speak !" 0
})

# End of Language Selection box
#endregion
#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.frmSpeechGUI.Add_Loaded({
})

#Things to load when the WPF form is rendered aka drawn on screen
$wpf.frmSpeechGUI.Add_ContentRendered({
Working-Label "Busy. I'm speaking, wait..."
Say "Bonjour Liam!"
Working-Label "Ready. Make me speak !" 0
})
$wpf.frmSpeechGUI.add_Closing({
    Working-Label "Busy. I'm speaking, wait..."
    $msg = "Sssichering !"
    Say $msg
    Working-Label "Ready. Make me speak !" 0
})
# End of load, draw and closing form events
#endregion

#END OF EVENTS HANDLING
#endregion

Init-Speech

#$wpf.frmSpeechGUI.ShowDialog() | Out-null
