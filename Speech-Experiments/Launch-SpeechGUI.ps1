<# Functions #>
Function Say {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    # $InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
    # $InstalledVoices
    # Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior',[int32]'Position which voices are ordered','fr/en')
    switch ($wpf.lstBoxLanguage.SelectedItem.Content) {
        "Francais" {$Language = 'fr'}
        "English" {$Language = 'en'}
        "" {$language = 'en'}
        $null {$Language = 'en'}
    }
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
    Say $wpf.txtInputBox.text
})

$wpf.btnCancel.add_Click({
    $msg = "Exiting..."
    Say $msg
    $wpf.frmSpeechGUI.Close()
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
    Say $msg})

# End of Language Selection box
#endregion

#END OF EVENTS HANDLING
#endregion

$wpf.frmSpeechGUI.ShowDialog() | Out-null
